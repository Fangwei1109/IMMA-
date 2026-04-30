const path = require("path");
const os = require("os");
const fs = require("fs/promises");
const { spawn } = require("child_process");
const { buildTerminologyHints } = require("../obsidian");

const ROOT = path.join(__dirname, "..", "..");
const LOCAL_WHISPER_SCRIPT = path.join(ROOT, "scripts", "transcribe_local_whisper.py");
const LOCAL_WHISPER_PYTHON =
  process.env.LOCAL_WHISPER_PYTHON ||
  "C:\\Users\\HMATC\\AppData\\Local\\Programs\\Python\\Python311\\python.exe";
const ALI_SHORT_AUDIO_LIMIT_BYTES = 10 * 1024 * 1024;

function buildProviders() {
  return {
    auto: {
      id: "auto",
      name: "Auto Select",
      description: "Choose SenseVoice/FunASR when configured, otherwise prefer Alibaba ASR for Chinese-heavy calls, then local Whisper.",
      configured: true,
      mode: "virtual",
    },
    chinese: {
      id: "chinese",
      name: "SenseVoice / FunASR",
      description: "Chinese-first open-source ASR via a SenseVoice-compatible HTTP service.",
      configured: Boolean(process.env.SENSEVOICE_API_URL || process.env.CHINESE_ASR_API_URL),
      mode: "remote",
    },
    whisper: {
      id: "whisper",
      name: "Whisper API",
      description: "OpenAI-compatible audio transcription endpoint for multilingual audio.",
      configured: Boolean(
        process.env.WHISPER_API_KEY || process.env.OPENAI_API_KEY,
      ),
      mode: "remote",
    },
    ali: {
      id: "ali",
      name: "Alibaba / Qwen ASR",
      description: "Qwen3-ASR-Flash via DashScope. Long audio is automatically compressed and chunked with ffmpeg.",
      configured: Boolean(process.env.ALI_API_KEY),
      mode: "remote",
    },
    "local-whisper": {
      id: "local-whisper",
      name: "Local Whisper / faster-whisper",
      description: "Run faster-whisper locally with CPU int8 inference and Obsidian terminology hints.",
      configured: true,
      mode: "local",
    },
  };
}

function getAsrProviders() {
  return Object.values(buildProviders());
}

function resolveAsrSelection(requested) {
  const providers = buildProviders();
  const wanted = requested || process.env.ASR_PROVIDER_DEFAULT || "auto";

  if (wanted === "auto") {
    if (providers.chinese.configured) {
      return "chinese";
    }

    if (providers.ali.configured) {
      return "ali";
    }

    if (providers["local-whisper"].configured) {
      return "local-whisper";
    }

    if (providers.whisper.configured) {
      return "whisper";
    }

    return "ali";
  }

  if (!providers[wanted]) {
    throw createError(400, `Unknown ASR provider "${wanted}".`);
  }

  return wanted;
}

async function transcribeAudio({ providerId, audioFile, context = {} }) {
  const chosen = resolveAsrSelection(providerId);
  const providers = buildProviders();

  if (!audioFile) {
    throw createError(400, "Audio file is required for ASR transcription.");
  }

  if (!providers[chosen].configured) {
    throw createError(
      400,
      `ASR provider "${chosen}" is not configured. Set the matching environment variables first.`,
    );
  }

  if (chosen === "chinese") {
    return transcribeWithChineseAsr(audioFile);
  }

  if (chosen === "whisper") {
    return transcribeWithWhisper(audioFile);
  }

  if (chosen === "ali") {
    return transcribeWithAliAsr(audioFile);
  }

  if (chosen === "local-whisper") {
    return transcribeWithLocalWhisper(audioFile, context);
  }

  throw createError(500, `ASR provider "${chosen}" is not implemented.`);
}

async function transcribeWithChineseAsr(audioFile) {
  const endpoint =
    process.env.SENSEVOICE_API_URL || process.env.CHINESE_ASR_API_URL;
  const apiKey =
    process.env.SENSEVOICE_API_KEY || process.env.CHINESE_ASR_API_KEY || "";
  const language =
    process.env.SENSEVOICE_LANGUAGE ||
    process.env.CHINESE_ASR_LANGUAGE ||
    "zh";
  const useItn =
    (process.env.SENSEVOICE_USE_ITN || "true").toLowerCase() !== "false";
  const fileName = audioFile.originalname || "meeting-audio.wav";
  const mimeType = audioFile.mimetype || guessMimeType(fileName);

  const form = new FormData();
  form.append("file", new Blob([audioFile.buffer], { type: mimeType }), fileName);
  form.append("language", language);
  form.append("use_itn", String(useItn));

  const response = await fetch(endpoint, {
    method: "POST",
    headers: apiKey
      ? {
          Authorization: `Bearer ${apiKey}`,
        }
      : undefined,
    body: form,
  });

  if (!response.ok) {
    const detail = await response.text();
    throw createError(502, `Chinese ASR request failed: ${detail}`);
  }

  const payload = await response.json();
  const text =
    payload.text ||
    payload.transcript ||
    payload.result?.[0]?.text ||
    payload.result?.text ||
    payload.data?.result?.[0]?.text ||
    payload.data?.text ||
    "";

  if (!text) {
    throw createError(
      502,
      "Chinese ASR response did not contain a transcript text field.",
    );
  }

  return {
    providerId: "chinese",
    text,
    raw: payload,
  };
}

async function transcribeWithWhisper(audioFile) {
  const endpoint =
    process.env.WHISPER_API_URL ||
    process.env.OPENAI_AUDIO_TRANSCRIPTIONS_URL ||
    "https://api.openai.com/v1/audio/transcriptions";
  const apiKey = process.env.WHISPER_API_KEY || process.env.OPENAI_API_KEY;
  const model = process.env.WHISPER_MODEL || "whisper-1";
  const language = process.env.WHISPER_LANGUAGE || "";
  const prompt = process.env.WHISPER_PROMPT || "";
  const fileName = audioFile.originalname || "meeting-audio.wav";
  const mimeType = audioFile.mimetype || guessMimeType(fileName);

  const form = new FormData();
  form.append("file", new Blob([audioFile.buffer], { type: mimeType }), fileName);
  form.append("model", model);
  if (language) {
    form.append("language", language);
  }
  if (prompt) {
    form.append("prompt", prompt);
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    body: form,
  });

  if (!response.ok) {
    const detail = await response.text();
    throw createError(502, `Whisper ASR request failed: ${detail}`);
  }

  const payload = await response.json();
  const text = payload.text || "";

  if (!text) {
    throw createError(502, "Whisper ASR response did not contain transcript text.");
  }

  return {
    providerId: "whisper",
    text,
    raw: payload,
  };
}

async function transcribeWithAliAsr(audioFile) {
  const config = getAliAsrConfig();

  if (!config.apiKey) {
    throw createError(400, 'ASR provider "ali" is not configured. Set ALI_API_KEY first.');
  }

  if ((audioFile.buffer?.length || 0) > ALI_SHORT_AUDIO_LIMIT_BYTES) {
    return transcribeWithAliAsrChunks(audioFile, config);
  }

  return transcribeAliAsrSingle(audioFile, config);
}

function getAliAsrConfig() {
  return {
    endpoint: `${(process.env.ALI_ASR_BASE_URL || process.env.ALI_BASE_URL || "https://dashscope.aliyuncs.com/compatible-mode/v1").replace(/\/$/, "")}/chat/completions`,
    apiKey: process.env.ALI_API_KEY,
    model: process.env.ALI_ASR_MODEL || "qwen3-asr-flash",
    language: process.env.ALI_ASR_LANGUAGE || "",
    enableItn: (process.env.ALI_ASR_ENABLE_ITN || "true").toLowerCase() !== "false",
  };
}

async function transcribeAliAsrSingle(audioFile, config) {
  const fileName = audioFile.originalname || "meeting-audio.wav";
  const mimeType = audioFile.mimetype || guessMimeType(fileName);

  if ((audioFile.buffer?.length || 0) > ALI_SHORT_AUDIO_LIMIT_BYTES) {
    throw createError(
      400,
      `Alibaba ASR chunk is still over 10 MB after compression (${Math.round(audioFile.buffer.length / 1024 / 1024)} MB). Lower ALI_ASR_CHUNK_SECONDS or ALI_ASR_CHUNK_BITRATE.`,
    );
  }

  const dataUri = `data:${mimeType};base64,${audioFile.buffer.toString("base64")}`;
  const body = {
    model: config.model,
    messages: [
      {
        role: "user",
        content: [
          {
            type: "input_audio",
            input_audio: {
              data: dataUri,
            },
          },
        ],
      },
    ],
    stream: false,
    asr_options: {
      enable_itn: config.enableItn,
      ...(config.language ? { language: config.language } : {}),
    },
  };

  const response = await fetch(config.endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw createError(502, `Alibaba ASR request failed: ${detail}`);
  }

  const payload = await response.json();
  const text = payload?.choices?.[0]?.message?.content || "";

  if (!text) {
    throw createError(502, "Alibaba ASR response did not contain transcript text.");
  }

  return {
    providerId: "ali",
    text,
    raw: payload,
  };
}

async function transcribeWithAliAsrChunks(audioFile, config) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "ali-asr-"));
  const originalExt = path.extname(audioFile.originalname || "") || ".audio";
  const inputPath = path.join(tempDir, `input${originalExt}`);
  const chunkPattern = path.join(tempDir, "chunk-%04d.mp3");
  const chunkSeconds = Number(process.env.ALI_ASR_CHUNK_SECONDS || 300);
  const chunkBitrate = process.env.ALI_ASR_CHUNK_BITRATE || "32k";

  await fs.writeFile(inputPath, audioFile.buffer);

  try {
    await runProcess("ffmpeg", [
      "-hide_banner",
      "-y",
      "-i",
      inputPath,
      "-vn",
      "-ac",
      "1",
      "-ar",
      "16000",
      "-b:a",
      chunkBitrate,
      "-f",
      "segment",
      "-segment_time",
      String(chunkSeconds),
      "-reset_timestamps",
      "1",
      chunkPattern,
    ]);

    const chunkFiles = (await fs.readdir(tempDir))
      .filter((file) => /^chunk-\d+\.mp3$/.test(file))
      .sort();

    if (!chunkFiles.length) {
      throw createError(502, "ffmpeg did not produce audio chunks for Alibaba ASR.");
    }

    const chunks = [];
    for (const [index, file] of chunkFiles.entries()) {
      const chunkPath = path.join(tempDir, file);
      const buffer = await fs.readFile(chunkPath);
      const result = await transcribeAliAsrSingle(
        {
          originalname: file,
          mimetype: "audio/mpeg",
          buffer,
        },
        config,
      );
      chunks.push({
        index,
        file,
        size: buffer.length,
        text: result.text,
      });
    }

    return {
      providerId: "ali",
      text: chunks.map((chunk) => chunk.text).filter(Boolean).join("\n"),
      raw: {
        chunked: true,
        originalSize: audioFile.buffer.length,
        chunkSeconds,
        chunkBitrate,
        chunks,
      },
    };
  } catch (error) {
    if (/ENOENT|not recognized|not found/i.test(error.message || "")) {
      throw createError(
        400,
        "Long-audio Alibaba ASR requires ffmpeg. Install ffmpeg or choose Local Whisper for this audio.",
      );
    }
    throw error;
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function transcribeWithLocalWhisper(audioFile, context = {}) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "local-whisper-"));
  const fileName = audioFile.originalname || "meeting-audio.wav";
  const audioPath = path.join(tempDir, fileName);
  const optionsPath = path.join(tempDir, "options.json");
  const terminology = await buildTerminologyHints({
    company: context.company || process.env.LOCAL_WHISPER_HINT_COMPANY || "",
    title: context.title || process.env.LOCAL_WHISPER_HINT_TITLE || "",
    userInstructions: context.userInstructions || process.env.LOCAL_WHISPER_HINTS || "",
    limit: 18,
  });

  await fs.writeFile(audioPath, audioFile.buffer);
  await fs.writeFile(
    optionsPath,
        JSON.stringify(
      {
        model: process.env.LOCAL_WHISPER_MODEL || "medium",
        device: process.env.LOCAL_WHISPER_DEVICE || "cpu",
        compute_type: process.env.LOCAL_WHISPER_COMPUTE_TYPE || "int8",
        language: process.env.LOCAL_WHISPER_LANGUAGE || "",
        beam_size: Number(process.env.LOCAL_WHISPER_BEAM_SIZE || 8),
        best_of: Number(process.env.LOCAL_WHISPER_BEST_OF || 5),
        temperature: Number(process.env.LOCAL_WHISPER_TEMPERATURE || 0),
        condition_on_previous_text:
          (process.env.LOCAL_WHISPER_CONDITION_ON_PREVIOUS_TEXT || "false").toLowerCase() !== "false",
        vad_filter: (process.env.LOCAL_WHISPER_VAD_FILTER || "true").toLowerCase() !== "false",
        initial_prompt: terminology.length
          ? `Domain terms and names: ${terminology.join(", ")}.`
          : "",
      },
      null,
      2,
    ),
    "utf8",
  );

  try {
    const payload = await runLocalWhisper(audioPath, optionsPath);
    const text = payload?.text || "";
    if (!text) {
      throw createError(502, "Local Whisper returned no transcript text.");
    }

    return {
      providerId: "local-whisper",
      text,
      raw: payload,
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function runLocalWhisper(audioPath, optionsPath) {
  const stdout = await new Promise((resolve, reject) => {
    const child = spawn(LOCAL_WHISPER_PYTHON, [LOCAL_WHISPER_SCRIPT, audioPath, optionsPath], {
      stdio: ["ignore", "pipe", "pipe"],
      env: {
        ...process.env,
        HF_ENDPOINT: process.env.HF_ENDPOINT || "https://hf-mirror.com",
        HF_HUB_DISABLE_SYMLINKS_WARNING: "1",
        PYTHONIOENCODING: "utf-8",
      },
    });

    let out = "";
    let err = "";
    child.stdout.on("data", (chunk) => {
      out += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      err += chunk.toString();
    });
    child.on("error", (error) => {
      reject(createError(500, `Failed to start local Whisper: ${error.message}`));
    });
    child.on("exit", (code) => {
      if (code === 0) {
        resolve(out);
        return;
      }
      reject(createError(502, `Local Whisper failed: ${err || `exit code ${code}`}`));
    });
  });

  try {
    return JSON.parse(stdout);
  } catch (error) {
    throw createError(502, `Local Whisper returned invalid JSON: ${error.message}`);
  }
}

async function runProcess(command, args) {
  await new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: ["ignore", "pipe", "pipe"],
    });

    let err = "";
    child.stderr.on("data", (chunk) => {
      err += chunk.toString();
    });
    child.on("error", (error) => {
      reject(error);
    });
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(createError(502, `${command} failed: ${err || `exit code ${code}`}`));
    });
  });
}

function guessMimeType(fileName) {
  const ext = path.extname(fileName).toLowerCase();
  const mapping = {
    ".mp3": "audio/mpeg",
    ".m4a": "audio/mp4",
    ".wav": "audio/wav",
    ".mp4": "audio/mp4",
    ".aac": "audio/aac",
    ".ogg": "audio/ogg",
    ".webm": "audio/webm",
  };

  return mapping[ext] || "application/octet-stream";
}

function createError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

module.exports = {
  getAsrProviders,
  resolveAsrSelection,
  transcribeAudio,
};
