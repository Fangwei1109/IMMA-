const test = require("node:test");
const assert = require("node:assert/strict");
const { transcribeAudio, resolveAsrSelection } = require("./asr");

test("resolveAsrSelection prefers Alibaba ASR when chinese provider is not configured", () => {
  const originalSenseVoiceUrl = process.env.SENSEVOICE_API_URL;
  const originalChineseUrl = process.env.CHINESE_ASR_API_URL;
  const originalAliKey = process.env.ALI_API_KEY;

  delete process.env.SENSEVOICE_API_URL;
  delete process.env.CHINESE_ASR_API_URL;
  process.env.ALI_API_KEY = "test-key";

  assert.equal(resolveAsrSelection("auto"), "ali");

  restoreEnv("SENSEVOICE_API_URL", originalSenseVoiceUrl);
  restoreEnv("CHINESE_ASR_API_URL", originalChineseUrl);
  restoreEnv("ALI_API_KEY", originalAliKey);
});

test("transcribeAudio parses SenseVoice-style nested response", async () => {
  const originalSenseVoiceUrl = process.env.SENSEVOICE_API_URL;
  const originalFetch = global.fetch;

  process.env.SENSEVOICE_API_URL = "http://127.0.0.1:50000/asr";

  global.fetch = async () => ({
    ok: true,
    async json() {
      return {
        result: [{ text: "这是一个中文会议转录结果。" }],
      };
    },
  });

  const result = await transcribeAudio({
    providerId: "chinese",
    audioFile: {
      originalname: "meeting.wav",
      mimetype: "audio/wav",
      buffer: Buffer.from("fake"),
    },
  });

  assert.equal(result.providerId, "chinese");
  assert.equal(result.text, "这是一个中文会议转录结果。");

  global.fetch = originalFetch;
  restoreEnv("SENSEVOICE_API_URL", originalSenseVoiceUrl);
});

function restoreEnv(key, value) {
  if (typeof value === "undefined") {
    delete process.env[key];
    return;
  }

  process.env[key] = value;
}
