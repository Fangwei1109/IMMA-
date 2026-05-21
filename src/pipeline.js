const path = require("path");
const fs = require("fs/promises");
const os = require("os");
const crypto = require("crypto");
const { PDFParse } = require("pdf-parse");
const { spawn } = require("child_process");
const iconv = require("iconv-lite");
const mammoth = require("mammoth");
const JSZip = require("jszip");
const XLSX = require("xlsx");
const { buildTerminologyHints, getObsidianConfig } = require("./obsidian");
const { DATA_DIR, OUTPUTS_DIR, ensureFolders, getStorageAdapter } = require("./storage");
const { getAsrProviders, resolveAsrSelection, transcribeAudio } = require("./providers/asr");
const {
  enrichCompanyContext,
  getResearchProviders,
  resolveResearchProvider,
} = require("./providers/research");
const {
  extractImageMaterialText,
  getLlmProviders,
  getResolvedLlmConfig,
  resolveLlmProvider,
} = require("./providers/llm");

const ROOT = path.join(__dirname, "..");
const TEMPLATES_DIR = path.join(ROOT, "templates");
const SCRIPTS_DIR = path.join(ROOT, "scripts");
const storage = getStorageAdapter();
const JOB_LIST_LIMIT = Number(process.env.JOB_LIST_LIMIT || 50);
const LLM_TIMEOUT_MS = Number(process.env.LLM_TIMEOUT_MS || 120000);
const LLM_INPUT_CHAR_LIMIT = Number(process.env.LLM_INPUT_CHAR_LIMIT || 18000);
const INTERVIEW_MEMO_INPUT_CHAR_LIMIT = Number(process.env.INTERVIEW_MEMO_INPUT_CHAR_LIMIT || 60000);
const INTERVIEW_MEMO_MIN_EFFECTIVE_RATIO = Number(process.env.INTERVIEW_MEMO_MIN_EFFECTIVE_RATIO || 0.3);
const INTERVIEW_MEMO_MAX_BACKFILL_LINES = Number(process.env.INTERVIEW_MEMO_MAX_BACKFILL_LINES || 80);
const ROLE_LABEL_CHAR_LIMIT = Number(process.env.ROLE_LABEL_CHAR_LIMIT || 12000);
const ROLE_LABEL_SKIP_CHAR_LIMIT = Number(process.env.ROLE_LABEL_SKIP_CHAR_LIMIT || 20000);
const ENABLE_LLM_ROLE_LABELING = String(process.env.ENABLE_LLM_ROLE_LABELING || "").toLowerCase() === "true";
const PDF_TEXT_PYTHON = process.env.PDF_TEXT_PYTHON || process.env.LOCAL_WHISPER_PYTHON || "python";

async function ensureAppFolders() {
  await storage.ensureFolders();
}

async function listJobs() {
  await ensureAppFolders();
  const files = await storage.listStructuredRecords();
  const jobs = await Promise.all(
    files
      .filter((file) => file.endsWith(".json"))
      .sort((a, b) => (a < b ? 1 : -1))
      .slice(0, JOB_LIST_LIMIT)
      .map(async (file) => {
        const raw = await storage.readStructuredRecord(file);
        return normalizeJobListItem(JSON.parse(raw), file);
      }),
  );

  return jobs
    .filter((job) => job?.job?.createdAt)
    .sort((a, b) => (a.job.createdAt < b.job.createdAt ? 1 : -1));
}

function normalizeJobListItem(record, fileName = "") {
  if (record?.job?.id) {
    return {
      job: {
        id: record.job.id,
        createdAt: record.job.createdAt,
        title: record.job.title || "Untitled meeting",
        company: record.job.company || "Unknown company",
        templateId: record.job.templateId || record.structured?.processing?.templateId || "interview-knowledge-base",
        templateName: record.job.templateName || "Interview Memo",
        summary: record.job.summary || record.structured?.summary?.oneSentence || "",
        modelMode: record.job.modelMode || record.structured?.processing?.modelMode || "unknown",
        llmProvider: record.job.llmProvider || record.structured?.processing?.llmProvider || "none",
        sourceType: record.job.sourceType || record.structured?.processing?.sourceType || "missing",
        asrProvider: record.job.asrProvider || record.structured?.processing?.asrProvider || "manual",
        researchProvider: record.job.researchProvider || record.structured?.processing?.researchProvider || "none",
        transcriptPath: record.job.transcriptPath || "",
        markdownPath: record.job.markdownPath || "",
        pptPath: record.job.pptPath || null,
        structuredPath: record.job.structuredPath || storage.getStructuredPath(`${record.job.id}.json`),
      },
    };
  }

  const structured = record || {};
  const structuredId = structured.id || path.basename(fileName, ".json");
  const processing = structured.processing || {};
  const meta = structured.meeting?.meta || {};
  const templateId = processing.templateId || "interview-knowledge-base";
  return {
    job: {
      id: structuredId,
      createdAt: structured.createdAt || inferCreatedAtFromFileName(fileName),
      title: meta.title || "Untitled meeting",
      company: meta.company || "Unknown company",
      templateId,
      templateName: processing.templateName || templateId,
      summary: structured.summary?.oneSentence || "",
      modelMode: processing.modelMode || "unknown",
      llmProvider: processing.llmProvider || "none",
      sourceType: processing.sourceType || "missing",
      asrProvider: processing.asrProvider || "manual",
      researchProvider: processing.researchProvider || "none",
      transcriptPath: meta.transcriptPath || "",
      markdownPath: storage.getOutputPath(`${structuredId}.md`),
      pptPath: ["weekly-report", "oi-news-report"].includes(templateId)
        ? storage.getOutputPath(`${structuredId}.pptx`)
        : null,
      structuredPath: storage.getStructuredPath(`${structuredId}.json`),
    },
  };
}

function inferCreatedAtFromFileName(fileName = "") {
  const match = String(fileName).match(/^(\d{4}-\d{2}-\d{2}t\d{2}-\d{2}-\d{2})/i);
  if (!match) {
    return new Date(0).toISOString();
  }

  return `${match[1].replace("t", "T").replace(/-(\d{2})-(\d{2})$/, ":$1:$2")}Z`;
}

async function fetchWithTimeout(url, options = {}, timeoutMs = LLM_TIMEOUT_MS) {
  if (typeof AbortSignal !== "undefined" && AbortSignal.timeout) {
    return fetch(url, {
      ...options,
      signal: options.signal || AbortSignal.timeout(timeoutMs),
    });
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      ...options,
      signal: options.signal || controller.signal,
    });
  } finally {
    clearTimeout(timer);
  }
}

function compactTextForLLM(text, maxChars = LLM_INPUT_CHAR_LIMIT) {
  const value = String(text || "").trim();
  if (value.length <= maxChars) {
    return value;
  }

  const lines = value
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  const evidenceLines = lines.filter((line) =>
    /\b(fund|financing|valuation|ipo|pre-ipo|revenue|margin|customer|contract|pilot|product|technology|model|data|architecture|roadmap|manufacturing|deployment|founder|ceo|cto|team|risk|competition)\b|融资|估值|上市|客户|收入|毛利|产品|技术|模型|数据|架构|量产|部署|创始|团队|风险/i.test(line),
  );

  const headBudget = Math.floor(maxChars * 0.34);
  const middleBudget = Math.floor(maxChars * 0.2);
  const tailBudget = Math.floor(maxChars * 0.26);
  const evidenceBudget = maxChars - headBudget - middleBudget - tailBudget - 600;
  const middleStart = Math.max(0, Math.floor(value.length / 2 - middleBudget / 2));

  return [
    `[Long source compacted from ${value.length} chars to fit the LLM request. Preserve facts from all excerpts.]`,
    "",
    "BEGINNING EXCERPT:",
    value.slice(0, headBudget).trim(),
    "",
    "MIDDLE EXCERPT:",
    value.slice(middleStart, middleStart + middleBudget).trim(),
    "",
    "ENDING EXCERPT:",
    value.slice(Math.max(0, value.length - tailBudget)).trim(),
    "",
    "KEY EVIDENCE LINES:",
    evidenceLines.join("\n").slice(0, Math.max(0, evidenceBudget)).trim(),
  ]
    .filter(Boolean)
    .join("\n");
}

async function readJob(jobId) {
  const filePath = `${jobId}.json`;
  const raw = await storage.readStructuredRecord(filePath);
  return normalizeStoredJob(JSON.parse(raw), `${jobId}.json`);
}

async function saveJob(job) {
  return storage.writeStructuredRecord(`${job.job.id}.json`, JSON.stringify(job, null, 2));
}

function toClientJobRecord(record) {
  const structured = record.structured || {};
  const inputs = record.inputs || {};
  return {
    ...record,
    inputs: {
      ...inputs,
      transcriptText: previewText(inputs.transcriptText, 4000),
      roleAnalysis: undefined,
      evidenceBank: undefined,
    },
    structured: {
      id: structured.id,
      createdAt: structured.createdAt,
      processing: structured.processing,
      meeting: {
        meta: structured.meeting?.meta || {},
        transcriptPreview: previewText(structured.meeting?.transcript, 4000),
        materialsPreview: previewText(structured.meeting?.materialsText, 3000),
      },
      materialInsights: structured.materialInsights,
      summary: structured.summary,
      sections: structured.sections,
      memoCategories: structured.memoCategories,
      actionItems: structured.actionItems,
      risks: structured.risks,
      quotes: structured.quotes,
      dataPoints: structured.dataPoints,
      fundingTable: structured.fundingTable,
      uncategorized: structured.uncategorized,
      fundraisingNotes: structured.fundraisingNotes,
      weeklyReportDraft: structured.weeklyReportDraft,
      oiNewsDraft: structured.oiNewsDraft,
      focusProfile: structured.focusProfile,
      templateRecommendation: structured.templateRecommendation,
      renderChecks: structured.renderChecks,
      provenance: structured.provenance,
      research: structured.research,
    },
  };
}

function previewText(value, maxChars = 4000) {
  const text = String(value || "");
  if (text.length <= maxChars) {
    return text;
  }

  return `${text.slice(0, maxChars).trim()}\n\n[Preview truncated. Use the download links for the full artifact.]`;
}

function normalizeStoredJob(record, fileName = "") {
  if (record?.job?.id) {
    if (!record.structured?.renderChecks && record.structured?.processing?.templateId) {
      record.structured.renderChecks = buildRenderChecks(record.structured.processing.templateId, record.structured);
    }
    if (!record.structured?.provenance && record.structured?.processing?.templateId) {
      record.structured.provenance = buildProvenanceMetadata({
        templateId: record.structured.processing.templateId,
        transcriptText: record.structured.meeting?.transcript || "",
        materialText: record.structured.meeting?.materialsText || "",
        research: record.structured.research,
        structured: record.structured,
        selectedEnrichmentFields: [],
      });
    }
    if (!record.structured?.focusProfile && record.structured?.processing?.templateId) {
      record.structured.focusProfile = deriveFocusProfile(record.structured);
    }
    if (!record.structured?.templateRecommendation && record.structured?.processing?.templateId) {
      record.structured.templateRecommendation = deriveTemplateRecommendation(record.structured);
    }
    return record;
  }

  const structured = record || {};
  const structuredId = structured.id || path.basename(fileName, ".json");
  const createdAt = structured.createdAt || new Date().toISOString();
  const meta = structured.meeting?.meta || {};
  const processing = structured.processing || {};
  const templateId = processing.templateId || "interview-knowledge-base";
  const templateName =
    processing.templateName ||
    (templateId === "weekly-report"
      ? "Weekly Report"
      : templateId === "oi-news-report"
        ? "OI News Report"
        : templateId === "interview-free-style"
          ? "Free Style Memo"
          : "Interview Memo");
  const markdownPath = storage.getOutputPath(`${structuredId}.md`);
  const pptPath = ["weekly-report", "oi-news-report"].includes(templateId)
    ? storage.getOutputPath(`${structuredId}.pptx`)
    : null;

  return {
    job: {
      id: structuredId,
      createdAt,
      title: meta.title || "Untitled meeting",
      company: meta.company || "Unknown company",
      templateId,
      templateName,
      transcriptPath: meta.transcriptPath || "",
      markdownPath,
      summary: structured.summary?.oneSentence || "",
      modelMode: processing.modelMode || "unknown",
      llmProvider: processing.llmProvider || "none",
      sourceType: processing.sourceType || "missing",
      asrProvider: processing.asrProvider || "manual",
      researchProvider: processing.researchProvider || "none",
      pptPath,
      structuredPath: storage.getStructuredPath(`${structuredId}.json`),
    },
    inputs: {
      meetingTitle: meta.title || "Untitled meeting",
      company: meta.company || "Unknown company",
      meetingType: meta.meetingType || "general",
      participants: meta.participants || [],
      templateId,
      asrProvider: processing.asrProvider || "manual",
      llmProvider: processing.llmProvider || "none",
      researchProvider: processing.researchProvider || "none",
      notesText: "",
      userInstructions: processing.userInstructions || "",
      transcriptText: structured.meeting?.transcript || "",
      materials: meta.materials || [],
    },
    review: buildCoverageReview({
      templateId,
      sourceText: buildSourcePacket({
        transcriptText: structured.meeting?.transcript || "",
        materialText: structured.meeting?.materialsText || "",
        research: structured.research,
      }),
      structured,
    }),
    structured,
    markdown: "",
    artifacts: {
      markdownPath,
      pptPath,
      obsidianPath: null,
    },
  };
}

function getRuntimeProviders() {
  const obsidian = getObsidianConfig();
  return {
    asr: getAsrProviders(),
    llm: {
      ...getLlmProviders(),
    },
    research: getResearchProviders(),
    storage: {
      mode: storage.mode,
      rootDir: storage.rootDir,
    },
    integrations: {
      obsidian,
      feishu: {
        configured: Boolean(process.env.FEISHU_APP_ID && process.env.FEISHU_APP_SECRET),
      },
    },
  };
}

async function processMeetingInput(payload) {
  return createDraftJob(payload);
}

async function processTranscriptInput(payload) {
  await ensureAppFolders();

  const title = sanitizeText(payload.meetingTitle, "Untitled meeting");
  const company = sanitizeText(payload.company, "Unknown company");
  const requestedAsrProvider = sanitizeText(payload.asrProvider, "auto");
  const userInstructions = sanitizeText(payload.userInstructions);
  const slug = slugify(`${new Date().toISOString()}-${company}-${title}-transcript`);

  const transcriptResolution = await resolveTranscript({
    title,
    company,
    audioFile: payload.audioFile,
    transcriptFile: payload.transcriptFile,
    transcriptText: payload.transcriptText,
    slug,
    asrProvider: requestedAsrProvider,
    userInstructions,
  });

  if (!transcriptResolution.text || transcriptResolution.text.trim().length < 20) {
    throw createError(
      400,
      "The source was processed, but the cleaned transcript is too thin. Please upload a longer audio/transcript file.",
    );
  }

  const transcriptPath = await storage.writeTranscript(`${slug}.txt`, transcriptResolution.text);

  return {
    ok: true,
    transcriptText: transcriptResolution.text,
    transcriptPath,
    sourceType: transcriptResolution.sourceType,
    asrProvider: transcriptResolution.providerId,
  };
}

async function createDraftJob(payload) {
  await ensureAppFolders();

  const title = sanitizeText(payload.meetingTitle, "Untitled meeting");
  const templateId = sanitizeText(payload.templateId);
  const company = sanitizeText(payload.company, "Unknown company");
  const meetingType = sanitizeText(payload.meetingType, "general");
  const participants = splitParticipants(payload.participants);
  const requestedAsrProvider = sanitizeText(payload.asrProvider, "auto");
  const requestedLlmProvider = sanitizeText(payload.llmProvider, "auto");
  const requestedResearchProvider = sanitizeText(payload.researchProvider, "auto");
  const notesText = sanitizeText(payload.notesText);
  const userInstructions = sanitizeText(payload.userInstructions);

  if (!templateId) {
    throw createError(400, "Template is required.");
  }

  const template = await readTemplate(templateId);
  const slug = slugify(`${new Date().toISOString()}-${company}-${title}`);
  const transcriptResolution = await resolveTranscript({
    title,
    company,
    audioFile: payload.audioFile,
    transcriptFile: payload.transcriptFile,
    transcriptText: payload.transcriptText,
    slug,
    asrProvider: requestedAsrProvider,
    userInstructions,
  });
  const materialsResolution = await resolveMaterials({
    slug,
    materialFiles: payload.materialFiles,
    notesText,
    llmProvider: requestedLlmProvider,
  });

  if (
    (!transcriptResolution.text || transcriptResolution.text.trim().length < 50) &&
    materialsResolution.combinedText.trim().length < 80
  ) {
    throw createError(
      400,
      "The uploaded audio was transcribed, but the resulting text is still too thin for note generation. Please add supporting materials or analyst notes, or try a longer recording.",
    );
  }

  const transcriptPath = await storage.writeTranscript(`${slug}.txt`, transcriptResolution.text);

  const structured = await buildStructuredMeeting({
    title,
    company,
    meetingType,
    participants,
    template,
    transcriptText: transcriptResolution.text,
    materialText: materialsResolution.combinedText,
    materials: materialsResolution.materials,
    materialInsights: materialsResolution.materialInsights,
    transcriptPath,
    asrProvider: transcriptResolution.providerId,
    llmProvider: requestedLlmProvider,
    researchProvider: requestedResearchProvider,
    selectedEnrichmentFields: [],
    sourceType: transcriptResolution.sourceType,
    userInstructions,
  });

  const output = renderOutput(template, structured);
  const outputBase = `${slug}-${template.id}`;
  const markdownPath = await storage.writeOutputText(`${outputBase}.md`, output);
  const roleTranscriptPath = await storage.writeOutputText(
    `${outputBase}-role-labeled-transcript.md`,
    renderRoleLabeledTranscript(structured.roleAnalysis),
  );
  const evidenceBankPath = await storage.writeOutputJson(`${outputBase}-evidence-bank.json`, structured.evidenceBank || {});

  const coverageReview = buildCoverageReview({
    templateId: template.id,
    sourceText: buildSourcePacket({
      transcriptText: transcriptResolution.text,
      materialText: materialsResolution.combinedText,
      research: null,
      roleAnalysis: structured.roleAnalysis,
    }),
    structured,
  });

  const jobRecord = {
    job: {
      id: outputBase,
      createdAt: structured.createdAt,
      title: structured.meeting.meta.title,
      company: structured.meeting.meta.company,
      templateId: template.id,
      templateName: template.name,
      transcriptPath,
      roleTranscriptPath,
      evidenceBankPath,
      markdownPath,
      summary: structured.summary.oneSentence,
      modelMode: structured.processing.modelMode,
      llmProvider: structured.processing.llmProvider,
      sourceType: structured.processing.sourceType,
      asrProvider: structured.processing.asrProvider,
      researchProvider: structured.processing.researchProvider,
      pptPath: null,
    },
    inputs: {
      meetingTitle: title,
      company,
      meetingType,
      participants,
      templateId: template.id,
      asrProvider: requestedAsrProvider,
      llmProvider: requestedLlmProvider,
      researchProvider: requestedResearchProvider,
      notesText,
      userInstructions,
      transcriptText: transcriptResolution.text,
      roleAnalysis: structured.roleAnalysis,
      evidenceBank: structured.evidenceBank,
      materials: materialsResolution.materials,
    },
    review: coverageReview,
    structured,
    markdown: output,
    artifacts: {
      transcriptPath,
      roleTranscriptPath,
      evidenceBankPath,
      markdownPath,
      pptPath: null,
      obsidianPath: null,
    },
  };

  const structuredPath = await saveJob(jobRecord);
  jobRecord.job.structuredPath = structuredPath;
  await saveJob(jobRecord);

  return toClientJobRecord(jobRecord);
}

async function resolveTranscript({ title, company, audioFile, transcriptFile, transcriptText, slug, asrProvider, userInstructions }) {
  if (transcriptText && transcriptText.trim()) {
    return {
      text: normalizeTranscriptLikeTextV2(transcriptText.trim()),
      sourceType: "provided-transcript",
      providerId: "manual",
    };
  }

  if (transcriptFile) {
    const originalName = transcriptFile.originalname || `${slug}-transcript.txt`;
    const transcriptInputPath = await storage.writeInputBinary(
      `${slug}-${sanitizeFileName(originalName)}`,
      transcriptFile.buffer,
    );
    const extraction = await extractMaterialText(transcriptFile, "none");
    const extractedText = normalizeTranscriptLikeTextV2(extraction.text || "");

    if (!extractedText) {
      throw createError(
        400,
        `Could not read transcript file "${originalName}". Please upload a .txt, .md, .docx, or readable PDF transcript.`,
      );
    }

    return {
      text: extractedText,
      sourceType: "uploaded-transcript",
      providerId: "manual",
      transcriptInputPath,
      title,
    };
  }

  if (!audioFile) {
    return {
      text: "",
      sourceType: "missing",
      providerId: resolveAsrSelection(asrProvider),
    };
  }

  const originalName = audioFile.originalname || `${slug}.bin`;
  const audioPath = await storage.writeInputBinary(`${slug}-${sanitizeFileName(originalName)}`, audioFile.buffer);

  const transcription = await transcribeAudio({
    providerId: asrProvider,
    audioFile,
    context: {
      title,
      company,
      userInstructions,
    },
  });

  return {
    text: normalizeTranscriptLikeTextV2(transcription.text),
    sourceType: "audio-transcription",
    providerId: transcription.providerId,
    audioPath,
    title,
  };
}

async function resolveMaterials({ slug, materialFiles, notesText, llmProvider }) {
  const materials = [];
  const extractedChunks = [];
  const materialInsights = createEmptyMaterialInsights();

  for (const materialFile of materialFiles || []) {
    const originalName = materialFile.originalname || `${slug}-material.bin`;
    const safeName = `${slug}-${sanitizeFileName(originalName)}`;
    const filePath = await storage.writeMaterialBinary(safeName, materialFile.buffer);

    const extraction = await extractMaterialText(materialFile, llmProvider);
    const extractedText = typeof extraction.text === "string" ? extraction.text : "";
    materials.push({
      originalName,
      mimeType: materialFile.mimetype,
      filePath,
      extractedTextLength: extractedText.length,
      extractionMethod: extraction.method || "unknown",
    });
    mergeMaterialInsights(materialInsights, extraction.insights);

    if (extractedText) {
      extractedChunks.push(`Material: ${originalName}\n${extractedText}`);
    }
  }

  if (notesText) {
    extractedChunks.push(`Notes:\n${notesText}`);
    materials.push({
      originalName: "notesText",
      mimeType: "text/plain",
      filePath: null,
      extractedTextLength: notesText.length,
    });
  }

  return {
    materials,
    combinedText: extractedChunks.join("\n\n").trim(),
    materialInsights,
  };
}

async function extractMaterialText(materialFile, llmProvider) {
  const mimeType = materialFile.mimetype || "";
  const fileName = (materialFile.originalname || "").toLowerCase();

  if (mimeType === "application/pdf" || fileName.endsWith(".pdf")) {
    const text = await extractPdfText(materialFile.buffer);
    return {
      text,
      method: "pdf-parse",
      insights: deriveMaterialInsightsFromText(text),
    };
  }

  if (
    mimeType.startsWith("text/") ||
    fileName.endsWith(".md") ||
    fileName.endsWith(".txt") ||
    fileName.endsWith(".json")
  ) {
    const decodedText = decodeTextBuffer(materialFile.buffer);
    return {
      text: normalizeTranscriptLikeTextV2(decodedText),
      method: "plain-text",
      insights: deriveMaterialInsightsFromText(decodedText),
    };
  }

  if (
    mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    fileName.endsWith(".docx")
  ) {
    const extracted = await extractDocxText(materialFile.buffer);
    const normalized = normalizeTranscriptLikeTextV2(extracted);
    return {
      text: normalized,
      method: "docx-mammoth",
      insights: deriveMaterialInsightsFromText(normalized),
    };
  }

  if (
    mimeType === "application/vnd.openxmlformats-officedocument.presentationml.presentation" ||
    fileName.endsWith(".pptx")
  ) {
    const extracted = await extractPptxText(materialFile.buffer);
    return {
      text: extracted,
      method: "pptx-zip",
      insights: deriveMaterialInsightsFromText(extracted),
    };
  }

  if (
    mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
    mimeType === "application/vnd.ms-excel" ||
    fileName.endsWith(".xlsx") ||
    fileName.endsWith(".xls") ||
    fileName.endsWith(".csv")
  ) {
    const extracted = extractSpreadsheetText(materialFile.buffer);
    return {
      text: extracted,
      method: "spreadsheet-xlsx",
      insights: deriveMaterialInsightsFromText(extracted),
    };
  }

  if (fileName.endsWith(".doc") || fileName.endsWith(".ppt")) {
    return {
      text: "",
      method: "legacy-office-unsupported",
      insights: createEmptyMaterialInsights(),
    };
  }

  if (mimeType.startsWith("image/") || /\.(png|jpg|jpeg|webp|gif|bmp)$/i.test(fileName)) {
    try {
      const result = await extractImageMaterialText({
        requestedProvider: llmProvider,
        mimeType,
        buffer: materialFile.buffer,
        fileName: materialFile.originalname,
      });

      return {
        text: result.text,
        method: result.usedVision ? `${result.providerId}-vision` : "image-unprocessed",
        insights: deriveMaterialInsightsFromText(result.text || ""),
      };
    } catch (_error) {
      return {
        text: "",
        method: "image-unprocessed",
        insights: createEmptyMaterialInsights(),
      };
    }
  }

  return {
    text: "",
    method: "unsupported",
    insights: createEmptyMaterialInsights(),
  };
}

async function extractPdfText(buffer) {
  const pythonText = await extractPdfTextWithPython(buffer).catch(() => "");
  if (pythonText) {
    return pythonText;
  }

  const parser = new PDFParse({ data: buffer });
  let parsedText = "";
  try {
    const parsed = await parser.getText();
    parsedText = cleanExtractedPdfText(parsed.text || "");
  } finally {
    await parser.destroy().catch(() => {});
  }

  return parsedText;
}

async function extractPdfTextWithPython(buffer) {
  const tempPath = path.join(os.tmpdir(), `imma-pdf-${crypto.randomUUID()}.pdf`);
  await fs.writeFile(tempPath, buffer);
  try {
    const output = await runCaptureProcess(PDF_TEXT_PYTHON, [path.join(SCRIPTS_DIR, "extract_pdf_text.py"), tempPath]);
    return cleanExtractedPdfText(output);
  } finally {
    await fs.unlink(tempPath).catch(() => {});
  }
}

function cleanExtractedPdfText(text) {
  return String(text || "")
    .replace(/\u200b/g, "")
    .replace(/\u0001/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function buildStructuredMeeting({
  title,
  company,
  meetingType,
  participants,
  template,
  transcriptText,
  materialText,
  materials,
  materialInsights,
  transcriptPath,
  asrProvider,
  llmProvider,
  researchProvider,
  selectedEnrichmentFields,
  sourceType,
  userInstructions,
}) {
  const terminologyHints = await buildTerminologyHints({
    company,
    title,
    userInstructions,
    limit: 20,
  });
  const roleAnalysis = await buildRoleAwareTranscript({
    title,
    company,
    transcriptText,
    llmProvider,
    userInstructions,
  });
  const companyEvidenceText = buildCompanyEvidenceText(roleAnalysis);
  const research = await maybeEnrichResearch({
    company,
    meetingType,
    templateId: template.id,
    transcriptText: companyEvidenceText || transcriptText,
    materialText,
    researchProvider,
    selectedEnrichmentFields,
  });
  const sourcePacket = buildSourcePacket({
    transcriptText,
    materialText,
    research,
    roleAnalysis: isInterviewMemoTemplate(template.id) ? null : roleAnalysis,
  });
  const llmResult = await tryLLMExtraction({
    title,
    company,
    meetingType,
    participants,
    transcriptText: sourcePacket,
    template,
    llmProvider,
    userInstructions,
    terminologyHints,
  }).catch((error) => ({
    providerId: resolveLlmProvider(llmProvider),
    structured: null,
    error: error.message || "LLM extraction failed.",
  }));
  const llmSucceeded = Boolean(llmResult?.structured);

  const extracted = normalizeExtractedPayload(
    llmResult?.structured || fallbackExtraction(sourcePacket, template),
    template.id,
  );

  const structured = {
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    processing: {
      templateId: template.id,
      templateName: template.name,
      modelMode: llmSucceeded ? "llm" : "fallback",
      llmProvider: llmResult?.providerId || resolveLlmProvider(llmProvider),
      llmError: llmSucceeded ? "" : llmResult?.error || "",
      sourceType,
      asrProvider,
      researchProvider: resolveResearchProvider(researchProvider),
      researchUsed: Boolean(research),
      storageMode: process.env.STORAGE_MODE || "local",
      userInstructions,
      terminologyHints,
    },
    meeting: {
      meta: {
        title,
        company,
        meetingType,
        participants,
        transcriptPath,
        materials,
      },
      transcript: transcriptText,
      roleLabeledTranscript: roleAnalysis.roleLabeledTranscript,
      companyEvidenceText,
      materialsText: materialText,
    },
    roleAnalysis,
    evidenceBank: roleAnalysis.evidenceBank,
    materialInsights: materialInsights || createEmptyMaterialInsights(),
    summary: {
      oneSentence:
        extracted.oneSentence ||
        firstMeaningfulSentence(transcriptText) ||
        `${company} ${meetingType} meeting summary`,
      executiveSummary: extracted.executiveSummary || extracted.paragraphSummary || "",
    },
    sections: extracted.sections || [],
    memoCategories: extracted.memoCategories || [],
    actionItems: extracted.actionItems || [],
    risks: extracted.risks || [],
    quotes: extracted.quotes || [],
    dataPoints: extracted.dataPoints || [],
    fundingTable: extracted.fundingTable || [],
    uncategorized: extracted.uncategorized || [],
    fundraisingNotes: extracted.fundraisingNotes || [],
    research: research
      ? {
          providerId: research.providerId,
          searches: research.searches,
        }
      : null,
    weeklySlide: extracted.weeklySlide || {
      headline: extracted.oneSentence || "",
      updates: extracted.sections?.slice(0, 3).map((section) => section.title) || [],
      risks: extracted.risks || [],
      nextSteps: extracted.actionItems || [],
    },
    weeklyReportDraft:
      hasWeeklyReportDraftContent(extracted.weeklyReportDraft)
        ? extracted.weeklyReportDraft
        : buildWeeklyReportDraft({
        title,
        company,
        meetingType,
        participants,
        sourcePacket,
        extracted,
      }),
    oiNewsDraft: extracted.oiNewsDraft || null,
  };

  if (template.id === "weekly-report" && structured.weeklyReportDraft) {
    structured.weeklyReportDraft = enrichWeeklyReportDraftWithStructuredEvidence(structured.weeklyReportDraft, structured);
  }

  structured.focusProfile = deriveFocusProfile(structured);
  structured.templateRecommendation = deriveTemplateRecommendation(structured);
  structured.renderChecks = buildRenderChecks(template.id, structured);
  structured.provenance = buildProvenanceMetadata({
    templateId: template.id,
    transcriptText,
    materialText,
    research,
    structured,
    selectedEnrichmentFields,
  });

  return structured;
}

async function maybeEnrichResearch({
  company,
  meetingType,
  templateId,
  transcriptText,
  materialText,
  researchProvider,
  selectedEnrichmentFields = [],
}) {
  if (!selectedEnrichmentFields.length) {
    return null;
  }

  return enrichCompanyContext({
    company,
    meetingType,
    templateId,
    transcriptText,
    materialText,
    focusAreas: selectedEnrichmentFields,
    providerId: researchProvider,
  });
}

async function buildRoleAwareTranscript({ title, company, transcriptText, llmProvider, userInstructions }) {
  const cleanedTranscript = normalizeTranscriptLikeTextV2(transcriptText || "");
  if (!cleanedTranscript.trim()) {
    return createFallbackRoleAnalysis("", []);
  }

  const fallbackTranscript = cleanedTranscript;

  if (!ENABLE_LLM_ROLE_LABELING || cleanedTranscript.length > ROLE_LABEL_SKIP_CHAR_LIMIT) {
    return createFallbackRoleAnalysis(fallbackTranscript, splitIntoMemoSentences(fallbackTranscript));
  }

  const llmAnalysis = await tryLLMRoleAnalysis({
    title,
    company,
    transcriptText: cleanedTranscript,
    llmProvider,
    userInstructions,
  }).catch(() => null);

  return normalizeRoleAnalysis(llmAnalysis, llmAnalysis ? cleanedTranscript : fallbackTranscript);
}

async function tryLLMRoleAnalysis({ title, company, transcriptText, llmProvider, userInstructions }) {
  const config = getResolvedLlmConfig(llmProvider);
  if (!config.apiKey || config.providerId === "none") {
    return null;
  }

  const prompt = [
    "Role-label this investment interview transcript for downstream memo generation.",
    `Meeting title: ${title}`,
    `Company: ${company}`,
    `User instructions: ${userInstructions || "None"}`,
    "",
    "Critical rules:",
    "- Identify the interviewer/user/HMG/Hyundai/Cradle side. Their speech is context only unless explicitly analyst POV.",
    "- HMG, Hyundai Motor, Hyundai, Cradle, investor introduction, how we invest, our portfolio, our business, and question setup should NOT become company evidence.",
    "- Company/interviewee statements, data points, claims, and quotes are the primary evidence.",
    "- Interviewer questions can be preserved as context but should not be quoted as company claims.",
    "- If unsure, mark unknown and keep lower confidence.",
    "",
    "Return only JSON with this shape:",
    JSON.stringify({
      turns: [
        {
          role: "company | interviewer_question | interviewer_context | hmg_intro | analyst_pov | unknown",
          text: "cleaned turn text",
          evidenceUse: "primary | context_only | exclude | analyst_only",
          confidence: 0.8,
        },
      ],
      companyEvidence: ["fact or claim from company/interviewee only"],
      interviewerContext: ["question or framing context from interviewer"],
      excludedContext: ["HMG/Hyundai/Cradle intro or non-evidence talk"],
      analystPov: ["analyst/user POV only if explicitly present"],
    }),
    "",
    "Transcript:",
    compactTextForLLM(transcriptText, ROLE_LABEL_CHAR_LIMIT),
  ].join("\n");

  const response = await fetchWithTimeout(`${config.baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: config.model,
      temperature: 0.1,
      ...jsonResponseFormatForProvider(config.providerId),
      messages: [
        {
          role: "system",
          content: "You label interview transcript speaker roles. Return only valid JSON. Do not invent facts.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
    }),
  });

  if (!response.ok) {
    return null;
  }

  const payload = await response.json();
  const content = payload.choices?.[0]?.message?.content;
  if (!content) {
    return null;
  }

  return parseJsonContent(content);
}

function normalizeRoleAnalysis(rawAnalysis, transcriptText) {
  const fallback = createFallbackRoleAnalysis(transcriptText, splitIntoMemoSentences(transcriptText));
  if (!rawAnalysis || typeof rawAnalysis !== "object") {
    return fallback;
  }

  const turns = Array.isArray(rawAnalysis.turns)
    ? rawAnalysis.turns
        .map((turn) => {
          const role = normalizeSpeakerRole(turn?.role);
          const text = cleanTranscriptArtifact(turn?.text || "");
          const evidenceUse = normalizeEvidenceUse(turn?.evidenceUse, role);
          if (!text) {
            return null;
          }
          return {
            role,
            text,
            evidenceUse,
            confidence: Number.isFinite(Number(turn?.confidence)) ? Number(turn.confidence) : 0.65,
          };
        })
        .filter(Boolean)
    : fallback.turns;

  const companyEvidence = sanitizeRoleLines(rawAnalysis.companyEvidence).length
    ? sanitizeRoleLines(rawAnalysis.companyEvidence, 80)
    : turns.filter((turn) => turn.evidenceUse === "primary").map((turn) => turn.text);
  const interviewerContext = sanitizeRoleLines(rawAnalysis.interviewerContext, 50);
  const excludedContext = sanitizeRoleLines(rawAnalysis.excludedContext, 50);
  const analystPov = sanitizeRoleLines(rawAnalysis.analystPov, 30);

  return {
    provider: rawAnalysis.provider || "llm",
    turns,
    companyEvidence,
    interviewerContext,
    excludedContext,
    analystPov,
    roleLabeledTranscript: renderRoleTurns(turns),
    evidenceBank: {
      companyEvidence,
      interviewerContext,
      excludedContext,
      analystPov,
    },
  };
}

function createFallbackRoleAnalysis(transcriptText, sentences) {
  const turns = (sentences.length ? sentences : splitIntoMemoSentences(transcriptText)).map((sentence) => {
    const role = inferFallbackSpeakerRole(sentence);
    return {
      role,
      text: sentence,
      evidenceUse: normalizeEvidenceUse("", role),
      confidence: role === "unknown" ? 0.35 : 0.55,
    };
  });
  const companyEvidence = turns.filter((turn) => turn.evidenceUse === "primary").map((turn) => turn.text);
  const interviewerContext = turns.filter((turn) => turn.evidenceUse === "context_only").map((turn) => turn.text);
  const excludedContext = turns.filter((turn) => turn.evidenceUse === "exclude").map((turn) => turn.text);
  const analystPov = turns.filter((turn) => turn.evidenceUse === "analyst_only").map((turn) => turn.text);

  return {
    provider: "fallback",
    turns,
    companyEvidence,
    interviewerContext,
    excludedContext,
    analystPov,
    roleLabeledTranscript: renderRoleTurns(turns),
    evidenceBank: {
      companyEvidence,
      interviewerContext,
      excludedContext,
      analystPov,
    },
  };
}

function normalizeSpeakerRole(role) {
  const value = cleanGeneratedText(role).toLowerCase();
  if (/company|interviewee|management|founder|expert/.test(value)) {
    return "company";
  }
  if (/question/.test(value)) {
    return "interviewer_question";
  }
  if (/hmg|hyundai|cradle|intro|exclude/.test(value)) {
    return "hmg_intro";
  }
  if (/analyst|pov|opinion/.test(value)) {
    return "analyst_pov";
  }
  if (/interviewer|context/.test(value)) {
    return "interviewer_context";
  }
  return "unknown";
}

function inferFallbackSpeakerRole(text) {
  const line = cleanGeneratedText(text);
  if (looksLikeHmgIntro(line)) {
    return "hmg_intro";
  }
  if (/\?$|^(can you|could you|how do|what do|why do|when do|do you|tell me|maybe can|let's talk|help me understand)\b/i.test(line)) {
    return "interviewer_question";
  }
  if (/\b(i think|my view|our view|we think|from our side|as an investor)\b/i.test(line)) {
    return "analyst_pov";
  }
  if (/\b(management said|founder said|ceo said|cto said|company said|we have|we are|we built|our product|our revenue|our customers)\b/i.test(line)) {
    return "company";
  }
  return "company";
}

function looksLikeHmgIntro(text) {
  return /\b(HMG|Hyundai Motor|Hyundai|Cradle|we invest|our investment|our portfolio|from our side|introduce ourselves|we are looking at|how we invest|corporate venture|CVC)\b/i.test(
    text,
  );
}

function normalizeEvidenceUse(value, role) {
  const explicit = cleanGeneratedText(value).toLowerCase();
  if (/primary/.test(explicit)) {
    return "primary";
  }
  if (/analyst/.test(explicit)) {
    return "analyst_only";
  }
  if (/exclude/.test(explicit)) {
    return "exclude";
  }
  if (/context/.test(explicit)) {
    return "context_only";
  }
  if (role === "company") {
    return "primary";
  }
  if (role === "analyst_pov") {
    return "analyst_only";
  }
  if (role === "hmg_intro") {
    return "exclude";
  }
  if (role === "interviewer_question" || role === "interviewer_context") {
    return "context_only";
  }
  return "primary";
}

function sanitizeRoleLines(lines, limit = 80) {
  return sanitizeOutputLines(Array.isArray(lines) ? lines : [], limit);
}

function buildCompanyEvidenceText(roleAnalysis) {
  const companyEvidence = roleAnalysis?.companyEvidence || [];
  if (!companyEvidence.length) {
    return "";
  }
  return companyEvidence.map((line) => `- ${line}`).join("\n");
}

function renderRoleTurns(turns) {
  return (turns || [])
    .map((turn, index) => `### ${index + 1}. ${turn.role} | ${turn.evidenceUse} | confidence ${turn.confidence}\n${turn.text}`)
    .join("\n\n");
}

function renderRoleLabeledTranscript(roleAnalysis) {
  return [
    "# Role-Labeled Transcript",
    "",
    "## Company Evidence",
    ...(roleAnalysis?.companyEvidence || []).map((line) => `- ${line}`),
    "",
    "## Interviewer Context",
    ...(roleAnalysis?.interviewerContext || []).map((line) => `- ${line}`),
    "",
    "## Excluded HMG / Interviewer Intro",
    ...(roleAnalysis?.excludedContext || []).map((line) => `- ${line}`),
    "",
    "## Analyst POV",
    ...(roleAnalysis?.analystPov || []).map((line) => `- ${line}`),
    "",
    "## Full Role-Labeled Turns",
    roleAnalysis?.roleLabeledTranscript || "",
  ].join("\n");
}

function buildSourcePacket({ transcriptText, materialText, research, roleAnalysis }) {
  const chunks = [];

  const companyEvidenceText = buildCompanyEvidenceText(roleAnalysis);
  if (companyEvidenceText) {
    chunks.push(`Transcript:\n${companyEvidenceText}`);
  } else if (transcriptText?.trim()) {
    chunks.push(`Transcript:\n${transcriptText.trim()}`);
  }

  if (materialText?.trim()) {
    chunks.push(`Supporting materials:\n${materialText.trim()}`);
  }

  if (research?.searches?.length) {
    const researchLines = research.searches.flatMap((search) => {
      const answerLine = search.answer ? [`Answer: ${search.answer}`] : [];
      const resultLines = search.results.map(
        (item, index) => `Source ${index + 1}: ${item.title} | ${item.url} | ${item.content || ""}`,
      );

      return [`Query: ${search.query}`, ...answerLine, ...resultLines];
    });

    chunks.push(`Public research enrichment:\n${researchLines.join("\n")}`);
  }

  return chunks.join("\n\n").trim();
}

function buildWeeklyReportDraft({ title, company, meetingType, participants, sourcePacket, extracted }) {
  const foundedYear = firstRegexMatch(sourcePacket, /\b(19|20)\d{2}\b/g) || "TBD";
  const hqCity =
    firstRegexMatch(sourcePacket, /\b(Shanghai|Beijing|Shenzhen|Suzhou|Hangzhou|Singapore|Seoul|Tokyo|San Francisco|New York)\b/gi) ||
    "TBD";
  const category =
    inferCategory(sourcePacket, meetingType) || title || company;

  const extractedLines = [
    ...(extracted.sections || []).flatMap((section) => [
      ...(section.bullets || []).map((bullet) => `${section.title}: ${bullet}`),
    ]),
    ...(extracted.dataPoints || []),
    ...(extracted.fundraisingNotes || []),
    ...(extracted.quotes || []).map((quote) => `Management POV: "${stripWrappingQuotes(quote)}"`),
  ];
  const weeklySource = dedupeLines([...extractedLines, ...sourcePacket.split(/\r?\n+/)])
    .map(cleanGeneratedText)
    .filter(Boolean)
    .filter((line) => !/^#|^---$|^type:|^company:|^date:|^meeting_type:|^participants:|^source:|^analyst:|^tags:/i.test(line))
    .filter((line) => !/(?:Source Notes|User-provided materials|Public-source enrichment|Inference \/ synthesis|Current Interview Memo Content|Current Memo Facts|Company Evidence Lines|User Weekly Report Instructions|Derived weekly-report source packet|Supporting materials|Material: notesText)\b/i.test(line))
    .filter((line) => !/do not copy labels|Create a weekly report|overall prompt|preserve all|highlight/i.test(line));

  const sectionCandidates = [
    {
      label: "Company Background",
      patterns: [/builds|company|founded|hq|headquarter|team|founder|ceo|cto|chief scientist|leadership|background|category|market position/i],
    },
    {
      label: "Product / Technology",
      patterns: [/platform|product|tech|technology|architecture|model|data|sensor|software|hardware|roadmap|rollout|sku|payload|range|lidar|orin|bom|scenario/i],
    },
    {
      label: "Commercialization / Funding",
      patterns: [/market|customer(?!-specific)|partner|commercial|commercialization|gtm|go-to-market|cities|mileage|tender|backlog|traction|pilot|revenue|pricing|unit economics|payback|cost|ipo|listing|pre-ipo|valuation|fundraising|financing|raise|investor|shareholder|capital-market|hmg|hyundai|singapore|china|collaboration|strategic|partnership|expansion|ecosystem/i],
    },
  ];

  const usedLines = new Set();
  const sections = sectionCandidates
    .map((candidate) => {
      const lines = weeklySource
        .filter((line) => candidate.patterns.some((pattern) => pattern.test(line)))
        .filter((line) => {
          const key = line.toLowerCase();
          if (usedLines.has(key)) {
            return false;
          }
          usedLines.add(key);
          return true;
        })
        .map((line) => toWeeklySlideLine(line))
        .filter(Boolean);
      return { label: candidate.label, lines: padWeeklySectionLines(candidate.label, dedupeLines(lines), weeklySource) };
    })
    .slice(0, 3);

  return {
    title: `[Startup] ${company}`,
    meetingInformation: {
      date: "TBD",
      participants: participants.join(", ") || "TBD",
    },
    companyDescription: {
      line1: extracted.oneSentence || `${company} update based on the current source packet`,
      foundedYear,
      hqCity,
    },
    category,
    sourceContact: participants[0] || "TBD",
    opinion: extracted.oneSentence || `${company} requires follow-up`,
    nextStep: (extracted.actionItems || []).slice(0, 2).join("; ") || "Confirm key open questions and refresh next draft.",
    sections,
  };
}

function getCoverageChecks(templateId) {
  return templateId === "weekly-report" || templateId === "oi-news-report"
    ? [
        { id: "team", label: "Team / leadership", patterns: [/team|founder|chief scientist|ceo|cto|co-founder/i] },
        { id: "product", label: "Product / platform", patterns: [/platform|product|sku|payload|range|tech|lidar|orin/i] },
        { id: "commercialization", label: "Commercialization / traction", patterns: [/customer|partner|deployment|backlog|mileage|cities|commercial/i] },
        { id: "market", label: "Market / GTM context", patterns: [/market|usecase|distribution|express|tender/i] },
        { id: "funding", label: "Funding history", patterns: [/funding|raised|valuation|round|investor/i] },
        { id: "ipo", label: "IPO / listing plan", patterns: [/ipo|listing|pre-ipo|hong kong|filing/i] },
      ]
    : [
        { id: "team", label: "Team / management", patterns: [/founder|ceo|cto|management|team|executive/i] },
        { id: "timeline", label: "Timeline / milestones", patterns: [/founded|launch|pilot|milestone|timeline|roadmap/i] },
        { id: "business", label: "Business / strategy", patterns: [/business|gtm|go-to-market|customer|strategy|commercial/i] },
        { id: "product", label: "Product / technology", patterns: [/product|platform|technology|tech|manufacturing|deployment/i] },
        { id: "funding", label: "Funding / valuation", patterns: [/funding|raised|valuation|round|investor|shareholder|pre-ipo/i] },
        { id: "risks", label: "Risks / review queue", patterns: [/risk|challenge|concern|verify|open question|follow-up|uncertain/i] },
      ];
}

function buildProvenanceMetadata({
  templateId,
  transcriptText,
  materialText,
  research,
  structured,
  selectedEnrichmentFields = [],
}) {
  const coverage = getCoverageChecks(templateId).map((check) => buildFieldSource(check, transcriptText, materialText, research, structured));

  return {
    packetSummary: {
      transcriptChars: String(transcriptText || "").trim().length,
      materialChars: String(materialText || "").trim().length,
      researchEnabled: Boolean(research),
      researchProvider: research?.providerId || "none",
    },
    fieldSources: Object.fromEntries(
      coverage.map((item) => [
        item.id,
        {
          label: item.label,
          status: item.status,
          source: item.source,
          evidence: item.evidence,
        },
      ]),
    ),
    enhancementHistory: selectedEnrichmentFields.length
      ? [
          {
            timestamp: new Date().toISOString(),
            fields: selectedEnrichmentFields,
            provider: research?.providerId || "none",
          },
        ]
      : [],
  };
}

function buildFieldSource(check, transcriptText, materialText, research, structured) {
  const transcriptHits = extractEvidenceSnippets(transcriptText, check.patterns, 2);
  const materialHits = extractEvidenceSnippets(materialText, check.patterns, 2);
  const researchHits = extractEvidenceSnippets((research?.searches || []).map((item) => item.summary || "").join("\n"), check.patterns, 2);
  const structuredHits = extractStructuredEvidence(structured, check.patterns, 2);

  const source =
    transcriptHits.length || materialHits.length
      ? "user-provided"
      : researchHits.length
        ? "public-source-enriched"
        : structuredHits.length
          ? "model-synthesized"
          : "missing";

  const status =
    source === "missing"
      ? "missing"
      : source === "model-synthesized"
        ? "thin"
        : "covered";

  return {
    id: check.id,
    label: check.label,
    status,
    source,
    evidence: {
      transcript: transcriptHits,
      materials: materialHits,
      research: researchHits,
      structured: structuredHits,
    },
  };
}

function extractEvidenceSnippets(text, patterns, limit = 2) {
  const lines = String(text || "")
    .split(/\r?\n+/)
    .map((line) => cleanTranscriptArtifact(line))
    .filter(Boolean);

  return dedupeLines(
    lines.filter((line) => patterns.some((pattern) => pattern.test(line))).slice(0, limit),
  );
}

function extractStructuredEvidence(structured, patterns, limit = 2) {
  const lines = [
    ...(structured.sections || []).flatMap((section) => [section.title, ...(section.bullets || [])]),
    ...(structured.dataPoints || []),
    ...(structured.risks || []),
    ...(structured.actionItems || []),
    ...(structured.quotes || []),
    ...(structured.uncategorized || []),
    ...(structured.fundraisingNotes || []),
  ]
    .map(cleanGeneratedText)
    .filter(Boolean);

  return dedupeLines(lines.filter((line) => patterns.some((pattern) => pattern.test(line))).slice(0, limit));
}

function buildRenderChecks(templateId, structured) {
  if (templateId === "weekly-report") {
    const sections = structured.weeklyReportDraft?.sections || [];
    const overBudgetSections = sections
      .map((section) => ({
        label: section.label,
        lineCount: (section.lines || []).length,
        tooLongLines: (section.lines || []).filter((line) => String(line || "").length > 145).length,
      }))
      .filter((section) => section.lineCount > 8 || section.tooLongLines > 0);

    return {
      templateId,
      status: overBudgetSections.length ? "warning" : "ok",
      warnings: overBudgetSections.map((section) =>
        `${section.label} is above the slide-safe budget (${section.lineCount} lines, ${section.tooLongLines} long lines).`,
      ),
    };
  }

  if (templateId === "oi-news-report") {
    const sections = structured.oiNewsDraft?.sections || [];
    const warnings = sections.flatMap((section) => {
      const lines = section.lines || [];
      const sectionWarnings = [];
      if (lines.length > 3) {
        sectionWarnings.push(`${section.label} exceeds the portrait-slide line budget.`);
      }
      if (lines.some((line) => String(line || "").length > 130)) {
        sectionWarnings.push(`${section.label} still contains a line that may wrap too aggressively.`);
      }
      return sectionWarnings;
    });

    return {
      templateId,
      status: warnings.length ? "warning" : "ok",
      warnings,
    };
  }

  const warnings = [];
  if ((structured.uncategorized || []).length > 12) {
    warnings.push("Many notes remain uncategorized. Consider adding supporting materials or running enhancement.");
  }
  if ((structured.fundingTable || []).length === 0 && (structured.fundraisingNotes || []).length > 0) {
    warnings.push("Funding notes were preserved, but the fundraising table still looks thin.");
  }

  return {
    templateId,
    status: warnings.length ? "warning" : "ok",
    warnings,
  };
}

function deriveFocusProfile(structured) {
  const sections = Array.isArray(structured?.sections) ? structured.sections : [];
  const sectionTitles = sections.map((section) => String(section?.title || ""));
  const bulletLines = sections.flatMap((section) => section?.bullets || []);
  const corpus = [
    ...sectionTitles,
    ...bulletLines,
    ...(structured?.fundraisingNotes || []),
    ...(structured?.dataPoints || []),
    ...(structured?.quotes || []),
    ...(structured?.uncategorized || []),
    structured?.meeting?.transcript || "",
    structured?.meeting?.materialsText || "",
    structured?.summary?.executiveSummary || "",
    structured?.summary?.oneSentence || "",
  ]
    .map((item) => cleanGeneratedText(item))
    .filter(Boolean)
    .join("\n");

  const buckets = [
    {
      id: "tech-deep-dive",
      label: "Tech Deep Dive",
      patterns: [/tech|technical|architecture|model|training|inference|algorithm|sensor|lidar|perception|autonomy|stack|engineering|deployment|roadmap|product platform/i],
    },
    {
      id: "founder-background",
      label: "Founder / Management Background",
      patterns: [/founder|ceo|cto|chairman|background|career|experience|resume|joined|worked at|education|previous role|previously/i],
    },
    {
      id: "fundraising-pre-ipo",
      label: "Fundraising / Pre-IPO",
      patterns: [/\b(fundraising|financing|raised|round|valuation|investor|shareholder|ipo|pre-ipo|listing|filing)\b|cap table|capital markets/i],
    },
    {
      id: "gtm-customer-traction",
      label: "GTM / Customer Traction",
      patterns: [/customer|commercial|commercialization|gtm|go-to-market|sales|pipeline|partner|market|traction|deployment|order|backlog|revenue/i],
    },
    {
      id: "mixed-diligence",
      label: "Mixed Diligence",
      patterns: [/business model|strategy|competition|unit economics|manufacturing|supply chain|risk|margin|cost|expansion/i],
    },
  ];

  const scored = buckets
    .map((bucket) => ({
      ...bucket,
      score: scoreTextAgainstPatterns(corpus, bucket.patterns),
    }))
    .sort((a, b) => b.score - a.score);

  const primary = scored[0];
  const secondary = scored.filter((bucket) => bucket.id !== primary?.id && bucket.score > 0).slice(0, 2);
  const dominant = primary?.score || 0;
  const total = scored.reduce((sum, bucket) => sum + bucket.score, 0);
  const dominanceRatio = total ? dominant / total : 0;

  return {
    primaryFocus: primary?.id || "mixed-diligence",
    primaryLabel: primary?.label || "Mixed Diligence",
    secondaryFocuses: secondary.map((item) => item.id),
    secondaryLabels: secondary.map((item) => item.label),
    scoredFocuses: scored.filter((item) => item.score > 0).map((item) => ({
      id: item.id,
      label: item.label,
      score: item.score,
    })),
    dominanceRatio,
    isNarrowInterview: dominanceRatio >= 0.58,
  };
}

function deriveTemplateRecommendation(structured) {
  const focusProfile = structured?.focusProfile || deriveFocusProfile(structured);
  const availableSectionCount = (structured?.sections || []).filter((section) => (section?.bullets || []).length).length;
  const classicCoverageScore = [
    hasKeywordEvidence(structured, [/founder|ceo|cto|management|team|executive/i]),
    hasKeywordEvidence(structured, [/business|strategy|customer|commercial|gtm|market/i]),
    hasKeywordEvidence(structured, [/product|technology|platform|engineering|roadmap|manufacturing/i]),
    hasKeywordEvidence(structured, [/\b(fundraising|valuation|round|investor|ipo|pre-ipo|revenue|margin)\b/i]),
  ].filter(Boolean).length;

  const recommendedTemplateId =
    focusProfile.isNarrowInterview || (classicCoverageScore <= 2 && availableSectionCount <= 2)
      ? "interview-free-style"
      : "interview-knowledge-base";

  return {
    recommendedTemplateId,
    recommendedTemplateName: recommendedTemplateId === "interview-free-style" ? "Free Style Memo" : "Interview Memo",
    reason:
      recommendedTemplateId === "interview-free-style"
        ? `This interview is concentrated on ${focusProfile.primaryLabel.toLowerCase()} and does not look like full company coverage.`
        : "This interview appears broad enough to support the fuller Interview Memo structure.",
  };
}

function scoreTextAgainstPatterns(text, patterns) {
  const haystack = String(text || "");
  return patterns.reduce((sum, pattern) => {
    const matches = haystack.match(new RegExp(pattern.source, pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`));
    return sum + (matches?.length || 0);
  }, 0);
}

function hasKeywordEvidence(structured, patterns) {
  const corpus = [
    ...(structured?.sections || []).flatMap((section) => [section?.title || "", ...(section?.bullets || [])]),
    ...(structured?.fundraisingNotes || []),
    ...(structured?.uncategorized || []),
    ...(structured?.dataPoints || []),
    structured?.meeting?.transcript || "",
    structured?.meeting?.materialsText || "",
    structured?.summary?.executiveSummary || "",
    structured?.summary?.oneSentence || "",
  ]
    .map((item) => cleanGeneratedText(item))
    .filter(Boolean)
    .join("\n");

  return patterns.some((pattern) => pattern.test(corpus));
}

function buildProvenanceSummary(provenance) {
  const fieldSources = Object.values(provenance?.fieldSources || {});
  if (!fieldSources.length) {
    return "No provenance summary available yet.";
  }

  const counts = fieldSources.reduce(
    (acc, item) => {
      acc[item.source] = (acc[item.source] || 0) + 1;
      return acc;
    },
    { "user-provided": 0, "public-source-enriched": 0, "model-synthesized": 0, missing: 0 },
  );

  return `User-provided: ${counts["user-provided"]}; Public-source enriched: ${counts["public-source-enriched"]}; Model-synthesized: ${counts["model-synthesized"]}; Missing: ${counts.missing}.`;
}

function buildCoverageReview({ templateId, sourceText, structured }) {
  const checks = getCoverageChecks(templateId);
  const fieldSources = structured?.provenance?.fieldSources || {};
  const renderWarnings = structured?.renderChecks?.warnings || [];
  const focusProfile = structured?.focusProfile || deriveFocusProfile(structured);
  const recommendation = structured?.templateRecommendation || deriveTemplateRecommendation(structured);
  const reviewItems = checks.map((check) => {
    const coverage = fieldSources[check.id] || buildFieldSource(check, sourceText, "", null, structured);
    const enrichable = ["team", "commercialization", "market", "funding", "ipo", "product", "timeline", "business", "risks"].includes(check.id);
    const sourceLabel =
      coverage.source === "user-provided"
        ? "User-provided"
        : coverage.source === "public-source-enriched"
          ? "Public-source enriched"
          : coverage.source === "model-synthesized"
            ? "Model-synthesized"
            : "Missing";

    return {
      id: check.id,
      label: check.label,
      status: coverage.status,
      enrichable,
      source: sourceLabel,
      evidencePreview: buildReviewEvidencePreview(coverage),
      proposedAddition: buildProposedEnhancementText(check, coverage, enrichable),
      recommendation:
        coverage.status === "covered"
          ? "Looks grounded in the current packet."
          : coverage.status === "thin"
            ? enrichable
              ? "Some content exists, but it still looks thin. Enhancement can strengthen this field."
              : "Some content exists, but it should be reviewed manually."
            : enrichable
              ? "Missing from the current packet. Enhancement can pull public context."
              : "Needs manual input.",
      evidenceCount:
        (coverage.evidence?.transcript?.length || 0) +
        (coverage.evidence?.materials?.length || 0) +
        (coverage.evidence?.research?.length || 0) +
        (coverage.evidence?.structured?.length || 0),
    };
  });

  return {
    summary: reviewItems.some((item) => item.status !== "covered")
      ? templateId === "weekly-report" || templateId === "oi-news-report"
        ? "Some slide fields are still missing or thin. Review provenance before exporting."
        : "Some interview-memo fields are still missing or thin. Review provenance before finalizing."
      : templateId === "weekly-report" || templateId === "oi-news-report"
        ? "Core slide fields look covered from the current packet."
        : "Core interview-memo fields look covered from the current packet.",
    items: reviewItems,
    warnings: renderWarnings,
    recommendation,
    focusProfile,
    provenanceSummary: buildProvenanceSummary(structured?.provenance),
  };
}

function buildReviewEvidencePreview(coverage) {
  const userProvided = [
    ...(coverage.evidence?.transcript || []),
    ...(coverage.evidence?.materials || []),
  ].slice(0, 2);
  const publicSource = (coverage.evidence?.research || []).slice(0, 2);
  const currentDraft = (coverage.evidence?.structured || []).slice(0, 2);

  return {
    userProvided,
    publicSource,
    currentDraft,
  };
}

function buildProposedEnhancementText(check, coverage, enrichable) {
  const publicSource = coverage.evidence?.research || [];
  if (publicSource.length) {
    return `Use public-source context: ${truncateForSlide(publicSource[0], 180)}`;
  }

  const currentDraft = coverage.evidence?.structured || [];
  if (coverage.status === "covered") {
    return currentDraft.length
      ? `Already covered in current draft: ${truncateForSlide(currentDraft[0], 160)}`
      : "Already covered by user-provided materials.";
  }

  if (!enrichable) {
    return "Manual input recommended.";
  }

  return `Prepared enhancement: search and add source-backed ${check.label.toLowerCase()} context before refreshing the weekly report.`;
}

async function finalizeJob(jobId, { selectedEnrichmentFields = [], llmProvider, researchProvider }) {
  const jobRecord = await readJob(jobId);

  const template = await readTemplate(jobRecord.job.templateId);
  const participants = Array.isArray(jobRecord.inputs.participants)
    ? jobRecord.inputs.participants
    : splitParticipants(jobRecord.inputs.participants?.join ? jobRecord.inputs.participants.join(",") : "");

  const materialsText =
    jobRecord.structured.meeting.materialsText ||
    [jobRecord.inputs.notesText || "", ...(jobRecord.inputs.materials || []).map((item) => item.extractedText || "")]
      .filter(Boolean)
      .join("\n\n");

  const structured = await buildStructuredMeeting({
    title: jobRecord.inputs.meetingTitle,
    company: jobRecord.inputs.company,
    meetingType: jobRecord.inputs.meetingType,
    participants,
    template,
    transcriptText: jobRecord.inputs.transcriptText || "",
    materialText: materialsText,
    materials: jobRecord.inputs.materials || [],
    materialInsights: jobRecord.structured.materialInsights || createEmptyMaterialInsights(),
    transcriptPath: jobRecord.job.transcriptPath,
    asrProvider: jobRecord.job.asrProvider,
    llmProvider: llmProvider || jobRecord.job.llmProvider,
    researchProvider: researchProvider || jobRecord.job.researchProvider,
    selectedEnrichmentFields,
    sourceType: jobRecord.job.sourceType,
    userInstructions: jobRecord.inputs.userInstructions || "",
  });
  structured.provenance = {
    ...(structured.provenance || {}),
    enhancementHistory: [
      ...((jobRecord.structured?.provenance?.enhancementHistory || []).filter(Boolean)),
      ...((structured.provenance?.enhancementHistory || []).filter(Boolean)),
    ],
  };

  const markdown = renderOutput(template, structured);
  await storage.writeOutputText(path.basename(jobRecord.artifacts.markdownPath), markdown);

  let pptPath = null;
  if (template.id === "weekly-report") {
    const pptDraft = buildWeeklyPptDraftFromStructured(structured);
    const pptDraftPath = await storage.writeOutputJson(`${jobId}-ppt-draft.json`, pptDraft);
    pptPath = storage.getOutputPath(`${jobId}.pptx`);
    await renderWeeklyReportPpt({
      draftPath: pptDraftPath,
      outputPath: pptPath,
    });
  } else if (template.id === "oi-news-report") {
    const pptDraft = buildOiNewsPptDraftFromStructured(structured);
    const pptDraftPath = await storage.writeOutputJson(`${jobId}-ppt-draft.json`, pptDraft);
    pptPath = storage.getOutputPath(`${jobId}.pptx`);
    await renderOiNewsReportPpt({
      draftPath: pptDraftPath,
      outputPath: pptPath,
    });
  }

  const review = buildCoverageReview({
    templateId: template.id,
    sourceText: buildSourcePacket({
      transcriptText: jobRecord.inputs.transcriptText || "",
      materialText: materialsText,
      research: structured.research,
      roleAnalysis: structured.roleAnalysis,
    }),
    structured,
  });

  const updatedRecord = {
    ...jobRecord,
    review,
    structured,
    markdown,
    artifacts: {
      ...jobRecord.artifacts,
      markdownPath: jobRecord.artifacts.markdownPath,
      pptPath,
      obsidianPath: jobRecord.artifacts?.obsidianPath || null,
    },
    job: {
      ...jobRecord.job,
      summary: structured.summary.oneSentence,
      modelMode: structured.processing.modelMode,
      llmProvider: structured.processing.llmProvider,
      researchProvider: structured.processing.researchProvider,
      pptPath,
    },
  };

  await saveJob(updatedRecord);
  return toClientJobRecord(updatedRecord);
}

async function deriveOutputFromJob(jobId, { targetTemplateId = "weekly-report", materialFiles = [], notesText = "", userInstructions = "", llmProvider, researchProvider = "none" }) {
  const sourceJob = await readJob(jobId);
  const template = await readTemplate(targetTemplateId);
  const sourceStructured = sourceJob.structured || {};
  const sourceMeta = sourceStructured.meeting?.meta || {};
  const participants = Array.isArray(sourceJob.inputs?.participants)
    ? sourceJob.inputs.participants
    : splitParticipants(sourceJob.inputs?.participants?.join ? sourceJob.inputs.participants.join(",") : "");
  const slug = slugify(`${new Date().toISOString()}-${sourceJob.job.company || sourceMeta.company || "company"}-${targetTemplateId}`);
  const evidencePromptText = renderEvidenceBankForPrompt(sourceStructured.evidenceBank);
  const structuredMemoFacts = renderStructuredMemoFactsForWeekly(sourceStructured);
  const derivedNotes = [
    "Derived weekly-report source packet. Use this packet as evidence; do not copy labels, markdown headings, or internal workflow names into the slide.",
    structuredMemoFacts ? `Current Memo Facts:\n${structuredMemoFacts}` : "",
    evidencePromptText ? `Company Evidence Lines:\n${evidencePromptText}` : "",
    notesText,
  ]
    .filter(Boolean)
    .join("\n\n");
  const additionalMaterials = await resolveMaterials({
    slug,
    materialFiles,
    notesText: derivedNotes,
    llmProvider: llmProvider || sourceJob.job.llmProvider,
  });
  const materialText = [
    sourceStructured.meeting?.materialsText || "",
    additionalMaterials.combinedText,
  ]
    .filter(Boolean)
    .join("\n\n");
  const mergedInsights = mergeMaterialInsightSets(sourceStructured.materialInsights, additionalMaterials.materialInsights);

  const structured = await buildStructuredMeeting({
    title: `${sourceJob.job.title || sourceMeta.title || "Meeting"} -> ${template.name}`,
    company: sourceJob.job.company || sourceMeta.company || sourceJob.inputs.company || "Unknown company",
    meetingType: sourceJob.inputs.meetingType || sourceMeta.meetingType || "management-interview",
    participants,
    template,
    transcriptText: sourceJob.inputs.transcriptText || sourceStructured.meeting?.transcript || "",
    materialText,
    materials: [...(sourceJob.inputs.materials || []), ...additionalMaterials.materials],
    materialInsights: mergedInsights,
    transcriptPath: sourceJob.job.transcriptPath,
    asrProvider: sourceJob.job.asrProvider,
    llmProvider: llmProvider || sourceJob.job.llmProvider,
    researchProvider,
    selectedEnrichmentFields: [],
    sourceType: sourceJob.job.sourceType,
    userInstructions,
  });

  const output = renderOutput(template, structured);
  const outputBase = `${slug}-${template.id}`;
  const markdownPath = await storage.writeOutputText(`${outputBase}.md`, output);
  const roleTranscriptPath = await storage.writeOutputText(
    `${outputBase}-role-labeled-transcript.md`,
    renderRoleLabeledTranscript(structured.roleAnalysis),
  );
  const evidenceBankPath = await storage.writeOutputJson(`${outputBase}-evidence-bank.json`, structured.evidenceBank || {});
  let pptPath = null;

  if (template.id === "weekly-report") {
    const pptDraft = buildWeeklyPptDraftFromStructured(structured);
    const pptDraftPath = await storage.writeOutputJson(`${outputBase}-ppt-draft.json`, pptDraft);
    pptPath = storage.getOutputPath(`${outputBase}.pptx`);
    await renderWeeklyReportPpt({ draftPath: pptDraftPath, outputPath: pptPath });
  } else if (template.id === "oi-news-report") {
    const pptDraft = buildOiNewsPptDraftFromStructured(structured);
    const pptDraftPath = await storage.writeOutputJson(`${outputBase}-ppt-draft.json`, pptDraft);
    pptPath = storage.getOutputPath(`${outputBase}.pptx`);
    await renderOiNewsReportPpt({ draftPath: pptDraftPath, outputPath: pptPath });
  }

  const review = buildCoverageReview({
    templateId: template.id,
    sourceText: buildSourcePacket({
      transcriptText: sourceJob.inputs.transcriptText || "",
      materialText,
      research: structured.research,
      roleAnalysis: structured.roleAnalysis,
    }),
    structured,
  });

  const jobRecord = {
    job: {
      id: outputBase,
      createdAt: structured.createdAt,
      title: structured.meeting.meta.title,
      company: structured.meeting.meta.company,
      templateId: template.id,
      templateName: template.name,
      transcriptPath: sourceJob.job.transcriptPath,
      roleTranscriptPath,
      evidenceBankPath,
      markdownPath,
      summary: structured.summary.oneSentence,
      modelMode: structured.processing.modelMode,
      llmProvider: structured.processing.llmProvider,
      sourceType: structured.processing.sourceType,
      asrProvider: structured.processing.asrProvider,
      researchProvider: structured.processing.researchProvider,
      pptPath,
      derivedFromJobId: sourceJob.job.id,
    },
    inputs: {
      meetingTitle: structured.meeting.meta.title,
      company: structured.meeting.meta.company,
      meetingType: structured.meeting.meta.meetingType,
      participants,
      templateId: template.id,
      asrProvider: sourceJob.inputs.asrProvider || sourceJob.job.asrProvider,
      llmProvider: llmProvider || sourceJob.job.llmProvider,
      researchProvider,
      notesText: derivedNotes,
      userInstructions,
      transcriptText: sourceJob.inputs.transcriptText || sourceStructured.meeting?.transcript || "",
      roleAnalysis: structured.roleAnalysis,
      evidenceBank: structured.evidenceBank,
      materials: [...(sourceJob.inputs.materials || []), ...additionalMaterials.materials],
    },
    review,
    structured,
    markdown: output,
    artifacts: {
      transcriptPath: sourceJob.artifacts?.transcriptPath || sourceJob.job.transcriptPath,
      roleTranscriptPath,
      evidenceBankPath,
      markdownPath,
      pptPath,
      obsidianPath: null,
    },
  };

  const structuredPath = await saveJob(jobRecord);
  jobRecord.job.structuredPath = structuredPath;
  await saveJob(jobRecord);
  return toClientJobRecord(jobRecord);
}

function renderEvidenceBankForPrompt(evidenceBank) {
  if (!evidenceBank || typeof evidenceBank !== "object") {
    return "";
  }

  return sanitizeOutputLines(evidenceBank.companyEvidence || [], 40)
    .map((line) => `- ${line}`)
    .join("\n");
}

function renderStructuredMemoFactsForWeekly(structured) {
  if (!structured || typeof structured !== "object") {
    return "";
  }

  const lines = [
    structured.summary?.oneSentence ? `Summary: ${structured.summary.oneSentence}` : "",
    ...(structured.sections || []).flatMap((section) =>
      (section.bullets || []).map((bullet) => `${section.title}: ${bullet}`),
    ),
    ...(structured.dataPoints || []).map((item) => `Evidence: ${item}`),
    ...(structured.fundraisingNotes || []).map((item) => `Financing: ${item}`),
    ...(structured.fundingTable || []).map((row) =>
      `Financing table: ${[row.round, row.raised, row.valuation, row.keyShareholders].filter(Boolean).join(" | ")}`,
    ),
    ...(structured.risks || []).map((item) => `Risk / watch: ${item}`),
    ...(structured.actionItems || []).map((item) => `Next step: ${item}`),
    ...(structured.quotes || []).map((item) => `Management POV: "${stripWrappingQuotes(item)}"`),
  ];

  return sanitizeOutputLines(lines, 80)
    .map((line) => `- ${line}`)
    .join("\n");
}

function buildWeeklyPptDraftFromStructured(structured) {
  const draft = normalizeWeeklyReportDraft(structured.weeklyReportDraft, "weekly-report");
  const participants = structured.meeting.meta.participants || [];
  const appendix = buildWeeklyAppendixDraft(structured);

  return {
    title: draft.title,
    meetingInformation: {
      date: "TBD",
      participants: participants.join(", ") || draft.meetingInformation.participants,
    },
    companyDescription: draft.companyDescription,
    category: draft.category,
    sourceContact: draft.sourceContact,
    opinion: draft.opinion,
    nextStep: draft.nextStep,
    sections: draft.sections,
    appendix,
  };
}

function buildOiNewsPptDraftFromStructured(structured) {
  const draft = structured.oiNewsDraft || buildOiNewsDraft(structured);

  return {
    title: structured.meeting.meta.company || "OI News Report",
    headline: draft.headline,
    dateline: draft.dateline,
    sourceLine: draft.sourceLine,
    sections: draft.sections,
    page: {
      widthCm: 19.05,
      heightCm: 27.517,
      portrait: true,
      fontName: "Arial Narrow",
      bodyFontSize: 10.5,
      contentFillRatio: 0.62,
    },
  };
}

function buildWeeklyAppendixDraft(structured) {
  const company = structured.meeting.meta.company || "Company";
  const insights = structured.materialInsights || createEmptyMaterialInsights();

  const roadmapLines = dedupeLines([
    ...insights.roadmap,
    ...structured.dataPoints.filter((item) => /roadmap|timeline|milestone|launch|pilot|sop|mass production/i.test(item)),
  ]).slice(0, 6);

  const unitEconomicsLines = dedupeLines([
    ...insights.unitEconomics,
    ...structured.sections.flatMap((section) =>
      (section.bullets || []).filter((item) => /cost|margin|payback|economics|price|revenue|utilization/i.test(item)),
    ),
  ]).slice(0, 6);

  const fundraisingLines = dedupeLines([
    ...insights.fundraising,
    ...structured.sections.flatMap((section) =>
      (section.bullets || []).filter((item) => /fundraising|series|pre-ipo|valuation|investor|shareholder|listing/i.test(item)),
    ),
    ...structured.dataPoints.filter((item) => /fundraising|series|pre-ipo|valuation|investor|shareholder|listing/i.test(item)),
  ]).slice(0, 5);

  const fallbackRoadmap = roadmapLines.length
    ? roadmapLines.map((line) => truncateForSlide(line, 150)).slice(0, 6)
    : [structured.summary.oneSentence || `${company} roadmap details were limited in the current packet.`];
  const fallbackUnitEconomics = unitEconomicsLines.length
    ? unitEconomicsLines.map((line) => truncateForSlide(line, 150)).slice(0, 6)
    : ["Unit-economics detail was limited in the current packet and should be validated in follow-up materials."];
  const fundingRows = deriveFundraisingRows(structured).slice(0, 4);

  return {
    title: `[Appendix] ${company}`,
    subtitle: "Material-derived highlights",
    roadmap: {
      title: "Product Roadmap",
      subtitle: roadmapLines.length
        ? "Pulled from uploaded materials and normalized into slide-ready milestones."
        : "No explicit roadmap page was detected, so this section reflects the clearest available launch signals.",
      lines: fallbackRoadmap,
    },
    economics: {
      title: "Unit Economics / Commercial Signals",
      subtitle: unitEconomicsLines.length
        ? "Commercial and operating signals surfaced from uploaded materials."
        : "Key commercial signals still look thin and may need manual follow-up.",
      lines: fallbackUnitEconomics,
    },
    funding: {
      title: "Funding Snapshot",
      subtitle: fundraisingLines[0] || "Financing context extracted from materials and structured notes.",
      rows: fundingRows,
    },
  };
}

function buildOiNewsDraft(structured) {
  const meta = structured.meeting.meta;
  const sections = structured.sections || [];
  const cleanedDataPoints = (structured.dataPoints || []).map(cleanTranscriptArtifact).filter(Boolean);
  const cleanedRisks = (structured.risks || []).map(cleanTranscriptArtifact).filter(Boolean);
  const cleanedActionItems = (structured.actionItems || []).map(cleanTranscriptArtifact).filter(Boolean);

  const keyNewsLines = dedupeLines([
    ...sections.flatMap((section) =>
      (section.bullets || []).filter((item) => /launch|announce|partner|deal|product|customer|fund|ipo|valuation|policy|expansion|rollout/i.test(item)),
    ),
    ...cleanedDataPoints,
  ])
    .map(cleanTranscriptArtifact)
    .map((line) => truncateForSlide(line, 130))
    .slice(0, 3);

  const relevanceLines = dedupeLines([
    ...sections.flatMap((section) =>
      (section.bullets || []).filter((item) => /strategy|market|implication|gtm|relevance|positioning|commercial|ecosystem/i.test(item)),
    ),
    ...toBulletLines(structured.summary.executiveSummary, 4),
  ])
    .map(cleanTranscriptArtifact)
    .map((line) => truncateForSlide(line, 130))
    .slice(0, 3);

  const watchLines = dedupeLines([
    ...cleanedRisks,
    ...cleanedActionItems,
  ])
    .map(cleanTranscriptArtifact)
    .map((line) => truncateForSlide(line, 120))
    .slice(0, 2);

  return normalizeOiNewsDraft({
    headline: cleanTranscriptArtifact(structured.summary.oneSentence || meta.title || `${meta.company} update`),
    dateline: `${structured.createdAt.slice(0, 10)} | ${meta.company || "Unknown company"}`,
    sourceLine: meta.title || structured.processing.sourceType || "Source packet",
    sections: [
      {
        label: "Key News",
        lines: keyNewsLines.length ? keyNewsLines : ["No major news item was isolated from the current packet."],
      },
      {
        label: "Strategic Relevance",
        lines: relevanceLines.length ? relevanceLines : ["Strategic relevance should be refined after source review."],
      },
      {
        label: "Watch Items",
        lines: watchLines.length ? watchLines : ["No explicit watch item was captured from the current packet."],
      },
    ].filter((section, index) => index < 2 || section.lines.some((line) => !/No explicit watch item/i.test(line))),
  });
}

async function renderWeeklyReportPpt({ draftPath, outputPath }) {
  const scriptPath = path.join(SCRIPTS_DIR, "render-weekly-report-ppt.ps1");
  await runPowerShellScript(scriptPath, [draftPath, outputPath]);
}

async function renderOiNewsReportPpt({ draftPath, outputPath }) {
  const scriptPath = path.join(SCRIPTS_DIR, "render-oi-news-report-ppt.ps1");
  await runPowerShellScript(scriptPath, [draftPath, outputPath]);
}

async function runPowerShellScript(scriptPath, args) {
  await new Promise((resolve, reject) => {
    const child = spawn("powershell", ["-ExecutionPolicy", "Bypass", "-File", scriptPath, ...args], {
      stdio: "pipe",
    });

    let stderr = "";
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      reject(createError(500, `PPT rendering failed: ${stderr || `exit code ${code}`}`));
    });
  });
}

async function runCaptureProcess(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: ["ignore", "pipe", "pipe"],
    });

    const stdout = [];
    const stderr = [];
    child.stdout.on("data", (chunk) => stdout.push(Buffer.from(chunk)));
    child.stderr.on("data", (chunk) => stderr.push(Buffer.from(chunk)));
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve(Buffer.concat(stdout).toString("utf8"));
        return;
      }

      reject(createError(500, `${command} failed: ${Buffer.concat(stderr).toString("utf8") || `exit code ${code}`}`));
    });
  });
}

function collectLines(sourceText, matchers) {
  const lines = sourceText
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean);

  return lines
    .filter((line) => matchers.some((matcher) => matcher.test(line)))
    .slice(0, 8);
}

function toWeeklySlideLine(line) {
  let cleaned = cleanGeneratedText(line)
    .replace(/^Existing Interview Memo\s*:\s*/i, "")
    .replace(/^Evidence Bank\s*:\s*/i, "")
    .replace(/^User instructions for derived output\s*:\s*/i, "")
    .trim();

  cleaned = cleaned
    .replace(/^(?:Headline|Key Updates|Next Steps|Core Updates|Risks?)\s*:\s*(?:Current Memo Facts|Company Evidence Lines)\s*:?\s*-?\s*/i, "")
    .replace(/(?:Current Interview Memo Content|Current Memo Facts|Company Evidence Lines|User Weekly Report Instructions|Evidence Bank|Source Notes)\s*:?\s*-?\s*/gi, "")
    .replace(/^(?:Headline|Key Updates|Next Steps|Core Updates)\s*:\s*(?:Summary|Technical \/ Product Discussion|Founder \/ Management Background|GTM \/ Customer Traction|Fundraising \/ Pre-IPO Discussion)\s*:\s*/i, "")
    .replace(/^(?:Headline|Key Updates|Next Steps|Core Updates|Risks?)\s*:\s*/i, "")
    .replace(/^(?:Technical \/ Product Discussion|Founder \/ Management Background|GTM \/ Customer Traction|Fundraising \/ Pre-IPO Discussion|Mixed Diligence Topics|Primary Discussion)\s*:\s*/i, "")
    .replace(/^Mini-head\s*:\s*/i, "")
    .replace(/^(?:Product \/ Tech|Product \/ Technology|Commercial|Evidence|Team)\s*:\s*/i, "")
    .replace(/^Management POV\s*:\s*["']?/i, "")
    .replace(/["']$/i, "")
    .replace(/^Summary\s*:\s*/i, "")
    .trim();

  if (!cleaned) {
    return "";
  }

  if (/do not copy labels|Create a weekly report|overall prompt|Evidence Bank|Source Notes|Role-labeled|companyEvidence|interviewerContext/i.test(cleaned)) {
    return "";
  }

  if (/^[A-Z][A-Za-z0-9 /&+-]{2,34}:/.test(cleaned)) {
    return cleaned;
  }

  if (looksLikeFundingLine(cleaned)) {
    return `Financing: ${cleaned}`;
  }
  if (/product|platform|tech|technology|model|data|sensor|architecture|roadmap|software|hardware|vision|ai inspection/i.test(cleaned)) {
    if (/data|sensor|model|architecture|foundation|vision|simulation|evaluation|emg/i.test(cleaned)) {
      return `Tech stack: ${cleaned}`;
    }
    if (/roadmap|launch|rollout|sku|payload|range|dof|hand|hardware|platform/i.test(cleaned)) {
      return `Platform: ${cleaned}`;
    }
    return `Product: ${cleaned}`;
  }
  if (/customer|pilot|deployment|commercial|gtm|go-to-market|revenue|pricing|market/i.test(cleaned)) {
    if (/revenue|pricing|payback|unit economics|margin|cost/i.test(cleaned)) {
      return `Economics: ${cleaned}`;
    }
    return `Traction: ${cleaned}`;
  }
  if (/founder|team|ceo|cto|management|leadership/i.test(cleaned)) {
    return `Team: ${cleaned}`;
  }

  return cleaned;
}

function padWeeklySectionLines(label, lines, weeklySource = []) {
  const rawLines = dedupeWeeklyLines(lines)
    .map((line) => truncateForSlide(line, 180))
    .filter(Boolean)
    .slice(0, 18);
  const nonVerifyLines = rawLines.filter((line) => !isWeeklyInputNoise(line));
  const sectionLines = nonVerifyLines.length >= 6 ? nonVerifyLines : rawLines;
  const minimumLines = 8;

  if (sectionLines.length >= minimumLines) {
    return buildHierarchicalWeeklyLines(label, sectionLines);
  }

  const fallbackLines = getWeeklySectionFallbackLines(label, weeklySource)
    .map((line) => toWeeklySlideLine(line))
    .filter(Boolean)
    .filter((line) => !sectionLines.map((item) => item.toLowerCase()).includes(line.toLowerCase()));

  for (const line of fallbackLines) {
    if (sectionLines.length >= minimumLines) {
      break;
    }
    sectionLines.push(truncateForSlide(line, 180));
  }

  while (sectionLines.length < minimumLines) {
    sectionLines.push(getWeeklyVerificationLine(label, sectionLines.length));
  }

  const finalLines = dedupeWeeklyLines(sectionLines);
  while (finalLines.length < minimumLines) {
    finalLines.push(getWeeklyVerificationLine(label, finalLines.length));
  }

  return buildHierarchicalWeeklyLines(label, finalLines);
}

function buildHierarchicalWeeklyLines(label, lines) {
  const source = dedupeWeeklyLines(lines)
    .map(stripWeeklyLinePrefix)
    .filter(Boolean);

  if (label === "Product / Technology") {
    return buildProductTechnologyHierarchy(source);
  }

  if (label === "Commercialization / Funding") {
    return buildCommercialFundingHierarchy(source);
  }

  return buildCompanyBackgroundHierarchy(source);
}

function stripWeeklyLinePrefix(line) {
  return cleanGeneratedText(line)
    .replace(/^#+\s*/, "")
    .replace(/^[•.\-\s]+/, "")
    .replace(/^[①②③④⑤⑥⑦⑧⑨]\s*/, "")
    .replace(/^Product Matrix\s*:\s*/i, "")
    .replace(/^(?:product\s*\/\s*technology|technology\s*&\s*product architecture|commercialization\s*\/\s*funding|commercial\s*&\s*financial signals|go-to-market\s*&\s*strategic positioning|data strategy\s*&\s*infrastructure|team\s*&\s*governance)\s*:\s*/i, "")
    .replace(/^(?:team|background|context|platform|product|tech stack|technology|financing|unit economics|commercialization|traction)\s*:\s*/i, "")
    .trim();
}

function isWeeklyVerificationNoise(line) {
  const cleaned = stripWeeklyLinePrefix(line);
  return !cleaned ||
    /^To verify\b/i.test(cleaned) ||
    /additional .* detail is needed/i.test(cleaned) ||
    /not stated in source|not fully captured|before external circulation/i.test(cleaned);
}

function isWeeklyInputNoise(line) {
  const cleaned = stripWeeklyLinePrefix(line);
  return isWeeklyVerificationNoise(line) || /^(Product|Technology|Commercialization|Funding|Company \/ Team)$/i.test(cleaned);
}

function takeWeeklyLines(lines, patterns, used, limit = 2) {
  const selected = [];
  for (const line of lines) {
    const key = line.toLowerCase();
    if (used.has(key)) {
      continue;
    }
    if (patterns.some((pattern) => pattern.test(line))) {
      used.add(key);
      selected.push(line);
    }
    if (selected.length >= limit) {
      break;
    }
  }
  return selected;
}

function weeklyFallback(lines, used, limit = 1) {
  return takeWeeklyLines(lines, [/.+/], used, limit);
}

function addWeeklyLine(out, line, marker = "  -") {
  const cleaned = truncateForSlide(stripWeeklyLinePrefix(line), marker === "  -" ? 150 : 165);
  if (cleaned) {
    out.push(`${marker} ${cleaned}`);
  }
}

function buildCompanyBackgroundHierarchy(lines) {
  const used = new Set();
  const team = takeWeeklyLines(lines, [/team|people|founder|ceo|cto|coo|cfo|leadership|researcher|tsinghua|horizon/i], used, 3);
  const thesis = takeWeeklyLines(lines, [/physical ai|industry|generalization|industrial|workstation|home|hotel|market|position|builds|company/i], used, 3);
  const governance = takeWeeklyLines(lines, [/red chip|governance|structure|founded|hq|headquarter|investor|board/i], used, 2);
  const residual = weeklyFallback(lines, used, 2);
  const out = ["• Core Team:"];

  addWeeklyLine(out, team[0] || residual[0] || getWeeklyVerificationLine("Company Background", 0));
  team.slice(1, 3).forEach((line) => addWeeklyLine(out, line));
  addWeeklyLine(out, thesis[0] || residual[1] || getWeeklyVerificationLine("Company Background", 1), "•");
  thesis.slice(1, 3).forEach((line) => addWeeklyLine(out, line));
  if (governance.length) {
    addWeeklyLine(out, governance[0], "•");
    governance.slice(1, 2).forEach((line) => addWeeklyLine(out, line));
  }

  return ensureWeeklyHierarchyDensity(dedupeWeeklyLines(out), "Company Background", 9).slice(0, 10);
}

function buildProductTechnologyHierarchy(lines) {
  const used = new Set();
  const product = takeWeeklyLines(lines, [/hand|hardware|platform|product|toolchain|poc|small-batch|mass production|roadmap|dof|tendon/i], used, 4);
  const data = takeWeeklyLines(lines, [/data|emg|pose|vision|force|touch|sensor|factory|collection|vocational|100,?000|full-modality/i], used, 4);
  const model = takeWeeklyLines(lines, [/model|foundation|strategy|world|vision-centric|force-centric|frequency|cross-attention|simulation|reasoning/i], used, 4);
  const control = takeWeeklyLines(lines, [/execution|control|closed-loop|trajectory|action-intent|reflexive|real-time|policy/i], used, 3);
  const residual = weeklyFallback(lines, used, 3);
  const productLead = product[0] || residual[0] || getWeeklyVerificationLine("Product / Technology", 0);
  const out = [`Product Matrix: ${truncateForSlide(stripWeeklyLinePrefix(productLead), 170)}`];

  addWeeklyLine(out, product[1] || productLead, "•");
  product.slice(2, 4).forEach((line) => addWeeklyLine(out, line));
  addWeeklyLine(out, model[0] || residual[1] || getWeeklyVerificationLine("Product / Technology", 1), "•");
  if (model[1]) {
    out.push(`Philosophy: ${truncateForSlide(stripWeeklyLinePrefix(model[1]), 170)}`);
  }
  addWeeklyLine(out, control[0] || model[2] || getWeeklyVerificationLine("Product / Technology", 2), "①");
  control.slice(1, 2).forEach((line) => addWeeklyLine(out, line));
  addWeeklyLine(out, model[2] || residual[2] || getWeeklyVerificationLine("Product / Technology", 3), "②");
  addWeeklyLine(out, data[0] || getWeeklyVerificationLine("Product / Technology", 4), "③");
  data.slice(1, 4).forEach((line) => addWeeklyLine(out, line));

  return ensureWeeklyHierarchyDensity(dedupeWeeklyLines(out), "Product / Technology", 12).slice(0, 14);
}

function buildCommercialFundingHierarchy(lines) {
  const used = new Set();
  const gtm = takeWeeklyLines(lines, [/hotel|customer|commercial|gtm|go-to-market|market|international|overseas|china|labor|traction|partner|pilot/i], used, 4);
  const business = takeWeeklyLines(lines, [/revenue|pricing|unit economics|data services|platform sales|developer|university|margin|asp|business model/i], used, 3);
  const funding = takeWeeklyLines(lines, [/financing|series|valuation|term sheet|tranche|raised|investor|shareholder|red chip|horizon|hillhouse|xiaomi|k3/i], used, 4);
  const risks = takeWeeklyLines(lines, [/risk|not disclosed|verify|pressure|tam|execution|pricing model/i], used, 2);
  const residual = weeklyFallback(lines, used, 2);
  const out = ["Commercialization"];

  addWeeklyLine(out, gtm[0] || residual[0] || getWeeklyVerificationLine("Commercialization / Funding", 0), "•");
  gtm.slice(1, 3).forEach((line) => addWeeklyLine(out, line));
  addWeeklyLine(out, business[0] || residual[1] || getWeeklyVerificationLine("Commercialization / Funding", 1), "•");
  business.slice(1, 2).forEach((line) => addWeeklyLine(out, line));
  out.push("Funding:");
  addWeeklyLine(out, funding[0] || getWeeklyVerificationLine("Commercialization / Funding", 2), "•");
  funding.slice(1, 4).forEach((line) => addWeeklyLine(out, line));
  risks.slice(0, 2).forEach((line) => addWeeklyLine(out, line));

  return ensureWeeklyHierarchyDensity(dedupeWeeklyLines(out), "Commercialization / Funding", 11).slice(0, 13);
}

function ensureWeeklyHierarchyDensity(lines, label, targetCount) {
  return dedupeWeeklyLines(lines || []).filter((line) => !isWeeklyVerificationNoise(line));
}

function dedupeWeeklyLines(lines) {
  const seen = new Set();
  return (lines || []).filter((line) => {
    const cleaned = String(line || "").trim();
    if (!cleaned) {
      return false;
    }
    const key = cleaned
      .toLowerCase()
      .replace(/^[a-z][a-z0-9 /&+-]{2,50}:\s*/i, "")
      .replace(/^(product\s*\/\s*technology|commercialization\s*\/\s*funding|technology\s*&\s*product architecture|commercial\s*&\s*financial signals|go-to-market\s*&\s*strategic positioning|data strategy\s*&\s*infrastructure)\s*:\s*/i, "")
      .replace(/[^\p{L}\p{N}]+/gu, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function compressWeeklyLineLabels(sectionLabel, lines) {
  const maxLabeledLines = sectionLabel === "Product / Technology" ? 3 : 2;
  const seenPrefixes = new Set();
  let labeledCount = 0;

  return (lines || []).map((line) => {
    const cleaned = String(line || "").trim();
    const match = cleaned.match(/^([A-Z][A-Za-z0-9 /&+-]{2,34}):\s*(.+)$/);
    if (!match) {
      return cleaned;
    }

    const prefix = normalizeWeeklyLinePrefix(match[1], sectionLabel);
    const body = match[2].trim();
    if (!prefix || !body) {
      return body || cleaned;
    }

    const prefixKey = prefix.toLowerCase();
    if (seenPrefixes.has(prefixKey) || labeledCount >= maxLabeledLines) {
      return body;
    }

    seenPrefixes.add(prefixKey);
    labeledCount += 1;
    return `${prefix}: ${body}`;
  });
}

function normalizeWeeklyLinePrefix(prefix, sectionLabel) {
  const text = String(prefix || "").trim().toLowerCase();

  if (sectionLabel === "Company Background") {
    if (/team|founder|ceo|cto|management|leadership/.test(text)) {
      return "Team";
    }
    if (/market|category|background|company|context/.test(text)) {
      return "Background";
    }
    return "";
  }

  if (sectionLabel === "Product / Technology") {
    if (/data|model|architecture|tech|stack|sensor|simulation|evaluation|foundation/.test(text)) {
      return "Tech stack";
    }
    if (/platform|product|roadmap|rollout|hardware|sku/.test(text)) {
      return "Platform";
    }
    return "";
  }

  if (/fund|financing|ipo|valuation|round/.test(text)) {
    return "Financing";
  }
  if (/economics|revenue|pricing|payback|margin|cost/.test(text)) {
    return "Unit economics";
  }
  if (/traction|customer|commercial|gtm|partner|deployment|market/.test(text)) {
    return "Commercialization";
  }
  return "";
}

function getWeeklySectionFallbackLines(label, weeklySource = []) {
  const patterns =
    label === "Company Background"
      ? [/builds|company|founded|hq|headquarter|team|founder|management|category|market/i]
      : label === "Product / Technology"
        ? [/product|platform|tech|technology|model|data|sensor|software|hardware|roadmap|accuracy|architecture/i]
        : [/customer(?!-specific)|pilot|commercial|gtm|go-to-market|revenue|pricing|fund|financing|series|ipo|valuation|singapore|hmg|hyundai|partner/i];

  return weeklySource.filter((line) => patterns.some((pattern) => pattern.test(line))).slice(0, 6);
}

function getWeeklyVerificationLine(label, index) {
  const missing = {
    "Company Background": [
      "To verify: founding year, HQ city, and leadership background were not fully captured in the current packet.",
      "To verify: confirm the company one-line positioning and source contact before external circulation.",
      "To verify: validate team and shareholder context with updated company materials.",
      "To verify: check whether public materials add stronger market-positioning evidence.",
    ],
    "Product / Technology": [
      "To verify: product architecture, performance metrics, and technical differentiation need stronger source evidence.",
      "To verify: confirm product roadmap, deployment readiness, and validation milestones.",
      "To verify: collect proof points on model/data stack, hardware requirements, and customer deployment constraints.",
      "To verify: add product screenshots or technical appendix if available.",
    ],
    "Commercialization / Funding": [
      "To verify: customer traction, revenue model, and GTM path need fuller evidence.",
      "To verify: confirm fundraising stage, planned round, valuation context, and key investors.",
      "To verify: assess HMG/Singapore collaboration relevance after follow-up material review.",
      "To verify: validate near-term next step with the deal owner before sharing.",
    ],
  };

  return missing[label]?.[index] || `To verify: additional ${label.toLowerCase()} detail is needed.`;
}

function dedupeLines(lines) {
  const seen = new Set();
  return lines.filter((line) => {
    const normalized = String(line || "").trim();
    if (!normalized) {
      return false;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function normalizeExtractedPayload(extracted, templateId) {
  const normalized = extracted && typeof extracted === "object" ? extracted : {};

  const sections = normalizeSections(normalized.sections);
  const memoCategories = normalizeMemoCategories(normalized.memoCategories);
  const actionItems = sanitizeOutputLines(normalized.actionItems, 12);
  const risks = sanitizeOutputLines(normalized.risks, 12);
  const quotes = sanitizeOutputLines(normalized.quotes, 8);
  const dataPoints = sanitizeOutputLines(normalized.dataPoints, 20);
  const uncategorized = sanitizeOutputLines(normalized.uncategorized, 20);
  const fundraisingNotes = sanitizeOutputLines(normalized.fundraisingNotes, 12);
  const fundingTable = normalizeFundingTable(normalized.fundingTable);

  const oneSentence = cleanGeneratedText(normalized.oneSentence);
  const executiveSummary = cleanGeneratedParagraph(normalized.executiveSummary);

  return {
    ...normalized,
    oneSentence,
    executiveSummary,
    sections,
    memoCategories,
    actionItems,
    risks,
    quotes,
    dataPoints,
    uncategorized,
    fundraisingNotes,
    fundingTable,
    weeklySlide: normalizeWeeklySlide(normalized.weeklySlide),
    weeklyReportDraft: hasWeeklyReportDraftContent(normalized.weeklyReportDraft)
      ? normalizeWeeklyReportDraft(normalized.weeklyReportDraft, templateId)
      : null,
    oiNewsDraft: normalizeOiNewsDraft(normalized.oiNewsDraft),
  };
}

function normalizeMemoCategories(categories) {
  if (!Array.isArray(categories)) {
    return [];
  }

  return categories
    .map((category) => {
      const title = cleanGeneratedText(category?.title);
      const groups = normalizeMemoCategoryGroups(category?.groups);
      const lines = sanitizeOutputLines(category?.lines, 80);
      const quotes = sanitizeOutputLines(category?.quotes, 12);
      if (!title || (!groups.length && !lines.length && !quotes.length)) {
        return null;
      }
      return { title, groups, lines, quotes };
    })
    .filter(Boolean)
    .slice(0, 12);
}

function normalizeMemoCategoryGroups(groups) {
  if (!Array.isArray(groups)) {
    return [];
  }

  return groups
    .map((group) => {
      const title = cleanGeneratedText(group?.title);
      const lines = sanitizeOutputLines(group?.lines, 80);
      const quotes = sanitizeOutputLines(group?.quotes, 12);
      if (!title || (!lines.length && !quotes.length)) {
        return null;
      }
      return { title, lines, quotes };
    })
    .filter(Boolean)
    .slice(0, 20);
}

function normalizeSections(sections) {
  if (!Array.isArray(sections)) {
    return [];
  }

  return sections
    .map((section) => {
      const title = cleanGeneratedText(section?.title);
      const bullets = sanitizeOutputLines(section?.bullets, 24);
      if (!title || bullets.length === 0) {
        return null;
      }
      return { title, bullets };
    })
    .filter(Boolean);
}

function sanitizeOutputLines(lines, limit = 20) {
  if (!Array.isArray(lines)) {
    return [];
  }

  return dedupeLines(
    lines
      .map(cleanGeneratedText)
      .filter(Boolean)
      .filter((line) => !looksLikeMostlyTranscriptGarbage(line)),
  ).slice(0, limit);
}

function cleanGeneratedParagraph(value) {
  const cleaned = cleanGeneratedText(value);
  return cleaned.length > 1200 ? cleaned.slice(0, 1200).trim() : cleaned;
}

function cleanGeneratedText(value) {
  return cleanTranscriptArtifact(String(value || ""))
    .replace(/^Transcript\s*:\s*/i, "")
    .replace(/^Supporting materials\s*:\s*/i, "")
    .replace(/^Material\s*:\s*/i, "")
    .replace(/^\s*(?:[-*•]|\d+[.)])\s+/, "")
    .replace(/\s+\|\s+/g, " | ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function looksLikeMostlyTranscriptGarbage(value) {
  const text = String(value || "").trim();
  if (!text) {
    return true;
  }

  if (/^(speaker|Transcript|Supporting materials|Material)\b/i.test(text)) {
    return true;
  }

  if (/^\d{1,2}:\d{2}(?::\d{2})?$/.test(text)) {
    return true;
  }

  return false;
}

function normalizeFundingTable(rows) {
  if (!Array.isArray(rows)) {
    return [];
  }

  return rows
    .map((row) => {
      if (!row || typeof row !== "object") {
        return null;
      }

      const normalizedRow = {
        round: cleanGeneratedText(row.round),
        raised: cleanGeneratedText(row.raised),
        valuation: cleanGeneratedText(row.valuation),
        keyShareholders: cleanGeneratedText(row.keyShareholders),
      };

      const values = Object.values(normalizedRow);
      const hasUsefulContent = values.some(Boolean);
      const hasBadCell = values.some((cell) => isUnsafeFundingCell(cell));

      if (!hasUsefulContent || hasBadCell) {
        return null;
      }

      return normalizedRow;
    })
    .filter(Boolean)
    .slice(0, 6);
}

function isUnsafeFundingCell(value) {
  const text = String(value || "").trim();
  if (!text) {
    return false;
  }

  if (text.length > 120) {
    return true;
  }

  if (/(Speaker\s*\d+|璇磋瘽浜?|\[\d{1,2}:\d{2}(?::\d{2})?\])/i.test(text)) {
    return true;
  }

  if (/[。！？.!?].+[。！？.!?].+[。！？.!?]/.test(text)) {
    return true;
  }

  return false;
}

function normalizeWeeklySlide(weeklySlide) {
  const normalized = weeklySlide && typeof weeklySlide === "object" ? weeklySlide : {};
  return {
    headline: cleanGeneratedText(normalized.headline),
    updates: sanitizeOutputLines(normalized.updates, 8),
    risks: sanitizeOutputLines(normalized.risks, 6),
    nextSteps: sanitizeOutputLines(normalized.nextSteps, 6),
  };
}

function normalizeWeeklyReportDraft(weeklyReportDraft, templateId) {
  const normalized = weeklyReportDraft && typeof weeklyReportDraft === "object" ? weeklyReportDraft : {};
  const sections = Array.isArray(normalized.sections) ? normalized.sections : [];
  const safeSections = sections
    .map((section) => ({
      label: normalizeWeeklySectionLabel(section?.label),
      lines: sanitizeOutputLines(section?.lines, 8)
        .map(cleanWeeklyReportDraftLine)
        .filter(Boolean)
        .map((line) => truncateForSlide(line, 180)),
    }))
    .filter((section) => section.lines.length > 0)
    .slice(0, 3);
  const canonicalSections = normalizeWeeklySections(safeSections);

  return templateId === "weekly-report" || normalized.weeklyReportDraft || sections.length
    ? {
        title: truncateForSlide(cleanGeneratedText(normalized.title), 80),
        meetingInformation: {
          date: truncateForSlide(cleanGeneratedText(normalized.meetingInformation?.date), 40),
          participants: truncateForSlide(cleanGeneratedText(normalized.meetingInformation?.participants), 100),
        },
        companyDescription: {
          line1: truncateForSlide(cleanGeneratedText(normalized.companyDescription?.line1), 130),
          foundedYear: truncateForSlide(cleanGeneratedText(normalized.companyDescription?.foundedYear), 12),
          hqCity: truncateForSlide(cleanGeneratedText(normalized.companyDescription?.hqCity), 32),
        },
        category: truncateForSlide(cleanGeneratedText(normalized.category), 50),
        sourceContact: truncateForSlide(cleanGeneratedText(normalized.sourceContact), 60),
        opinion: truncateForSlide(cleanGeneratedText(normalized.opinion), 130),
        nextStep: truncateForSlide(cleanGeneratedText(normalized.nextStep), 130),
        sections: canonicalSections,
      }
    : null;
}

function normalizeWeeklySections(sections) {
  const canonical = ["Company Background", "Product / Technology", "Commercialization / Funding"].map((label) => ({
    label,
    lines: [],
  }));

  (sections || []).forEach((section) => {
    const defaultLabel = normalizeWeeklySectionLabel(section.label);
    (section.lines || []).forEach((line) => {
      const label = inferWeeklySectionForLine(line, defaultLabel);
      const target = canonical.find((item) => item.label === label) || canonical[2];
      target.lines.push(line);
    });
  });

  return canonical.map((section) => ({
    label: section.label,
    lines: padWeeklySectionLines(section.label, dedupeLines(section.lines), section.lines),
  }));
}

function enrichWeeklyReportDraftWithStructuredEvidence(weeklyReportDraft, structured) {
  const draft = normalizeWeeklyReportDraft(weeklyReportDraft, "weekly-report");
  const evidenceLines = collectWeeklyEvidenceLinesFromStructured(structured);
  const sections = draft.sections.map((section) => {
    const matchingEvidence = evidenceLines
      .filter((line) => inferWeeklySectionForLine(line, section.label) === section.label)
      .slice(0, section.label === "Product / Technology" ? 14 : 10);
    const generalEvidence = section.label === "Company Background"
      ? [
          structured.summary?.oneSentence,
          ...(structured.quotes || []).filter((quote) => /physical ai|industry|generalization|home|hotel|industrial/i.test(quote)),
        ].filter(Boolean)
      : [];
    return {
      ...section,
      lines: padWeeklySectionLines(section.label, dedupeWeeklyLines([...section.lines, ...generalEvidence, ...matchingEvidence]), matchingEvidence),
    };
  });

  return {
    ...draft,
    sections,
  };
}

function collectWeeklyEvidenceLinesFromStructured(structured) {
  const sections = Array.isArray(structured.sections) ? structured.sections : [];
  const materialInsights = structured.materialInsights || createEmptyMaterialInsights();
  const fundingRows = Array.isArray(structured.fundingTable) ? structured.fundingTable : [];
  const fundingLines = fundingRows.map((row) =>
    [row.round, row.raised, row.valuation, row.keyShareholders].filter(Boolean).join(" | "),
  );

  return dedupeWeeklyLines([
    ...(sections || []).flatMap((section) =>
      (section.bullets || []).map((bullet) => `${section.title}: ${bullet}`),
    ),
    ...(structured.dataPoints || []),
    ...(structured.fundraisingNotes || []),
    ...(structured.quotes || []),
    ...fundingLines,
    ...(materialInsights.roadmap || []),
    ...(materialInsights.fundraising || []),
    ...(materialInsights.unitEconomics || []),
  ])
    .map(cleanGeneratedText)
    .filter(Boolean)
    .filter((line) => !/Evidence Bank|Source Notes|Role-labeled|companyEvidence|interviewerContext|User Weekly Report Instructions/i.test(line));
}

function inferWeeklySectionForLine(line, defaultLabel) {
  const text = cleanGeneratedText(line)
    .replace(/^(?:product\s*\/\s*technology|technology\s*&\s*product architecture|commercialization\s*\/\s*funding|commercial\s*&\s*financial signals|go-to-market\s*&\s*strategic positioning|data strategy\s*&\s*infrastructure|team\s*&\s*governance)\s*:\s*/i, "")
    .replace(/^(?:team|background|context|platform|product|tech stack|financing|unit economics|commercialization|traction)\s*:\s*/i, "")
    .toLowerCase();

  if (/\b(team|founder|co-founder|ceo|cto|cfo|coo|management|leadership|people|headcount|core member|researcher|founded|hq|headquarter|university|tsinghua)\b/.test(text)) {
    return "Company Background";
  }

  if (
    /fund|financ|valuation|investor|shareholder|ipo|listing|round|term sheet|tranche|revenue|pricing|unit economics|margin|payback|customer|hotel|commercial|commercialization|gtm|go-to-market|market|overseas|international|singapore|china|labor substitution|traction|data services|platform offering|contract|loi/.test(text)
  ) {
    return "Commercialization / Funding";
  }

  if (
    /product|platform|technology|tech stack|architecture|model|world model|foundation model|strategy model|deep thinking|data|emg|vision|pose|force|touch|sensor|simulation|cross-attention|dof|dexterous|hand|tendon|hardware|software|toolchain|roadmap|rollout|poc|prototype|mass production|hierarchical|action-intent|trajectory|execution layer|control loop|closed-loop/.test(text)
  ) {
    return "Product / Technology";
  }

  if (/red-chip|red chip|structure|horizon robotics/.test(text)) {
    return "Company Background";
  }

  return defaultLabel;
}

function normalizeWeeklySectionLabel(label) {
  const text = cleanGeneratedText(label).toLowerCase();
  if (/background|team|founder|company|profile|leadership/.test(text)) {
    return "Company Background";
  }
  if (/product|tech|technology|platform|model|data|roadmap|architecture/.test(text)) {
    return "Product / Technology";
  }
  return "Commercialization / Funding";
}

function hasWeeklyReportDraftContent(weeklyReportDraft) {
  if (!weeklyReportDraft || typeof weeklyReportDraft !== "object") {
    return false;
  }

  return Array.isArray(weeklyReportDraft.sections) &&
    weeklyReportDraft.sections.some((section) => Array.isArray(section?.lines) && section.lines.some((line) => String(line || "").trim()));
}

function cleanWeeklyReportDraftLine(line) {
  const cleaned = cleanGeneratedText(line)
    .replace(/^#+\s*/, "")
    .replace(/^(?:Current Interview Memo Content|Current Memo Facts|Company Evidence Lines|User Weekly Report Instructions|Existing Interview Memo|Evidence Bank)\s*:?\s*/i, "")
    .replace(/^(?:oneSentence|executiveSummary|sections|bullets|companyEvidence|interviewerContext|excludedContext|analystPov)\s*:?\s*/i, "")
    .trim();

  if (
    !cleaned ||
    /^(?:Source Notes|User-provided materials|Public-source enrichment|Inference \/ synthesis|Role-labeled Transcript|Evidence Bank JSON|Derived weekly-report source packet)$/i.test(cleaned) ||
    /do not copy labels|Create a weekly report|overall prompt/i.test(cleaned)
  ) {
    return "";
  }

  return toWeeklySlideLine(cleaned);
}

function normalizeOiNewsDraft(oiNewsDraft) {
  const normalized = oiNewsDraft && typeof oiNewsDraft === "object" ? oiNewsDraft : {};
  const sections = Array.isArray(normalized.sections) ? normalized.sections : [];

  return {
    headline: truncateForSlide(cleanGeneratedText(normalized.headline), 120),
    dateline: truncateForSlide(cleanGeneratedText(normalized.dateline), 80),
    sourceLine: truncateForSlide(cleanGeneratedText(normalized.sourceLine), 80),
    sections: sections
      .map((section) => ({
        label: truncateForSlide(cleanGeneratedText(section?.label), 32) || "Section",
        lines: sanitizeOutputLines(section?.lines, 3).map((line) => truncateForSlide(line, 130)),
      }))
      .filter((section) => section.lines.length > 0)
      .slice(0, 3),
  };
}

function truncateForSlide(value, maxLength = 120) {
  const text = String(value || "").trim();
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, maxLength - 1).trim()}…`;
}

function inferCategory(sourceText, meetingType) {
  if (/autonomous|robot|robovan|logistics/i.test(sourceText)) {
    return "Autonomous Logistics";
  }
  if (/battery|cathode|anode/i.test(sourceText)) {
    return "Battery";
  }
  return meetingType || "General";
}

function firstRegexMatch(text, regex) {
  const match = text.match(regex);
  return match?.[0] || "";
}

function jsonResponseFormatForProvider(providerId) {
  if (providerId === "deepseek") {
    return {};
  }

  return {
    response_format: { type: "json_object" },
  };
}

async function tryLLMExtraction({
  title,
  company,
  meetingType,
  participants,
  transcriptText,
  template,
  llmProvider,
  userInstructions,
  terminologyHints,
}) {
  const config = getResolvedLlmConfig(llmProvider);

  if (!config.apiKey || config.providerId === "none") {
    return null;
  }

  const prompt = buildExtractionPrompt({
    title,
    company,
    meetingType,
    participants,
    transcriptText: compactTextForLLM(transcriptText, getLlmInputLimitForTemplate(template.id)),
    template,
    userInstructions,
    terminologyHints,
  });

  const response = await fetchWithTimeout(`${config.baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: config.model,
      temperature: 0.2,
      ...jsonResponseFormatForProvider(config.providerId),
      messages: [
        {
          role: "system",
          content:
            "You are an investment research meeting processor. Return only valid JSON and keep grounded in the transcript.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
    }),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw createError(502, `LLM request failed: ${detail}`);
  }

  const payload = await response.json();
  const content = payload.choices?.[0]?.message?.content;

  if (!content) {
    throw createError(502, "LLM returned no content.");
  }

  return {
    providerId: config.providerId,
    structured: parseJsonContent(content),
  };
}

function parseJsonContent(content) {
  const raw = String(content || "").trim();
  if (!raw) {
    throw new Error("LLM returned an empty response.");
  }

  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = (fenced?.[1] || raw).trim();
  try {
    return JSON.parse(candidate);
  } catch (_error) {
    const start = candidate.indexOf("{");
    const end = candidate.lastIndexOf("}");
    if (start >= 0 && end > start) {
      return JSON.parse(candidate.slice(start, end + 1));
    }
    throw _error;
  }
}

function buildExtractionPrompt({ title, company, meetingType, participants, transcriptText, template, userInstructions, terminologyHints = [] }) {
  const templateSpecificRules =
    template.id === "interview-free-style"
      ? [
          "- For Free Style Memo, detect the actual focus and structure of the interview before drafting sections.",
          "- First classify the interview focus into the closest shape: tech deep dive, founder or management background, fundraising or pre-IPO, GTM or customer traction, or mixed diligence.",
          "- Do not force company-wide buckets like Team, Financials, or GTM unless the transcript really spends meaningful time on them.",
          "- Prefer 3 to 6 broad, content-led sections that follow the conversation's natural emphasis.",
          "- If 70% to 80% of the interview is technical, let the memo stay mostly technical instead of repeating the same facts across generic buckets.",
          "- Omit unsupported sections entirely instead of filling them with recycled content.",
          "- If a speaker is mainly describing personal experience or career path, let that become a founder or management background block instead of forcing a team section.",
          "- If the interview mostly discusses funding or listing, make financing or capital-markets discussion the main body rather than a side note.",
        ]
      : template.id === "interview-knowledge-base"
        ? [
            "- For Interview Memo, keep the structure compact and avoid repeating the same point across one-line take, summary, takeaways, and later sections.",
            "- Preserve interview completeness. Do not drop detailed technical, product, commercial, or management discussion just because it is too granular.",
            "- The memo is a lightly cleaned source-preserving record, not a high-level summary. Preserve material details and classify them; do not omit granular facts.",
            "- Use `memoCategories` as the main body. Create category titles from the actual discussion, e.g. `世界模型与人类模态数据`, `数据采集范式`, `规模化数据生产`, `硬件与评测闭环`, `客户与商业化`, `融资与治理`.",
            "- `memoCategories` should have enough density to feel like a detailed cleaned note. For a long transcript, each important topic should usually contain multiple lines, not one compressed bullet.",
            "- For transcripts above ~10k Chinese characters, produce a source-preserving memo with at least 60 detailed `memoCategories` lines if the source supports it.",
            "- Preserve concrete sub-discussions such as technology route debates, alternative approaches rejected, customer validation examples, pricing logic, hardware specs, data operations, evaluation setup, team background, compliance constraints, and fundraising context.",
            "- Use broad company buckets only if they naturally fit. Do not force fixed Team / Business / Product / Commercial sections if the transcript focus is narrower or more technical.",
            "- Split mixed topics into separate category groups. Do not put revenue, GTM, and technical architecture under the same generic section if they are separate facts.",
            "- If the interview spends time on ego-centric data, data collection devices, sensor rigs, data infrastructure, labeling, simulation, or evaluation tooling, create a dedicated data-related section under Product / Technology.",
            "- Put go-to-market, customer, channel, and deployment-market discussion under Business / Strategy, not Product / Technology.",
            "- Put revenue model, pricing, margin, unit economics, fundraising, valuation, and IPO discussion under Commercial / Financial Signals, not Product / Technology.",
            "- Do not create a separate key-evidence section. Put important metrics, dates, and proof points directly inside the most relevant module and make them stand out.",
            "- Identify important interviewee POV, statements, and comments as quotes. Keep them in quotes and attach them to the most relevant section. Include 3 to 8 quotes when the transcript has meaningful management POV.",
            "- uncategorized should be a last resort. Try to classify points into Team, Business, Product / Technology, Commercial / Financial Signals, or Open Questions first.",
            "- If funding is discussed, keep the funding table clean and short, then keep any messy residual detail in fundraisingNotes only.",
          ]
        : template.id === "weekly-report"
          ? [
              "- For Weekly Report, generate a slide-ready `weeklyReportDraft` as the primary output. Do not rely on generic sections only.",
              "- Treat the PowerPoint file as a fixed empty visual shell only. Never copy or preserve any company names, examples, placeholder text, logos, dates, or business facts from a prior template.",
              "- Use only the current source packet below: current memo, transcript, evidence bank, supporting materials, public research, and explicit user instructions.",
              "- If a fact appears in a prior template but is not present in the current source packet, exclude it.",
              "- The slide should read like an investor weekly report, not a transcript summary.",
              "- weeklyReportDraft.sections must always contain exactly three top-level sections in this order: Company Background, Product / Technology, Commercialization / Funding.",
              "- The three top-level sections are fixed, but content weighting is flexible: if the memo is mostly technical, write more specific Product / Technology proof lines; if it is mostly fundraising or GTM, make Commercialization / Funding more detailed.",
              "- Each of the three sections should contain 6 to 8 dense evidence lines so the core table is visually full. Do not leave any row empty.",
              "- If one of the three major areas is thin in the source, include a `To verify:` line that names the missing fact instead of inventing evidence.",
              "- Use compact investor hierarchy inside each core cell, not visible markdown headers. Do not output `# Product` or `# Technology`.",
              "- Product / Technology should normally start with `Product Matrix:` and then use a few strong bullets such as `• Dexterous Hand`, `• Foundation Model`, `①/②/③ named technology pillars`, and `• Data & training toolchain`.",
              "- Commercialization / Funding should use plain text anchors like `Commercialization`, `Vertical Focus:`, `Funding:` and compact proof bullets. Keep GTM, revenue model, customers, and financing distinct but in one dense cell.",
              "- Use compact evidence fragments, not long paragraphs. Use semantic prefixes sparingly, only when they clarify structure, e.g. `Team:`, `Platform:`, `Rollout Timeline:`, `Unit economics:`, `Financing:`.",
              "- Do not output the literal label `Mini-head:`. Do not create a mini-title for every bullet.",
              "- Avoid repeated subheads inside the same section. If several lines belong under the same theme, write one labeled lead line followed by plain proof bullets.",
              "- Highlight hard proof in wording: customer names, dates, deployment metrics, product specs, funding amounts, valuation, IPO timing, revenue, unit economics, partner names.",
              "- Company description line1 must be one sentence describing what the company does. foundedYear and hqCity must be source-backed when possible; use `TBD` if not in source.",
              "- opinion should state HMG impact or investment relevance based only on current evidence.",
              "- nextStep should be a concrete follow-up for the investor workflow.",
              "- Do not include transcript speaker labels, memo metadata, markdown headings, source-note labels, or Evidence Bank JSON field names in slide lines.",
              "- Do not mention `Role-labeled transcript`, `Source Notes`, `uncategorized`, `review queue`, or other internal workflow labels on the slide.",
            ]
          : [];

  return [
    "Process the following meeting packet into a structured JSON object for investment-research note organization.",
    `Meeting title: ${title}`,
    `Company: ${company}`,
    `Meeting type: ${meetingType}`,
    `Participants: ${participants.join(", ") || "Unknown"}`,
    `Output template: ${template.name}`,
    `User instructions: ${userInstructions || "None provided"}`,
    `Terminology hints from Obsidian: ${terminologyHints.length ? terminologyHints.join(", ") : "None"}`,
    "Terminology hints are vocabulary aids only. Use them to normalize terms already present in the source; do not introduce new facts, technologies, frameworks, products, or claims from the hints.",
    "JSON shape:",
    JSON.stringify(
      {
        oneSentence: "short top-line conclusion grounded in source material",
        executiveSummary: "short paragraph or dense bullets",
        sections: [
          {
            title: "broad section title",
            bullets: ["fact-based bullet preserving source detail", "fact-based bullet preserving source detail"],
          },
        ],
        actionItems: ["follow-up item"],
        risks: ["risk item"],
        quotes: ["short quote or management claim grounded in source material"],
        dataPoints: ["important metric, date, or evidence point"],
        fundingTable: [
          {
            round: "financing stage or round label",
            raised: "amount raised if clearly stated",
            valuation: "valuation if clearly stated",
            keyShareholders: "named investors or shareholders if clearly stated",
          },
        ],
        oiNewsDraft: {
          headline: "short headline for a one-slide portrait news report",
          dateline: "short date and company line",
          sourceLine: "short source label",
          sections: [
            {
              label: "Key News",
              lines: ["very short slide-safe bullet"],
            },
            {
              label: "Strategic Relevance",
              lines: ["very short slide-safe bullet"],
            },
            {
              label: "Watch Items",
              lines: ["very short slide-safe bullet"],
            },
          ],
        },
        uncategorized: ["important item that does not fit the main sections"],
        fundraisingNotes: ["fundraising detail that should not be lost"],
        weeklySlide: {
          headline: "weekly headline",
          updates: ["update"],
          risks: ["risk"],
          nextSteps: ["next step"],
        },
        memoCategories: [
          {
            title: "中文讨论主类目，必须基于实际访谈重点命名",
            groups: [
              {
                title: "更具体的二级主题，避免重复和泛化",
                lines: [
                  "保留原始细节的中文要点；不要为了漂亮而压缩掉数字、限定条件、例子、判断依据",
                  "如果一段话有多个独立事实，拆成多条 lines",
                ],
                quotes: ["被访谈者关键原话，保留中文原意"],
              },
            ],
            lines: ["没有二级主题时才使用的直接要点"],
            quotes: ["属于整个类目的关键原话"],
          },
        ],
        weeklyReportDraft: {
          title: "[Startup] Company",
          meetingInformation: {
            date: "YYYY-MM-DD or TBD",
            participants: "meeting participants or TBD",
          },
          companyDescription: {
            line1: "one-sentence description of what the company does",
            foundedYear: "founded year or TBD",
            hqCity: "HQ city only or TBD",
          },
          category: "company category",
          sourceContact: "source contact or TBD",
          opinion: "HMG impact or investor relevance in one compact sentence",
          nextStep: "concrete follow-up action",
          sections: [
            {
              label: "Company Background",
              lines: [
                "Team: compact source-grounded proof line",
                "To verify: missing background item if source is thin",
              ],
            },
            {
              label: "Product / Technology",
              lines: [
                "Product Matrix: product family + foundation model + data/toolchain",
                "• Product family or platform",
                "  - source-grounded proof line",
                "• Foundation Model: source-grounded model roadmap or architecture",
                "① Named technical pillar: source-grounded proof line",
                "② Named technical pillar: source-grounded proof line",
                "③ Named technical pillar: source-grounded proof line",
                "• Data & training toolchain",
              ],
            },
            {
              label: "Commercialization / Funding",
              lines: [
                "Commercialization",
                "Vertical Focus: GTM or customer wedge",
                "  - source-grounded proof line",
                "Funding:",
                "• Financing: compact source-grounded proof line",
              ],
            },
          ],
        },
      },
      null,
      2,
    ),
    "Rules:",
    "- For Interview Memo outputs, use Chinese as the primary working language unless the user explicitly asks for English. Keep English technical terms and company names as-is.",
    "- For Interview Memo outputs, populate `memoCategories` as the primary memo body. The category titles must follow the actual discussion focus, not a fixed template.",
    "- `memoCategories` should preserve details: use many precise lines rather than a few broad summaries. Keep numbers, examples, customer names, technical route details, constraints, and sequence logic.",
    "- For a long interview transcript, target a detailed memo body: usually 8 to 14 memo categories/groups total and 60 to 120 detailed lines across `memoCategories`. Do not collapse a long discussion into fewer than 40 detail lines unless the source is genuinely short.",
    "- Keep `oneSentence` and `executiveSummary` short. Do not hide the important details only in the summary; the detailed facts must live inside `memoCategories`.",
    "- Each `memoCategories.groups[].lines` item should preserve one atomic fact, argument, example, number, or constraint. If one source paragraph contains five facts, output five lines.",
    "- Include 8 to 16 meaningful quotes for long interviews when the transcript contains strong management POV. Place quotes under the relevant group.",
    "- Do only minimal MECE cleanup: merge exact duplicates, split mixed paragraphs, and group adjacent related points. Do not delete granular but meaningful facts.",
    "- If the transcript is mostly technical, `memoCategories` can be mostly technical. Do not force Team, GTM, or Financial sections when the transcript did not discuss them.",
    "- Keep interviewer/HMG context out of company facts unless it is necessary to explain the answer. Do not treat the interviewer opinion as company evidence.",
    "- Put quotes directly inside the relevant `memoCategories.groups[].quotes`; do not create a quote dump.",
    "- Avoid translating Chinese transcript content into polished English in `memoCategories`. Translate only if the user explicitly requests English output.",
    "- Stay factual and avoid making up numbers.",
    "- Preservation is more important than compression. Do not drop meaningful facts just to make the note cleaner.",
    "- Organize and clean the material, but do not merge separate points into one if that loses specificity.",
    "- Use broad sections as containers, then use specific section titles for the actual interview subtopics.",
    "- If something does not clearly fit, put it into uncategorized instead of omitting it, but uncategorized should be small.",
    "- Remove speaker labels, timestamps, and obvious transcript artifacts from the content you output.",
    "- Treat Company/interviewee evidence as the primary source for company facts.",
    "- Interviewer questions and context are only for understanding the answer. Do not turn interviewer framing into company claims.",
    "- Exclude HMG, Hyundai Motor, Hyundai, Cradle, investor introduction, how we invest, our portfolio, and interviewer self-introduction from memo facts unless the user explicitly asks for analyst context.",
    "- Keep fundraising details structured and separated when possible.",
    "- fundingTable must only contain short, structured rows. Do not paste transcript chunks into fundingTable cells.",
    "- If a financing detail is important but too messy to structure cleanly, keep it in fundraisingNotes instead of forcing it into fundingTable.",
    "- weeklySlide must be populated even if the main template is not weekly report.",
    "- If the output template is Weekly Report Slide, weeklyReportDraft must be populated and slide-safe.",
    "- weeklyReportDraft.sections must contain exactly 3 sections: Company Background, Product / Technology, Commercialization / Funding.",
    "- Every weeklyReportDraft section must have at least 6 non-empty lines. No empty lines or blank rows in the weekly report core table.",
    "- For Weekly Report Slide, lines should express compact investor hierarchy using labels, `•`, `-`, and `①/②/③` where useful. Do not use visible markdown headers such as `# Product` or `# Technology`.",
    "- weeklyReportDraft section lines must not contain source labels, markdown headings, timestamps, speaker names, JSON keys, or old template/example company facts.",
    "- weeklyReportDraft section lines must not contain the literal string `Mini-head:` and should not repeat the same subhead multiple times.",
    "- If the output template is OI News Report, oiNewsDraft must be concise and slide-safe.",
    "- For OI News Report, each section should normally have 2 to 3 bullets at most, and each bullet should be short enough to fit a portrait one-slide layout.",
    "- quotes should contain important interviewee POV, statements, or comments. Use the interviewee's wording when possible and keep each quote attached to a clear context.",
    "- quotes should contain 3 to 8 items when the transcript has meaningful interviewee statements.",
    "- If public research enrichment is included, prefer verified official or reputable-source facts and do not blur them with transcript claims.",
    ...templateSpecificRules,
    "Meeting packet:",
    transcriptText,
  ].join("\n");
}

function isInterviewMemoTemplate(templateId) {
  return templateId === "interview-knowledge-base" || templateId === "interview-free-style";
}

function getLlmInputLimitForTemplate(templateId) {
  return isInterviewMemoTemplate(templateId) ? INTERVIEW_MEMO_INPUT_CHAR_LIMIT : LLM_INPUT_CHAR_LIMIT;
}

function getArtifactPath(job, artifactType) {
  if (artifactType === "transcript") {
    return job.artifacts?.transcriptPath || job.job?.transcriptPath || "";
  }

  if (artifactType === "role-transcript") {
    return job.artifacts?.roleTranscriptPath || job.job?.roleTranscriptPath || "";
  }

  if (artifactType === "evidence-bank") {
    return job.artifacts?.evidenceBankPath || job.job?.evidenceBankPath || "";
  }

  if (artifactType === "markdown") {
    return job.artifacts?.markdownPath || "";
  }

  if (artifactType === "ppt") {
    return job.artifacts?.pptPath || "";
  }

  throw createError(400, `Unknown artifact type "${artifactType}".`);
}

async function exportJobToObsidian(jobId) {
  const job = await readJob(jobId);
  const obsidian = getObsidianConfig();

  if (!obsidian.configured || !obsidian.vaultPath) {
    throw createError(400, "Obsidian vault was not detected on this machine.");
  }

  const sourceMarkdownPath = job.artifacts?.markdownPath;
  if (!sourceMarkdownPath) {
    throw createError(404, "Markdown artifact not found for this job.");
  }

  const exportDir = path.join(obsidian.vaultPath, "Meeting Automation");
  await fs.mkdir(exportDir, { recursive: true });
  const safeFileName = `${slugify(job.job.company || "meeting")}-${slugify(job.job.title || job.job.id)}.md`;
  const targetPath = path.join(exportDir, safeFileName);
  const markdown = await fs.readFile(sourceMarkdownPath, "utf8");
  await fs.writeFile(targetPath, markdown, "utf8");

  job.artifacts = {
    ...job.artifacts,
    obsidianPath: targetPath,
  };
  await saveJob(job);

  return {
    path: targetPath,
    vaultPath: obsidian.vaultPath,
  };
}

function fallbackExtraction(transcriptText, template) {
  const paragraphs = transcriptText
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean);

  const sentences = transcriptText
    .split(/(?<=[.!?。！？])\s+/)
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence.length > 15);

  const bulletPool = sentences.slice(0, 12);
  const sections =
    template.id === "interview-free-style"
      ? buildFreeStyleFallbackSections(sentences, paragraphs)
      : template.sectionHints.map((hint, index) => ({
          title: hint,
          bullets: bulletPool.slice(index * 2, index * 2 + 2),
        }));
  const actionPattern =
    template.id === "interview-knowledge-base" || template.id === "interview-free-style"
      ? /follow[- ]?up|need(?:s)?\s+(?:to\s+)?(?:verify|validate|confirm|clarify|check)|open question|next step|what to|todo|跟进|确认|验证|问题/i
      : /next|follow|plan|action|todo|will|need|推进|跟进|计划/i;

  return {
    oneSentence: sentences[0] || paragraphs[0] || "Meeting summary generated from transcript.",
    executiveSummary: paragraphs.slice(0, 2).join(" "),
    sections,
    actionItems: sentences.filter((sentence) => actionPattern.test(sentence)).slice(0, 4),
    risks: sentences
      .filter((sentence) => /risk|challenge|uncertain|pressure|competition|风险|压力|挑战/.test(sentence.toLowerCase()))
      .slice(0, 3),
    quotes: sentences.slice(0, 2),
    dataPoints: sentences
      .filter((sentence) => /\d/.test(sentence))
      .slice(0, 4),
    fundraisingNotes: paragraphs.filter((line) => looksLikeFundingLine(line)).slice(0, 8),
    weeklySlide: {
      headline: sentences[0] || "Weekly meeting update",
      updates: bulletPool.slice(0, 3),
      risks: sentences
        .filter((sentence) => /risk|challenge|uncertain|压力|风险/.test(sentence.toLowerCase()))
        .slice(0, 2),
      nextSteps: sentences
        .filter((sentence) => /next|follow|plan|will|need|跟进|计划/.test(sentence.toLowerCase()))
        .slice(0, 2),
    },
    uncategorized: template.id === "interview-free-style" ? [] : paragraphs.slice(0, 12),
  };
}

function buildFreeStyleFallbackSections(sentences, paragraphs) {
  const technical = sentences.filter((sentence) =>
    /tech|technical|architecture|model|training|inference|sensor|platform|product|engineering|deployment|hardware|software|roadmap|algorithm|stack|perception|autonomy/i.test(
      sentence.toLowerCase(),
    ),
  );
  const founder = sentences.filter((sentence) =>
    /founder|ceo|cto|background|career|experience|management|joined|worked|built|resume|education|previously/i.test(
      sentence.toLowerCase(),
    ),
  );
  const commercial = sentences.filter((sentence) =>
    /customer|commercial|gtm|go-to-market|sales|partner|deployment|traction|market|pipeline|order|backlog|revenue/i.test(
      sentence.toLowerCase(),
    ),
  );
  const financing = sentences.filter((sentence) => looksLikeFundingLine(sentence) || /\b(capital|listing|cap table)\b/i.test(sentence));
  const mixedDiligence = sentences.filter((sentence) =>
    /business model|strategy|competition|unit economics|manufacturing|supply chain|margin|cost|expansion|risk/i.test(
      sentence.toLowerCase(),
    ),
  );

  const sections = [
    { title: technical.length ? "Technical / Product Discussion" : "Primary Discussion", bullets: (technical.length ? technical : sentences).slice(0, 6) },
    { title: founder.length ? "Founder / Management Background" : "", bullets: founder.slice(0, 5) },
    { title: commercial.length ? "GTM / Customer Traction" : "", bullets: commercial.slice(0, 5) },
    { title: financing.length ? "Fundraising / Pre-IPO Discussion" : "", bullets: financing.slice(0, 5) },
    { title: mixedDiligence.length ? "Mixed Diligence Topics" : "", bullets: mixedDiligence.slice(0, 4) },
  ].filter((section) => section.title && section.bullets.length);

  if (sections.length) {
    return sections;
  }

  return [
    {
      title: "Primary Discussion",
      bullets: (paragraphs.length ? paragraphs : sentences).slice(0, 6),
    },
  ];
}

function renderOutput(template, structured) {
  if (template.id === "weekly-report") {
    return renderWeeklyMarkdown(structured);
  }

  if (template.id === "oi-news-report") {
    return renderOiNewsMarkdown(structured);
  }

  if (template.id === "interview-free-style") {
    return renderFreeStyleInterviewMarkdown(structured);
  }

  return renderInterviewMarkdown(structured);
}

function buildFreeStyleSections(structured) {
  return dedupeByTitle(
    (structured.sections || [])
      .map((section, index) => {
        const bullets = dedupeLines((section.bullets || []).map(cleanTranscriptArtifact).filter(Boolean));
        if (!bullets.length) {
          return null;
        }

        return {
          title: inferFreeStyleSectionTitle(section.title, bullets, index),
          bullets,
        };
      })
      .filter(Boolean),
  ).slice(0, 6);
}

function inferFreeStyleSectionTitle(originalTitle, bullets, index) {
  const combined = [originalTitle, ...bullets].join(" ").toLowerCase();
  const generic = /^(interview focus|core takeaways|primary themes|discussion|section\s*\d+)$/i.test(String(originalTitle || "").trim());

  if (!generic && String(originalTitle || "").trim()) {
    return cleanTranscriptArtifact(originalTitle);
  }

  if (/tech|technical|architecture|model|training|inference|sensor|platform|product|engineering|deployment/i.test(combined)) {
    return "Technical / Product Discussion";
  }
  if (/go-to-market|customer|commercial|commercialization|pipeline|deployment|sales|partner|market/i.test(combined)) {
    return "GTM / Customer Traction";
  }
  if (/fund|financing|valuation|ipo|pre-ipo|capital|investor|shareholder/i.test(combined)) {
    return "Fundraising / Pre-IPO Discussion";
  }
  if (/founder|ceo|cto|team|management|background|career|experience|resume|education/i.test(combined)) {
    return "Founder / Management Background";
  }
  if (/business model|strategy|competition|unit economics|manufacturing|supply chain|margin|cost|expansion|risk/i.test(combined)) {
    return "Mixed Diligence Topics";
  }

  return index === 0 ? "Primary Discussion" : `Discussion Theme ${index + 1}`;
}

function dedupeByTitle(sections) {
  const seen = new Set();
  return sections.filter((section) => {
    const key = cleanGeneratedText(section.title).toLowerCase();
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function renderFreeStyleInterviewMarkdown(structured) {
  const meta = structured.meeting.meta;
  const focusSections = buildFreeStyleSections(structured);
  const sourceNotes = buildInterviewSourceNotes(structured);
  const focusProfile = structured.focusProfile || deriveFocusProfile(structured);
  const recommendation = structured.templateRecommendation || deriveTemplateRecommendation(structured);
  const reviewQueue = dedupeLines([
    ...deriveOpenQuestions(structured),
    ...(structured.risks || []).slice(0, 6).map((item) => `Risk: ${item}`),
    ...(structured.actionItems || []).slice(0, 6).map((item) => `Follow-up: ${item}`),
  ]).slice(0, 10);
  const evidenceLines = dedupeLines([
    ...(structured.dataPoints || []),
    ...deriveFundraisingNotes(structured),
  ]).slice(0, 12);
  const uncategorizedLines = filterFreeStyleAdditionalNotes(
    deriveUncategorizedLines(structured, focusSections.map((section) => section.title)),
    focusSections,
  );

  const lines = [
    "---",
    "type: interview-free-style",
    `company: ${meta.company}`,
    `date: ${structured.createdAt.slice(0, 10)}`,
    `meeting_type: ${meta.meetingType}`,
    `participants: [${meta.participants.join(", ")}]`,
    `source: ${structured.processing.sourceType}`,
    "analyst: {{owner}}",
    `tags: [${slugify(meta.company)}, ${slugify(meta.meetingType)}, free-style]`,
    "---",
    "",
    `# ${meta.company} | Free Style Memo`,
    "",
    "## Basic Info",
    "",
    `- **Date:** ${structured.createdAt.slice(0, 10)}`,
    `- **Meeting Type:** ${meta.meetingType}`,
    `- **Participants:** ${meta.participants.join(", ") || "TBD"}`,
    `- **Source / Context:** ${structured.processing.sourceType}`,
    `- **Prepared By:** {{owner}}`,
    `- **Detected Focus:** ${focusProfile.primaryLabel}${focusProfile.secondaryLabels?.length ? ` | Secondary: ${focusProfile.secondaryLabels.join(", ")}` : ""}`,
    `- **Template Fit:** ${recommendation.recommendedTemplateName} | ${recommendation.reason}`,
    "",
    "## One-Line Take",
    cleanTranscriptArtifact(structured.summary.oneSentence),
    "",
    "## Executive Summary",
    "",
    "> [!ABSTRACT]",
    "> **Executive Summary**",
  ];

  const executiveSummaryLines = toBulletLines(structured.summary.executiveSummary, 4);
  if (executiveSummaryLines.length === 0) {
    lines.push("> - N/A");
  } else {
    executiveSummaryLines.forEach((line) => lines.push(`> - ${cleanTranscriptArtifact(line)}`));
  }

  lines.push("");
  if (!focusSections.length) {
    lines.push("## Main Discussion");
    lines.push("- No clear dominant discussion block was isolated from the current packet.");
  } else {
    focusSections.forEach((section) => {
      lines.push(`### ${cleanTranscriptArtifact(section.title)}`);
      lines.push("");
      (section.bullets || []).slice(0, 8).forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
      lines.push("");
    });
  }

  if (evidenceLines.length) {
    lines.push("## Key Evidence / Data Points");
    evidenceLines.forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
    lines.push("");
  }

  if (uncategorizedLines.length) {
    lines.push("## Additional Notes / Out-of-Structure Items");
    uncategorizedLines.forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
    lines.push("");
  }

  if (reviewQueue.length) {
    lines.push("## Review Queue");
    reviewQueue.forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
    lines.push("");
  }

  lines.push("<details>");
  lines.push("<summary>Source Notes</summary>");
  lines.push("");
  sourceNotes.forEach((item) => lines.push(`- ${stripMarkdownStrong(item)}`));
  lines.push("");
  lines.push("</details>");
  lines.push("");

  return lines.join("\n");
}

function filterFreeStyleAdditionalNotes(lines, focusSections) {
  const reference = new Set(
    focusSections
      .flatMap((section) => [section.title, ...(section.bullets || [])])
      .map((item) => cleanGeneratedText(item).toLowerCase()),
  );

  return lines.filter((line) => {
    const normalized = cleanGeneratedText(line).toLowerCase();
    if (!normalized) {
      return false;
    }
    if (reference.has(normalized)) {
      return false;
    }
    return ![...reference].some((item) => item && normalized.includes(item));
  });
}

function renderInterviewMarkdown(structured) {
  return renderInterviewMarkdownV3(structured);
  const meta = structured.meeting.meta;
  const focusProfile = structured.focusProfile || deriveFocusProfile(structured);
  const recommendation = structured.templateRecommendation || deriveTemplateRecommendation(structured);
  const businessKeywords = ["business", "strategy", "market", "customer", "商业", "业务", "战略", "市场", "客户", "合作"];
  const productKeywords = ["product", "technology", "tech", "manufacturing", "platform", "roadmap", "产品", "技术", "平台", "研发", "量产", "模型"];
  const commercialKeywords = ["financial", "finance", "commercial", "fund", "ipo", "revenue", "valuation", "商业化", "财务", "融资", "上市", "收入", "估值"];
  const teamKeywords = ["team", "management", "founder", "executive", "leadership", "团队", "管理层", "创始人", "高管"];

  const safeBusinessKeywords = ["business", "strategy", "market", "customer", "commercial", "partnership", "go-to-market"];
  const safeProductKeywords = ["product", "technology", "tech", "manufacturing", "platform", "roadmap", "engineering", "deployment"];
  const safeCommercialKeywords = ["financial", "finance", "commercial", "fund", "ipo", "revenue", "valuation", "margin", "pricing"];
  const safeTeamKeywords = ["team", "management", "founder", "executive", "leadership", "chairman", "chief", "scientist"];

  const businessSection = findSection(structured.sections, safeBusinessKeywords);
  const productSection = findSection(structured.sections, safeProductKeywords);
  const commercialSection = findSection(structured.sections, safeCommercialKeywords);
  const teamSection = findSection(structured.sections, safeTeamKeywords);
  const businessFallback = collectStructuredLinesByKeywords(structured, safeBusinessKeywords, 6);
  const productFallback = collectStructuredLinesByKeywords(structured, safeProductKeywords, 6);
  const commercialFallback = collectStructuredLinesByKeywords(structured, safeCommercialKeywords, 8);
  const teamFallback = collectStructuredLinesByKeywords(structured, safeTeamKeywords, 6);
  const uncategorizedLines = deriveUncategorizedLines(structured, [
    businessSection?.title,
    productSection?.title,
    commercialSection?.title,
    teamSection?.title,
  ]);
  const fundraisingNotes = deriveFundraisingNotes(structured);
  const sourceNotes = buildInterviewSourceNotes(structured);
  const lines = [
    "---",
    "type: interview-memo",
    `company: ${meta.company}`,
    `date: ${structured.createdAt.slice(0, 10)}`,
    `meeting_type: ${meta.meetingType}`,
    `participants: [${meta.participants.join(", ")}]`,
    `source: ${structured.processing.sourceType}`,
    "analyst: {{owner}}",
    `tags: [${slugify(meta.company)}, ${slugify(meta.meetingType)}]`,
    "---",
    "",
    `# ${meta.company} | Interview Memo`,
    "",
    "## Basic Info",
    "",
    `- **Date:** ${structured.createdAt.slice(0, 10)}`,
    `- **Meeting Type:** ${meta.meetingType}`,
    `- **Participants:** ${meta.participants.join(", ") || "TBD"}`,
    `- **Source / Context:** ${structured.processing.sourceType}`,
    `- **Prepared By:** {{owner}}`,
    `- **Detected Focus:** ${focusProfile.primaryLabel}${focusProfile.secondaryLabels?.length ? ` | Secondary: ${focusProfile.secondaryLabels.join(", ")}` : ""}`,
    `- **Template Fit:** ${recommendation.recommendedTemplateName} | ${recommendation.reason}`,
    "",
    "## One-Line Take",
    cleanTranscriptArtifact(structured.summary.oneSentence),
    "",
    "## Executive Summary",
  ];

  lines.push("");
  lines.push("> [!ABSTRACT]");
  lines.push("> **Executive Summary**");
  const executiveSummaryLines = toBulletLines(structured.summary.executiveSummary, 3);
  if (executiveSummaryLines.length === 0) {
    lines.push("> - N/A");
  } else {
    executiveSummaryLines.forEach((line) => lines.push(`> - ${cleanTranscriptArtifact(line)}`));
  }

  lines.push("");
  lines.push("## Core Takeaways");
  const takeawayLines = collectTakeawayLines(structured);
  if (takeawayLines.length === 0) {
    lines.push("- N/A");
  } else {
    takeawayLines.forEach((line, index) => lines.push(`- **Takeaway ${index + 1}:** ${cleanTranscriptArtifact(line)}`));
  }

  lines.push("");
  lines.push("## Business / Strategy");
  pushLabeledSection(lines, [
    ["Business model", bulletAt(businessSection, 0, businessFallback)],
    ["Go-to-market", bulletAt(businessSection, 1, businessFallback)],
    ["Customer traction", bulletAt(businessSection, 2, businessFallback)],
    ["Strategic priorities", bulletAt(businessSection, 3, businessFallback)],
    ["Geographic expansion", bulletAt(businessSection, 4, businessFallback)],
  ]);

  lines.push("");
  lines.push("## Product / Technology");
  pushLabeledSection(lines, [
    ["Core product", bulletAt(productSection, 0, productFallback)],
    ["Technology stack", bulletAt(productSection, 1, productFallback)],
    ["Differentiation", bulletAt(productSection, 2, productFallback)],
    ["Roadmap", bulletAt(productSection, 3, productFallback)],
    ["Manufacturing / deployment readiness", bulletAt(productSection, 4, productFallback)],
  ]);

  lines.push("");
  lines.push("## Commercial / Financial Signals");
  pushLabeledSection(lines, [
    ["Revenue / monetization", bulletAt(commercialSection, 0, commercialFallback)],
    ["Margins / unit economics", bulletAt(commercialSection, 1, commercialFallback)],
    ["Funding status", bulletAt(commercialSection, 2, commercialFallback)],
    ["Valuation / financing context", bulletAt(commercialSection, 3, commercialFallback)],
    ["IPO / exit framing", bulletAt(commercialSection, 4, commercialFallback)],
  ]);

  lines.push("");
  lines.push("### Fundraising Snapshot");
  lines.push("");
  lines.push("| Round | Raised | Valuation | Key Shareholders |");
  lines.push("| --- | --- | --- | --- |");
  deriveFundraisingRows(structured).forEach((row) => lines.push(`| ${row.join(" | ")} |`));

  if (fundraisingNotes.length) {
    lines.push("");
    lines.push("### Fundraising Notes");
    fundraisingNotes.forEach((item) => lines.push(`- ${item}`));
  }

  lines.push("");
  lines.push("## Team");
  pushLabeledSection(lines, [
    ["Founder / CEO", bulletAt(teamSection, 0, teamFallback)],
    ["Key executives", bulletAt(teamSection, 1, teamFallback)],
    ["Timeline / Key milestones", bulletAt(teamSection, 2, teamFallback)],
    ["What seems strong", bulletAt(teamSection, 3, teamFallback)],
    ["What still needs validation", bulletAt(teamSection, 4, teamFallback)],
  ]);

  lines.push("");
  lines.push("## Full Section Capture");
  const fullSectionBlocks = deriveFullSectionCapture(structured);
  if (fullSectionBlocks.length === 0) {
    lines.push("- N/A");
  } else {
    fullSectionBlocks.forEach((line) => lines.push(line));
  }

  lines.push("");
  lines.push("## Key Evidence / Data Points");
  if (structured.dataPoints.length === 0) {
    lines.push("- N/A");
  } else {
    structured.dataPoints.slice(0, 12).forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
  }

  lines.push("");
  if (uncategorizedLines.length) {
    lines.push("## Additional Notes / Uncategorized");
    uncategorizedLines.forEach((item) => lines.push(`- ${item}`));
    lines.push("");
  }

  lines.push("");
  lines.push("## Quotes / Management Claims");
  if (structured.quotes.length === 0) {
    lines.push('- "N/A"');
  } else {
    structured.quotes.slice(0, 6).forEach((item) => lines.push(`- "${stripWrappingQuotes(cleanTranscriptArtifact(item))}"`));
  }

  lines.push("");
  lines.push("## Review Queue");
  lines.push("");
  lines.push("### Risks / Red Flags");
  if (structured.risks.length === 0) {
    lines.push("- **Risk 1:** None highlighted yet");
  } else {
    structured.risks.slice(0, 8).forEach((item, index) => lines.push(`- **Risk ${index + 1}:** ${cleanTranscriptArtifact(item)}`));
  }

  lines.push("");
  lines.push("### Open Questions");
  const openQuestions = deriveOpenQuestions(structured);
  if (openQuestions.length === 0) {
    lines.push("- N/A");
  } else {
    openQuestions.forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
  }

  lines.push("");
  lines.push("### What To Verify Next");
  if (structured.actionItems.length === 0) {
    lines.push("- N/A");
  } else {
    structured.actionItems.slice(0, 10).forEach((item) => lines.push(`- ${cleanTranscriptArtifact(item)}`));
  }

  lines.push("");
  lines.push("## Analyst View");
  lines.push("");
  lines.push("> [!NOTE]");
  lines.push("> **Analyst POV**");
  lines.push(">");
  lines.push("> **What Increased Conviction**");
  const convictionUp = structured.dataPoints.slice(0, 2);
  if (convictionUp.length === 0) {
    lines.push("> - N/A");
  } else {
    convictionUp.forEach((item) => lines.push(`> - ${item}`));
  }

  lines.push(">");
  lines.push("> **What Reduced Conviction**");
  const convictionDown = structured.risks.slice(0, 2);
  if (convictionDown.length === 0) {
    lines.push("> - N/A");
  } else {
    convictionDown.forEach((item) => lines.push(`> - ${item}`));
  }

  lines.push(">");
  lines.push("> **Current Assessment**");
  lines.push(`> ${structured.summary.executiveSummary || structured.summary.oneSentence || "N/A"}`);

  lines.push("");
  lines.push("<details>");
  lines.push("<summary>Source Notes</summary>");
  lines.push("");
  sourceNotes.forEach((item) => lines.push(`- ${stripMarkdownStrong(item)}`));
  lines.push("");
  lines.push("</details>");
  lines.push("");
  return lines.join("\n");
}

function renderInterviewMarkdownV2(structured) {
  const meta = structured.meeting.meta;
  const focusProfile = structured.focusProfile || deriveFocusProfile(structured);
  const sectionDefinitions = [
    { title: "Team", keywords: ["founder", "ceo", "cto", "executive", "leadership", "chairman", "chief", "scientist", "background", "career"] },
    { title: "Business / Strategy", keywords: ["business", "strategy", "market", "customer", "commercial", "partnership", "go-to-market"] },
    { title: "Product / Technology", keywords: ["product", "technology", "tech", "manufacturing", "platform", "roadmap", "engineering", "deployment"] },
    { title: "Commercial / Financial Signals", keywords: ["financial", "finance", "commercial", "funding", "fundraising", "financing", "ipo", "revenue", "valuation", "margin", "pricing"] },
  ];
  const compactSections = attachUnassignedQuotesToSections(buildCompactInterviewSections(structured, sectionDefinitions), structured);
  const sourceNotes = buildInterviewSourceNotes(structured);
  const investmentTakeLines = dedupeLines(
    collectTakeawayLines(structured)
      .map(cleanTranscriptArtifact)
      .filter(Boolean)
      .filter((line) => cleanGeneratedText(line).toLowerCase() !== cleanGeneratedText(structured.summary.oneSentence).toLowerCase()),
  ).slice(0, 4);
  const fundraisingRows = deriveFundraisingRows(structured);
  const fundraisingNotes = deriveFundraisingNotes(structured);
  const reviewQueue = buildCompactReviewQueue(structured);
  const additionalNotes = deriveCompactAdditionalNotes(structured, compactSections);
  const lines = [
    "---",
    "type: interview-memo",
    `company: ${meta.company}`,
    `date: ${structured.createdAt.slice(0, 10)}`,
    `meeting_type: ${meta.meetingType}`,
    `participants: [${meta.participants.join(", ")}]`,
    `source: ${structured.processing.sourceType}`,
    "analyst: {{owner}}",
    `tags: [${slugify(meta.company)}, ${slugify(meta.meetingType)}]`,
    "---",
    "",
    `# ${meta.company} | Interview Memo`,
    `_${structured.createdAt.slice(0, 10)} | ${meta.meetingType} | Focus: ${focusProfile.primaryLabel}_`,
    "",
    "## Investment Take",
    "",
    "> [!ABSTRACT]",
    `> **Bottom line:** ${emphasizeMemoLine(structured.summary.oneSentence || "N/A")}`,
  ];

  investmentTakeLines.forEach((line) => lines.push(`> - ${emphasizeMemoLine(line)}`));

  compactSections.forEach((section) => {
    lines.push("");
    lines.push(`## ${section.title}`);
    section.groups.forEach((group) => {
      if (group.title && group.title !== section.title) {
        lines.push(`- **${group.title}:**`);
        group.lines.forEach((item) => lines.push(`  - ${emphasizeMemoLine(item)}`));
      } else {
        group.lines.forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
      }
    });
    section.quotes.forEach((item) => lines.push(`> [!QUOTE] Management claim: ${emphasizeMemoLine(stripWrappingQuotes(item))}`));

    if (section.title === "Commercial / Financial Signals" && hasUsefulFundraisingContent(fundraisingRows, fundraisingNotes)) {
      lines.push("");
      lines.push("### Fundraising Snapshot");
      lines.push("");
      lines.push("| Round | Raised | Valuation | Key Shareholders |");
      lines.push("| --- | --- | --- | --- |");
      fundingRowsForDisplay(fundraisingRows).forEach((row) => lines.push(`| ${row.map(stripTablePipe).join(" | ")} |`));
      if (fundraisingNotes.length) {
        lines.push("");
        fundraisingNotes.slice(0, 4).forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
      }
    }
  });

  if (reviewQueue.length) {
    lines.push("");
    lines.push("## Open Items / Follow-ups");
    reviewQueue.forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
  }

  lines.push("");
  lines.push("## Analyst View");
  lines.push("");
  lines.push("> [!NOTE]");
  lines.push("> **Analyst POV**");
  lines.push(`> **Current assessment:** ${emphasizeMemoLine(cleanGeneratedText(structured.summary.oneSentence || "N/A"))}`);

  if (additionalNotes.length) {
    lines.push("");
    lines.push("<details>");
    lines.push("<summary>Unclassified Residual Notes</summary>");
    lines.push("");
    additionalNotes.forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
    lines.push("");
    lines.push("</details>");
  }

  lines.push("");
  lines.push("<details>");
  lines.push("<summary>Source Notes</summary>");
  lines.push("");
  sourceNotes.forEach((item) => lines.push(`- ${stripMarkdownStrong(item)}`));
  lines.push("");
  lines.push("</details>");
  lines.push("");
  return lines.join("\n");
}

function renderInterviewMarkdownV3(structured) {
  const meta = structured.meeting.meta;
  const focusProfile = structured.focusProfile || deriveFocusProfile(structured);
  const sourceNotes = buildInterviewSourceNotes(structured);
  const llmMemoCategories = buildLlmMemoCategorySections(structured);
  const groupedSections = llmMemoCategories.length ? llmMemoCategories : buildMemoV3Sections(structured);
  const investmentTakeLines = dedupeLines(
    [
      ...toBulletLines(structured.summary.executiveSummary, 5),
      ...collectTakeawayLines(structured),
    ]
      .map(cleanGeneratedText)
      .filter(Boolean)
      .filter((line) => cleanGeneratedText(line).toLowerCase() !== cleanGeneratedText(structured.summary.oneSentence).toLowerCase()),
  ).slice(0, 5);
  const openItems = buildCompactReviewQueue(structured);
  const residualNotes = buildMemoV3ResidualNotes(structured, groupedSections, openItems);
  const lines = [
    "---",
    "type: interview-memo",
    `company: ${meta.company}`,
    `date: ${structured.createdAt.slice(0, 10)}`,
    `meeting_type: ${meta.meetingType}`,
    `participants: [${meta.participants.join(", ")}]`,
    `source: ${structured.processing.sourceType}`,
    "analyst: {{owner}}",
    `tags: [${slugify(meta.company)}, ${slugify(meta.meetingType)}]`,
    "---",
    "",
    `# ${meta.company} | Interview Memo`,
    `_${structured.createdAt.slice(0, 10)} | ${meta.meetingType} | Focus: ${focusProfile.primaryLabel}${focusProfile.secondaryLabels?.length ? `; secondary: ${focusProfile.secondaryLabels.join(", ")}` : ""}_`,
    "",
    "## Investment Take",
    "",
    "> [!ABSTRACT]",
    `> **Bottom line:** ${emphasizeMemoLine(structured.summary.oneSentence || "N/A")}`,
  ];

  investmentTakeLines.forEach((line) => lines.push(`> - ${emphasizeMemoLine(line)}`));

  groupedSections.forEach((section) => {
    lines.push("");
    lines.push(`## ${section.title}`);
    section.groups.forEach((group) => {
      if (isGenericMemoGroupTitle(group.title, section.title)) {
        group.lines.forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
        group.quotes.forEach((quote) => lines.push(`> [!QUOTE] "${emphasizeMemoLine(stripWrappingQuotes(quote))}"`));
        return;
      }
      lines.push(`- **${group.title}:**`);
      group.lines.forEach((item) => lines.push(`  - ${emphasizeMemoLine(item)}`));
      group.quotes.forEach((quote) => lines.push(`  > [!QUOTE] "${emphasizeMemoLine(stripWrappingQuotes(quote))}"`));
    });

    if (section.shouldRenderFundingSnapshot && hasUsefulFundraisingContent(section.fundraisingRows, section.fundraisingNotes)) {
      lines.push("");
      lines.push("### Fundraising Snapshot");
      lines.push("");
      lines.push("| Round | Raised | Valuation | Key Shareholders |");
      lines.push("| --- | --- | --- | --- |");
      const displayRows = fundingRowsForDisplay(section.fundraisingRows);
      displayRows.forEach((row) => lines.push(`| ${row.map(stripTablePipe).join(" | ")} |`));
      if (!displayRows.length) {
        section.fundraisingNotes.slice(0, 6).forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
      }
    }
  });

  if (openItems.length) {
    lines.push("");
    lines.push("## Open Questions / Follow-ups");
    openItems.forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
  }

  lines.push("");
  lines.push("## Analyst View");
  lines.push("");
  lines.push("> [!NOTE]");
  lines.push("> **Analyst POV**");
  lines.push(`> **Current assessment:** ${emphasizeMemoLine(cleanGeneratedText(structured.summary.oneSentence || "N/A"))}`);

  if (residualNotes.length) {
    lines.push("");
    lines.push("## Other / Unclassified");
    residualNotes.forEach((item) => lines.push(`- ${emphasizeMemoLine(item)}`));
  }

  lines.push("");
  lines.push("<details>");
  lines.push("<summary>Source Notes</summary>");
  lines.push("");
  sourceNotes.forEach((item) => lines.push(`- ${stripMarkdownStrong(item)}`));
  lines.push("");
  lines.push("</details>");
  lines.push("");
  return lines.join("\n");
}

function buildLlmMemoCategorySections(structured) {
  if (!Array.isArray(structured.memoCategories) || !structured.memoCategories.length) {
    return [];
  }

  const sections = structured.memoCategories
    .map((category) => {
      const groups = [];

      (category.groups || []).forEach((group) => {
        const title = cleanTranscriptArtifact(group.title || category.title || "Discussion");
        const lines = sanitizeMemoCategoryLines(group.lines, 80);
        const quotes = sanitizeMemoCategoryLines(group.quotes, 16);
        if (lines.length || quotes.length) {
          groups.push({ title, lines, quotes });
        }
      });

      const directLines = sanitizeMemoCategoryLines(category.lines, 80);
      const directQuotes = sanitizeMemoCategoryLines(category.quotes, 16);
      if (directLines.length || directQuotes.length) {
        groups.push({
          title: cleanTranscriptArtifact(category.title || "Discussion"),
          lines: directLines,
          quotes: directQuotes,
        });
      }

      return {
        key: "llm-category",
        title: cleanTranscriptArtifact(category.title || "Discussion"),
        groups,
        fundraisingRows: [],
        fundraisingNotes: [],
        shouldRenderFundingSnapshot: false,
      };
    })
    .filter((section) => section.title && section.groups.length);

  const fundingRows = deriveFundraisingRows(structured);
  const fundingNotes = deriveFundraisingNotes(structured);
  if (hasUsefulFundraisingContent(fundingRows, fundingNotes)) {
    const fundingSection =
      sections.find((section) => /融资|估值|上市|股东|fund|financing|valuation|ipo|shareholder/i.test(section.title)) ||
      sections.find((section) =>
        section.groups.some((group) =>
          [group.title, ...group.lines].some((line) => /融资|估值|上市|股东|fund|financing|valuation|ipo|shareholder/i.test(line)),
        ),
      );
    if (fundingSection) {
      fundingSection.fundraisingRows = fundingRows;
      fundingSection.fundraisingNotes = fundingNotes;
      fundingSection.shouldRenderFundingSnapshot = true;
    }
  }

  return backfillMemoSectionsFromTranscript(structured, sections);
}

function backfillMemoSectionsFromTranscript(structured, sections) {
  const transcript = structured.meeting?.transcript || "";
  const transcriptEffectiveLength = effectiveMemoTextLength(transcript);
  if (!transcriptEffectiveLength || !sections.length) {
    return sections;
  }

  const currentEffectiveLength = effectiveMemoTextLength(
    sections
      .flatMap((section) => section.groups)
      .flatMap((group) => [...group.lines, ...group.quotes])
      .join("\n"),
  );
  const currentRatio = currentEffectiveLength / transcriptEffectiveLength;
  if (currentRatio >= INTERVIEW_MEMO_MIN_EFFECTIVE_RATIO) {
    return sections;
  }

  const targetLength = Math.floor(transcriptEffectiveLength * INTERVIEW_MEMO_MIN_EFFECTIVE_RATIO);
  const existingText = cleanGeneratedText(
    sections
      .flatMap((section) => [section.title, ...section.groups.flatMap((group) => [group.title, ...group.lines, ...group.quotes])])
      .join("\n"),
  ).toLowerCase();
  const backfillLines = extractBackfillMemoLines(transcript, existingText, targetLength - currentEffectiveLength);
  if (!backfillLines.length) {
    return sections;
  }

  backfillLines.forEach((line) => {
    const placement = findMemoBackfillPlacement(sections, line);
    const section = placement.section || ensureBackfillSection(sections);
    const group = placement.group || ensureBackfillGroup(section, deriveBackfillGroupTitle(line));
    group.lines.push(line);
  });

  return sections;
}

function effectiveMemoTextLength(value) {
  return String(value || "")
    .replace(/^---[\s\S]*?---\s*/, "")
    .replace(/<details>[\s\S]*?<\/details>/g, "")
    .replace(/\s/g, "")
    .length;
}

function extractBackfillMemoLines(transcript, existingText, targetAdditionalLength) {
  const seen = new Set();
  const candidates = splitIntoBackfillMemoUnits(transcript)
    .map((line) => normalizeMemoSectionLine(line))
    .filter((line) => isUsefulMemoBackfillLine(line))
    .filter((line) => {
      const key = cleanGeneratedText(line).toLowerCase();
      if (!key || seen.has(key)) {
        return false;
      }
      seen.add(key);
      if (existingText.includes(key) || key.length < 18) {
        return false;
      }
      const compactKey = key.slice(0, Math.min(32, key.length));
      return !existingText.includes(compactKey);
    });

  const selected = [];
  let addedLength = 0;
  for (const line of candidates) {
    selected.push(line);
    addedLength += effectiveMemoTextLength(line);
    if (selected.length >= INTERVIEW_MEMO_MAX_BACKFILL_LINES || addedLength >= targetAdditionalLength) {
      break;
    }
  }

  return selected;
}

function isUsefulMemoBackfillLine(line) {
  const text = cleanGeneratedText(line);
  if (text.length < 24 || text.length > 260 || looksLikeMostlyTranscriptGarbage(text) || isStrictMemoActionItem(text)) {
    return false;
  }

  if (/^(Keywords|Transcript|GENROBOT AI|Meeting|Source)\b/i.test(text)) {
    return false;
  }

  if (/^(话|型|们|建|据|望|在|断|够|只是|是一个东西|的模态|之后|Data acquisition system)/.test(text)) {
    return false;
  }

  if (/(skin load|sentry|声音\s*load|The shipping|uploaded)/i.test(text)) {
    return false;
  }

  if (!/[。！？!?]$/.test(text) && /[但我你他她它这那的了和与把将构装模设不]$/.test(text)) {
    return false;
  }

  return /数据|模型|模态|采集|硬件|传感器|视觉|触觉|评测|机器人|客户|收入|价格|融资|估值|股东|团队|创始|合规|海外|开源|小时|场景|精度|延迟|SLAM|VLA|VLM|world model|foundation model|customer|revenue|pricing|valuation|funding|sensor|hardware|robot|model|data/i.test(
    text,
  );
}

function splitIntoBackfillMemoUnits(text) {
  const allLines = String(text || "")
    .split(/\r?\n+/)
    .map((line) => normalizeMemoSectionLine(line))
    .filter(Boolean)
    .filter((line) => !/^(GENROBOT AI|Meeting|Source)\b/i.test(line));
  const transcriptMarkerIndex = allLines.findIndex((line) => /^Transcript:?$/i.test(line));
  const rawLines = (transcriptMarkerIndex >= 0 ? allLines.slice(transcriptMarkerIndex + 1) : allLines)
    .filter((line) => !/^(Keywords?|Transcript):?$/i.test(line))
    .filter((line) => !/^\d{4}年\d{1,2}月\d{1,2}日|^\d+\s*小时|硬件\s+模态\s+机器人|key milestone|商业模式\s+技术路线\s+训练数据/i.test(line));

  const units = [];
  let buffer = "";
  rawLines.forEach((line) => {
    if (!buffer) {
      buffer = line;
    } else if (shouldJoinBackfillLine(buffer, line)) {
      buffer = `${buffer}${line}`;
    } else {
      units.push(buffer);
      buffer = line;
    }

    if (/[。！？!?]$/.test(buffer) || buffer.length >= 220) {
      units.push(buffer);
      buffer = "";
    }
  });
  if (buffer) {
    units.push(buffer);
  }

  return units
    .flatMap((unit) => splitLongBackfillUnit(unit))
    .map((unit) => unit.trim())
    .filter(Boolean);
}

function shouldJoinBackfillLine(previous, next) {
  if (/[。！？!?]$/.test(previous)) {
    return false;
  }
  if (/^[-*•]|\d+[.)]/.test(next)) {
    return false;
  }
  if (/^[A-Z][A-Za-z\s]{0,30}:$/.test(next)) {
    return false;
  }
  return previous.length < 180;
}

function splitLongBackfillUnit(unit) {
  const text = String(unit || "").trim();
  if (text.length <= 260) {
    return [text];
  }

  const parts = text
    .split(/(?<=[。！？!?])/)
    .map((item) => item.trim())
    .filter(Boolean);
  if (parts.length <= 1) {
    return [];
  }

  const chunks = [];
  let buffer = "";
  parts.forEach((part) => {
    if (!buffer) {
      buffer = part;
    } else if ((buffer + part).length <= 240) {
      buffer += part;
    } else {
      chunks.push(buffer);
      buffer = part;
    }
  });
  if (buffer) {
    chunks.push(buffer);
  }
  return chunks;
}

function findMemoBackfillPlacement(sections, line) {
  const lineWords = significantMemoWords(line);
  let best = { section: null, group: null, score: 0 };
  sections.forEach((section) => {
    section.groups.forEach((group) => {
      const groupText = cleanGeneratedText([section.title, group.title, ...group.lines].join(" ")).toLowerCase();
      const score = lineWords.filter((word) => groupText.includes(word)).length;
      if (score > best.score) {
        best = { section, group, score };
      }
    });
  });

  if (best.score >= 2) {
    return best;
  }

  const title = deriveBackfillSectionTitle(line);
  const section =
    sections.find((item) => cleanGeneratedText(item.title).toLowerCase() === cleanGeneratedText(title).toLowerCase()) ||
    ensureBackfillSection(sections, title);
  return { section, group: ensureBackfillGroup(section, deriveBackfillGroupTitle(line)) };
}

function ensureBackfillSection(sections, title = "补充细节") {
  let section = sections.find((item) => item.title === title);
  if (!section) {
    section = {
      key: "backfill",
      title,
      groups: [],
      fundraisingRows: [],
      fundraisingNotes: [],
      shouldRenderFundingSnapshot: false,
    };
    sections.push(section);
  }
  return section;
}

function ensureBackfillGroup(section, title) {
  const normalized = cleanGeneratedText(title).toLowerCase();
  let group = section.groups.find((item) => cleanGeneratedText(item.title).toLowerCase() === normalized);
  if (!group) {
    group = { title, lines: [], quotes: [] };
    section.groups.push(group);
  }
  return group;
}

function deriveBackfillSectionTitle(line) {
  if (/融资|估值|股东|收入|价格|客户|商业|开源|海外|合规|funding|valuation|revenue|pricing|customer/i.test(line)) {
    return "商业化、融资与治理补充";
  }
  if (/团队|创始|CEO|负责人|小鹏|大疆|腾讯|Momentum|team|founder/i.test(line)) {
    return "团队与组织补充";
  }
  return "技术与数据细节补充";
}

function deriveBackfillGroupTitle(line) {
  if (/世界模型|人类模态|第一人称|Benchmark|VLA|VLM/i.test(line)) return "世界模型与人类模态";
  if (/采集|传感器|硬件|摄像头|手套|Eagle|Gripper|IMU|编码器|SLAM/i.test(line)) return "采集硬件与传感器";
  if (/标注|管线|Data Foundation|SSP|处理|利用率|数据平台/i.test(line)) return "数据平台与处理";
  if (/评测|机器人|成功率|闭环|5000|200台/i.test(line)) return "评测闭环";
  if (/客户|价格|收入|开源|商业|海外|合规/i.test(line)) return "商业化与客户";
  if (/融资|估值|股东|国资|主体|Pre|Post/i.test(line)) return "融资与治理";
  if (/团队|创始|CEO|负责人|背景/i.test(line)) return "团队背景";
  return "原文补充";
}

function sanitizeMemoCategoryLines(lines, limit = 80) {
  if (!Array.isArray(lines)) {
    return [];
  }

  return dedupeLines(lines)
    .map(normalizeMemoSectionLine)
    .filter(Boolean)
    .filter((line) => line !== "N/A")
    .filter((line) => !looksLikeMostlyTranscriptGarbage(line))
    .slice(0, limit);
}

function buildMemoV3Sections(structured) {
  const buckets = createMemoV3Buckets();
  const usedLines = new Set();
  const usedQuotes = new Set();
  const sectionQuotes = [];

  (structured.sections || []).forEach((section) => {
    if (isMemoQuoteSection(section?.title)) {
      (section.bullets || []).forEach((rawLine) => {
        const quote = normalizeMemoSectionLine(rawLine);
        if (quote) {
          sectionQuotes.push(quote);
        }
      });
      return;
    }

    const fallbackKey = classifyMemoSection(section);
    (section.bullets || []).forEach((rawLine) => {
      const line = sanitizeMemoV3Line(rawLine, usedLines);
      if (!line) {
        return;
      }
      if (isStrictMemoActionItem(line)) {
        return;
      }
      const targetKey = classifyMemoLine(line, fallbackKey);
      const target = buckets[targetKey] || buckets.other;
      const groupTitle = deriveMemoV3LineGroupTitle(line, targetKey, section.title, target.title, fallbackKey);
      ensureMemoGroup(target, groupTitle).lines.push(line);
    });
  });

  recoverUncoveredTranscriptFacts(structured, buckets, usedLines);
  distributeDataPointsToMemoGroups(structured, buckets, usedLines);
  distributeQuotesToMemoGroups(structured, buckets, usedQuotes, sectionQuotes);
  mergeDuplicateMemoGroups(buckets);

  buckets.commercial.fundraisingRows = deriveFundraisingRows(structured);
  buckets.commercial.fundraisingNotes = deriveFundraisingNotes(structured);
  reduceFundraisingDuplication(buckets.commercial);

  return Object.values(buckets)
    .map((bucket) => ({
      ...bucket,
      shouldRenderFundingSnapshot: bucket.key === "commercial",
      groups: bucket.groups.filter((group) => group.lines.length || group.quotes.length),
    }))
    .filter((bucket) => bucket.groups.length || (bucket.key === "commercial" && hasUsefulFundraisingContent(bucket.fundraisingRows, bucket.fundraisingNotes)));
}

function isMemoQuoteSection(title) {
  return /\b(quotes?|management claims?|management pov|statements?|comments?)\b/i.test(cleanGeneratedText(title || ""));
}

function createMemoV3Buckets() {
  return {
    team: { key: "team", title: "Team", groups: [], fundraisingRows: [], fundraisingNotes: [] },
    business: { key: "business", title: "Business / Strategy", groups: [], fundraisingRows: [], fundraisingNotes: [] },
    product: { key: "product", title: "Product / Technology", groups: [], fundraisingRows: [], fundraisingNotes: [] },
    commercial: { key: "commercial", title: "Commercial / Financial Signals", groups: [], fundraisingRows: [], fundraisingNotes: [] },
    other: { key: "other", title: "Other / Unclassified Themes", groups: [], fundraisingRows: [], fundraisingNotes: [] },
  };
}

function classifyMemoSection(section) {
  const title = cleanGeneratedText(section?.title || "");
  const body = cleanGeneratedText((section?.bullets || []).join(" "));
  const text = `${title}\n${body}`;
  const scores = {
    team: scoreMemoCategory(text, [
      "founder",
      "ceo",
      "cto",
      "executive",
      "management background",
      "career",
      "leadership",
      "team",
      "hiring",
      "organization",
    ]),
    business: scoreMemoCategory(text, [
      "go-to-market",
      "gtm",
      "market",
      "customer",
      "partnership",
      "channel",
      "strategy",
      "business model",
      "commercialization",
      "deployment scenario",
    ]),
    product: scoreMemoCategory(text, [
      "product",
      "technology",
      "technical",
      "data",
      "data strategy",
      "infrastructure",
      "ego centric",
      "ego-centric",
      "sensor",
      "data collection",
      "hardware",
      "device",
      "model",
      "model architecture",
      "world model",
      "foundation model",
      "end-to-end",
      "planning",
      "perception",
      "transformer",
      "algorithm",
      "architecture",
      "platform",
      "roadmap",
      "validation",
      "simulation",
      "labeling",
      "feature store",
    ]),
    commercial: scoreMemoCategory(text, [
      "revenue",
      "revenue model",
      "pricing",
      "margin",
      "unit economics",
      "funding",
      "fundraising",
      "financing",
      "valuation",
      "ipo",
      "pre-ipo",
      "investor",
      "shareholder",
      "monetization",
    ]),
  };

  if (/\b(go-to-market|gtm|customer|channel)\b/i.test(title)) {
    scores.business += 6;
  }
  if (/\b(revenue|pricing|margin|unit economics|funding|financing|valuation|ipo|pre-ipo)\b/i.test(title)) {
    scores.commercial += 6;
  }
  if (/\b(data|ego[- ]?centric|sensor|collection|infrastructure|model|algorithm|technical|technology)\b/i.test(title)) {
    scores.product += 6;
  }

  const ranked = Object.entries(scores).sort((a, b) => b[1] - a[1]);
  return ranked[0][1] > 0 ? ranked[0][0] : "other";
}

function scoreMemoCategory(text, keywords) {
  return keywords.reduce((score, keyword) => score + (keywordMatchesText(keyword, text) ? 1 : 0), 0);
}

function normalizeMemoV3GroupTitle(title, fallbackTitle) {
  const cleaned = cleanTranscriptArtifact(title)
    .replace(/^(Core Takeaways|Business and Strategy|Product and Technology|Commercial and Financial Signals|Quotes and Management Claims)$/i, fallbackTitle)
    .trim();
  return canonicalMemoGroupTitle(cleaned || fallbackTitle);
}

function isGenericMemoGroupTitle(groupTitle, sectionTitle) {
  const group = cleanGeneratedText(groupTitle).toLowerCase();
  const section = cleanGeneratedText(sectionTitle).toLowerCase();
  return !group || group === section || ["business / strategy", "product / technology", "commercial / financial signals", "team"].includes(group);
}

function sanitizeMemoV3Line(value, usedLines) {
  const line = normalizeMemoSectionLine(value);
  if (!line || line === "N/A" || looksLikeMostlyTranscriptGarbage(line)) {
    return "";
  }
  const key = cleanGeneratedText(line).toLowerCase();
  if (!key || usedLines.has(key)) {
    return "";
  }
  usedLines.add(key);
  return line;
}

function sanitizeMemoV3Lines(lines, usedLines, limit) {
  return dedupeLines(lines)
    .map(normalizeMemoSectionLine)
    .filter(Boolean)
    .filter((line) => line !== "N/A")
    .filter((line) => !looksLikeMostlyTranscriptGarbage(line))
    .filter((line) => {
      const key = cleanGeneratedText(line).toLowerCase();
      if (!key || usedLines.has(key)) {
        return false;
      }
      usedLines.add(key);
      return true;
    })
    .slice(0, limit);
}

function isStrictMemoActionItem(value) {
  return /follow[- ]?up|need(?:s)?\s+(?:to\s+)?(?:verify|validate|confirm|clarify|check)|open question|next step|what to|todo|跟进|确认|验证|问题/i.test(
    String(value || ""),
  );
}

function classifyMemoLine(line, fallbackKey = "other") {
  const text = cleanGeneratedText(line);
  if (/\b(revenue|revenue model|pricing|gross margin|margin profile|unit economics|monetization|subscription|service fee|arr|mrr|funding|fundraising|financing|valuation|post-money|ipo|pre-ipo|investor|investors|shareholder|shareholders|cap table|cash runway|raised)\b/i.test(text)) {
    return "commercial";
  }
  if (/\b(go-to-market|gtm|customer|client|pilot customer|oem pilot|channel|partnership|sales motion|commercialization|deployment market|market entry|geographic expansion|distribution)\b/i.test(text)) {
    return "business";
  }
  if (/\b(ego[- ]?centric|data strategy|data collection|sensor rig|sensor|in-vehicle|collection hardware|labeling|simulation|evaluation tooling|model evaluation|training data|data infrastructure|feature store|data pipeline|world model|model architecture|foundation model|end-to-end|e2e|planning model|perception stack|transformer|neural|algorithm|architecture|technical|technology|product roadmap|hardware|device|platform|validation|prototype|manufacturing readiness)\b/i.test(text)) {
    return "product";
  }
  if (/\b(founder|ceo|cto|cfo|management background|career|previously|joined|built|led|team|hiring|organization|leadership|advisor)\b/i.test(text)) {
    return "team";
  }
  return fallbackKey || "other";
}

function deriveMemoV3LineGroupTitle(line, targetKey, sourceTitle, bucketTitle, fallbackKey) {
  const text = cleanGeneratedText(line);
  if (targetKey === "product") {
    if (/\b(world model)\b/i.test(text)) {
      return "World Model";
    }
    if (/\b(model architecture|foundation model|end-to-end|e2e|planning model|perception stack|transformer|neural|algorithmic architecture|technical architecture|system architecture)\b/i.test(text)) {
      return "Technical Architecture";
    }
    if (/\b(sensor rig|in-vehicle|collection hardware|data collection|device|devices)\b/i.test(text)) {
      return "Data Collection";
    }
    if (/\b(ego[- ]?centric|data strategy|labeling|data infrastructure|training data|feature store|data pipeline)\b/i.test(text)) {
      return "Data Strategy";
    }
    if (/\b(model evaluation|evaluation tooling|validation metric|benchmark|simulation)\b/i.test(text)) {
      return "Model Evaluation";
    }
    if (/\b(product roadmap|roadmap|launch|prototype|device|hardware|platform)\b/i.test(text)) {
      return "Product Roadmap";
    }
    if (/\b(manufacturing|sop|production|supply chain|deployment readiness|validation)\b/i.test(text)) {
      return "Deployment Readiness";
    }
  }
  if (targetKey === "business") {
    if (/\b(go-to-market|gtm|customer|pilot customer|oem pilot|channel|sales|commercialization)\b/i.test(text)) {
      return "GTM";
    }
    if (/\b(partnership|ecosystem|strategic|geographic expansion|market entry)\b/i.test(text)) {
      return "Partnerships";
    }
  }
  if (targetKey === "commercial") {
    if (/\b(revenue|pricing|monetization|subscription|service fee|unit economics|margin)\b/i.test(text)) {
      return "Revenue Model";
    }
    if (/\b(funding|fundraising|financing|valuation|post-money|ipo|pre-ipo|investor|investors|shareholder|shareholders|cap table|raised)\b/i.test(text)) {
      return "Fundraising";
    }
  }
  if (targetKey === "team") {
    if (/\b(founder|ceo|cto|cfo|management background|career|previously|joined|built|led)\b/i.test(text)) {
      return "Founder Background";
    }
  }

  if (targetKey === fallbackKey) {
    return canonicalMemoGroupTitle(normalizeMemoV3GroupTitle(sourceTitle, bucketTitle));
  }
  return canonicalMemoGroupTitle(bucketTitle);
}

function canonicalMemoGroupTitle(title) {
  const cleaned = cleanTranscriptArtifact(title).trim();
  const normalized = cleanGeneratedText(cleaned).toLowerCase();
  if (!normalized) {
    return cleaned;
  }

  const mappings = [
    [/world model/, "World Model"],
    [/(model architecture|technical architecture|system architecture|foundation model|perception stack|planning model|end-to-end|e2e)/, "Technical Architecture"],
    [/(model evaluation|evaluation tooling|benchmark|simulation|validation metric)/, "Model Evaluation"],
    [/(data collection|sensor rig|in-vehicle|collection hardware|collection device)/, "Data Collection"],
    [/(data strategy|ego[- ]?centric|labeling|data infrastructure|training data|feature store|data pipeline)/, "Data Strategy"],
    [/(product roadmap|roadmap|launch cadence|prototype)/, "Product Roadmap"],
    [/(manufacturing|deployment readiness|readiness|validation|sop|production|supply chain)/, "Deployment Readiness"],
    [/(gtm|go-to-market|customer traction|customer|pilot customer|commercialization|channel|sales)/, "GTM"],
    [/(partnership|ecosystem|market expansion|geographic expansion|market entry)/, "Partnerships"],
    [/(revenue model|unit economics|pricing|monetization|margin|subscription|service fee)/, "Revenue Model"],
    [/(fundraising|funding|financing|capital markets|valuation|pre-ipo|ipo|shareholder|investor|cap table)/, "Fundraising"],
    [/(founder|management background|leadership|career)/, "Founder Background"],
    [/(team|organization|hiring)/, "Organization"],
  ];

  const match = mappings.find(([pattern]) => pattern.test(normalized));
  return match ? match[1] : cleaned;
}

function distributeDataPointsToMemoGroups(structured, buckets, usedLines) {
  (structured.dataPoints || []).forEach((item) => {
    const line = normalizeMemoSectionLine(item);
    if (!line || usedLines.has(cleanGeneratedText(line).toLowerCase())) {
      return;
    }
    const key = classifyMemoLine(line, classifyMemoSection({ title: "Data / Evidence", bullets: [line] }));
    const bucket = buckets[key] || buckets.other;
    const group = ensureMemoGroup(bucket, deriveMemoV3LineGroupTitle(line, key, "Evidence / Metrics", bucket.title, key));
    group.lines.push(line);
    usedLines.add(cleanGeneratedText(line).toLowerCase());
  });
}

function distributeQuotesToMemoGroups(structured, buckets, usedQuotes, sectionQuotes = []) {
  dedupeLines([...(structured.quotes || []), ...sectionQuotes, ...deriveSupplementalMemoQuotes(structured)])
    .map(cleanTranscriptArtifact)
    .filter(Boolean)
    .slice(0, 14)
    .forEach((quote) => {
      const keyValue = cleanGeneratedText(quote).toLowerCase();
      if (!keyValue || usedQuotes.has(keyValue) || isStrictMemoActionItem(quote)) {
        return;
      }
      const placement = findBestMemoQuotePlacement(buckets, quote);
      const bucketKey = placement?.bucket?.key || classifyMemoLine(quote, classifyMemoSection({ title: "Management claim", bullets: [quote] }));
      const bucket = buckets[bucketKey] && bucketKey !== "other" ? buckets[bucketKey] : chooseFallbackQuoteBucket(buckets, quote);
      const group =
        placement?.group ||
        findBestMemoQuoteGroup(bucket, quote) ||
        ensureMemoGroup(bucket, deriveMemoV3LineGroupTitle(quote, bucket.key, "Management POV / Claims", bucket.title, bucket.key));
      group.quotes.push(quote);
      usedQuotes.add(keyValue);
    });
}

function findBestMemoQuotePlacement(buckets, quote) {
  const quoteText = cleanGeneratedText(quote).toLowerCase();
  const quoteWords = significantMemoWords(quoteText);
  let best = null;

  Object.values(buckets).forEach((bucket) => {
    if (bucket.key === "other") {
      return;
    }
    bucket.groups.forEach((group) => {
      const groupText = cleanGeneratedText([group.title, ...group.lines].join(" ")).toLowerCase();
      const overlap = quoteWords.filter((word) => groupText.includes(word)).length;
      if (overlap > (best?.score || 0)) {
        best = { bucket, group, score: overlap };
      }
    });
  });

  return best?.score >= 2 ? best : null;
}

function findBestMemoQuoteGroup(bucket, quote) {
  const quoteText = cleanGeneratedText(quote).toLowerCase();
  return bucket.groups.find((group) =>
    group.lines.some((line) => {
      const lineWords = cleanGeneratedText(line).toLowerCase().split(/\s+/).filter((word) => word.length > 4);
      return lineWords.some((word) => quoteText.includes(word));
    }),
  );
}

function chooseFallbackQuoteBucket(buckets, quote) {
  const key = classifyMemoLine(quote, "product");
  if (key !== "other" && buckets[key]) {
    return buckets[key];
  }
  const populated = [buckets.product, buckets.business, buckets.commercial, buckets.team].find((bucket) => bucket.groups.length);
  return populated || buckets.product || buckets.other;
}

function significantMemoWords(text) {
  return cleanGeneratedText(text)
    .toLowerCase()
    .split(/[^a-z0-9\u4e00-\u9fff-]+/i)
    .filter((word) => word.length >= 5)
    .filter((word) => !["management", "company", "their", "there", "would", "could", "about", "because", "which", "should"].includes(word))
    .slice(0, 24);
}

function deriveSupplementalMemoQuotes(structured) {
  return splitIntoMemoSentences(structured.meeting?.companyEvidenceText || structured.meeting?.transcript || "")
    .filter((sentence) => /\b(management said|founder said|ceo said|cto said|we believe|we think|we see|we expect|they said|said that|noted that|emphasized|highlighted|confirmed)\b/i.test(sentence))
    .filter((sentence) => !isStrictMemoActionItem(sentence))
    .slice(0, 8);
}

function recoverUncoveredTranscriptFacts(structured, buckets, usedLines) {
  const existing = new Set(
    Object.values(buckets)
      .flatMap((bucket) => bucket.groups)
      .flatMap((group) => group.lines)
      .map((line) => cleanGeneratedText(line).toLowerCase()),
  );

  splitIntoMemoSentences(structured.meeting?.companyEvidenceText || structured.meeting?.transcript || "")
    .filter((sentence) => sentence.length >= 35)
    .filter((sentence) => !looksLikeMostlyTranscriptGarbage(sentence))
    .filter((sentence) => !isStrictMemoActionItem(sentence))
    .slice(0, 80)
    .forEach((sentence) => {
      const line = normalizeMemoSectionLine(sentence);
      const key = cleanGeneratedText(line).toLowerCase();
      if (!key || usedLines.has(key) || existing.has(key)) {
        return;
      }
      if ([...existing].some((item) => item.includes(key) || key.includes(item))) {
        return;
      }
      const targetKey = classifyMemoLine(line, "other");
      if (targetKey === "other") {
        return;
      }
      const bucket = buckets[targetKey] || buckets.other;
      const groupTitle = deriveMemoV3LineGroupTitle(line, targetKey, bucket.title, bucket.title, targetKey);
      ensureMemoGroup(bucket, groupTitle).lines.push(line);
      usedLines.add(key);
      existing.add(key);
    });
}

function splitIntoMemoSentences(text) {
  return String(text || "")
    .split(/\r?\n|(?<=[.!?。！？])\s+/)
    .map(normalizeMemoSectionLine)
    .filter(Boolean);
}

function reduceFundraisingDuplication(commercialBucket) {
  if (!hasUsefulFundraisingContent(commercialBucket.fundraisingRows, commercialBucket.fundraisingNotes)) {
    return;
  }

  commercialBucket.groups = commercialBucket.groups
    .map((group) => {
      if (!/\b(fundraising|funding|financing|capital markets|valuation|ipo|pre-ipo)\b/i.test(group.title)) {
        return group;
      }
      return {
        ...group,
        lines: group.lines.filter((line) => !looksLikeFundingLine(line)),
      };
    })
    .filter((group) => group.lines.length || group.quotes.length);
}

function ensureMemoGroup(bucket, title) {
  let group = bucket.groups.find((item) => cleanGeneratedText(item.title).toLowerCase() === cleanGeneratedText(title).toLowerCase());
  if (!group) {
    group = { title, lines: [], quotes: [] };
    bucket.groups.push(group);
  }
  return group;
}

function mergeDuplicateMemoGroups(buckets) {
  Object.values(buckets).forEach((bucket) => {
    const byTitle = new Map();
    bucket.groups.forEach((group) => {
      const key = cleanGeneratedText(group.title).toLowerCase();
      if (!byTitle.has(key)) {
        byTitle.set(key, { title: group.title, lines: [], quotes: [] });
      }
      const existing = byTitle.get(key);
      existing.lines = dedupeLines([...existing.lines, ...group.lines]);
      existing.quotes = dedupeLines([...existing.quotes, ...group.quotes]);
    });
    bucket.groups = [...byTitle.values()];
  });
}

function buildMemoV3ResidualNotes(structured, groupedSections, openItems) {
  const covered = new Set(
    groupedSections
      .flatMap((section) => section.groups)
      .flatMap((group) => [...group.lines, ...group.quotes])
      .map((item) => cleanGeneratedText(item).toLowerCase()),
  );
  const openItemSet = new Set(
    openItems.map((item) => cleanGeneratedText(String(item || "").replace(/^(Question|Next|Risk):\s*/i, "")).toLowerCase()),
  );

  return dedupeLines([...(structured.uncategorized || [])])
    .map(normalizeMemoSectionLine)
    .filter(Boolean)
    .filter((item) => {
      const key = cleanGeneratedText(item).toLowerCase();
      if (!key || covered.has(key) || openItemSet.has(key)) {
        return false;
      }
      return ![...covered].some((coveredItem) => coveredItem && key.includes(coveredItem));
    })
    .slice(0, 12);
}

function buildCompactInterviewSections(structured, sectionDefinitions) {
  const usedLines = new Set();

  return sectionDefinitions
    .map((definition) => {
      const sectionGroups = collectSectionGroupsForDefinition(structured, definition, usedLines);
      const flatLines = sectionGroups.flatMap((group) => group.lines);
      const normalizedLineSet = new Set(flatLines.map((line) => cleanGeneratedText(line).toLowerCase()));
      const quotes = selectRelevantQuotes(structured, definition.keywords, 2).filter(
        (quote) => !normalizedLineSet.has(cleanGeneratedText(quote).toLowerCase()),
      );

      if (!sectionGroups.length && !quotes.length) {
        return null;
      }

      return {
        title: definition.title,
        lines: flatLines,
        groups: sectionGroups,
        quotes,
      };
    })
    .filter(Boolean);
}

function collectSectionGroupsForDefinition(structured, definition, usedLines) {
  const groups = new Map();
  const addToGroup = (title, value) => {
    const parsed = parseMemoPrefixedLine(value);
    const effectiveTitle = parsed.title || title;
    const line = normalizeMemoSectionLine(parsed.body || value);
    if (!line || line === "N/A" || looksLikeMostlyTranscriptGarbage(line) || !lineMatchesSectionKeywords(line, definition.keywords)) {
      return;
    }

    const normalized = cleanGeneratedText(line).toLowerCase();
    if (!normalized || usedLines.has(normalized)) {
      return;
    }

    usedLines.add(normalized);
    const groupTitle = normalizeMemoGroupTitle(effectiveTitle, definition.title);
    if (!groups.has(groupTitle)) {
      groups.set(groupTitle, []);
    }
    groups.get(groupTitle).push(line);
  };

  (structured.sections || []).forEach((section) => {
    const titleScore = scoreSection(section, definition.keywords);
    if (titleScore <= 0) {
      return;
    }
    (section.bullets || []).forEach((bullet) => addToGroup(section.title, bullet));
  });

  collectStructuredLinesByKeywords(structured, definition.keywords, 8).forEach((line) => addToGroup(definition.title, line));

  return [...groups.entries()]
    .map(([title, lines]) => ({
      title,
      lines: dedupeLines(lines).slice(0, 5),
    }))
    .filter((group) => group.lines.length)
    .slice(0, 4);
}

function parseMemoPrefixedLine(value) {
  const text = cleanTranscriptArtifact(value);
  const match = text.match(/^([^:]{3,90}):\s*(.+)$/);
  if (!match) {
    return {
      title: "",
      body: text,
    };
  }

  const title = match[1].trim();
  const body = match[2].trim();
  if (!title || !body || /^\d+$/.test(title)) {
    return {
      title: "",
      body: text,
    };
  }

  return {
    title,
    body,
  };
}

function normalizeMemoGroupTitle(title, fallbackTitle) {
  const cleaned = cleanTranscriptArtifact(title)
    .replace(/^(Core Takeaways|Business and Strategy|Product and Technology|Commercial and Financial Signals|Quotes and Management Claims)$/i, fallbackTitle)
    .trim();
  return cleaned || fallbackTitle;
}

function selectRelevantQuotes(structured, keywords, limit = 2) {
  return dedupeLines(
    (structured.quotes || [])
      .map(normalizeMemoSectionLine)
      .filter(Boolean)
      .filter((quote) => keywords.some((keyword) => keywordMatchesText(keyword, quote))),
  ).slice(0, limit);
}

function attachUnassignedQuotesToSections(sections, structured) {
  if (!sections.length) {
    return sections;
  }

  const assigned = new Set(
    sections.flatMap((section) => section.quotes || []).map((quote) => cleanGeneratedText(quote).toLowerCase()),
  );
  const existingLines = new Set(
    sections.flatMap((section) => section.lines || []).map((line) => cleanGeneratedText(line).toLowerCase()),
  );
  const unassigned = dedupeLines((structured.quotes || []).map(cleanTranscriptArtifact).filter(Boolean))
    .filter((quote) => !assigned.has(cleanGeneratedText(quote).toLowerCase()))
    .filter((quote) => !existingLines.has(cleanGeneratedText(quote).toLowerCase()))
    .filter((quote) => /\b(said|expects|claimed|noted|confirmed|believes|management|founder|ceo|cto)\b/i.test(quote))
    .slice(0, 3);

  unassigned.forEach((quote) => {
    const target =
      sections.find((section) => lineMatchesSectionKeywords(quote, section.lines.join(" ").split(/\s+/).slice(0, 12))) ||
      sections[0];
    target.quotes = dedupeLines([...(target.quotes || []), quote]).slice(0, 3);
  });

  return sections;
}

function buildCompactReviewQueue(structured) {
  const items = [];
  const seen = new Set();
  const pushItem = (label, value) => {
    const normalizedValue = cleanTranscriptArtifact(String(value || "").replace(/^Need to verify:\s*/i, ""));
    const key = cleanGeneratedText(normalizedValue).toLowerCase();
    if (!normalizedValue || !key || seen.has(key)) {
      return;
    }
    seen.add(key);
    items.push(`${label}: ${normalizedValue}`);
  };

  deriveOpenQuestions(structured).slice(0, 3).forEach((item) => pushItem("Question", item));
  (structured.actionItems || []).filter(isStrictMemoActionItem).slice(0, 5).forEach((item) => pushItem("Next", item));
  (structured.risks || []).slice(0, 3).forEach((item) => pushItem("Risk", item));

  return items.slice(0, 8);
}

function deriveCompactAdditionalNotes(structured, compactSections) {
  const covered = new Set(
    compactSections
      .flatMap((section) => [...section.lines, ...section.quotes])
      .map((item) => cleanGeneratedText(normalizeMemoSectionLine(item)).toLowerCase()),
  );
  const reviewCovered = new Set(
    buildCompactReviewQueue(structured).map((item) =>
      cleanGeneratedText(String(item || "").replace(/^(Watchout|Open question|Next step|Question|Next|Risk):\s*/i, "")).toLowerCase(),
    ),
  );

  return dedupeLines(deriveUncategorizedLines(structured, compactSections.map((section) => section.title)).map(normalizeMemoSectionLine))
    .filter((item) => {
      const normalized = cleanGeneratedText(item).toLowerCase();
      if (covered.has(normalized)) {
        return false;
      }
      return ![...covered].some((coveredItem) => coveredItem && normalized.includes(coveredItem));
    })
    .filter((item) => !reviewCovered.has(cleanGeneratedText(item).toLowerCase()))
    .filter((item) => item.length <= 260)
    .filter((item) => !looksLikeMostlyTranscriptGarbage(item))
    .slice(0, 6);
}

function emphasizeMemoLine(value) {
  const cleaned = String(cleanTranscriptArtifact(value))
    .replace(/\b(Series\s+[A-D]|Pre-IPO|IPO|seed|angel)\b/gi, "**$1**")
    .replace(/\b(Q[1-4]\s*\d{4}|Q[1-4]|FY\d{2,4}|20\d{2})\b/g, "**$1**")
    .replace(/\b(US?\$[\d\.,]+\s*(?:m|mn|million|b|bn|billion)?|RMB[\d\.,]+\s*(?:m|mn|million|b|bn|billion)?|\d+(?:\.\d+)?%|\d+(?:,\d{3})+)\b/gi, "**$1**");
  return cleaned.replace(/\b(management said|management expects|company expects|founder said|CEO said|CTO said)\b/gi, "==$1==");
}

function hasUsefulFundraisingContent(rows, notes) {
  return fundingRowsForDisplay(rows).length > 0 || (notes || []).length > 0;
}

function fundingRowsForDisplay(rows) {
  return (rows || []).filter((row) => row.some((item) => String(item || "").trim() !== "N/A"));
}

function stripTablePipe(value) {
  return String(value || "").replace(/\|/g, "/").trim();
}

function normalizeMemoSectionLine(value) {
  return cleanTranscriptArtifact(value)
    .replace(/^(Core Takeaways|Business and Strategy|Product and Technology|Commercial and Financial Signals|Quotes and Management Claims)\s*:\s*/i, "")
    .replace(/^\s*(?:[-*•]|\d+[.)])\s+/, "")
    .trim();
}

function lineMatchesSectionKeywords(line, keywords) {
  return keywords.some((keyword) => keywordMatchesText(keyword, line));
}

function keywordMatchesText(keyword, text) {
  const needle = String(keyword || "").trim().toLowerCase();
  const haystack = String(text || "").toLowerCase();
  if (!needle || !haystack) {
    return false;
  }

  const escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`(^|[^a-z0-9])${escaped}($|[^a-z0-9])`, "i").test(haystack);
}

function renderWeeklyMarkdown(structured) {
  const meta = structured.meeting.meta;
  const slide = structured.weeklySlide;
  const lines = [
    `# Weekly Report | ${meta.company}`,
    "",
    `**Meeting:** ${meta.title}`,
    `**Type:** ${meta.meetingType}`,
    `**ASR:** ${structured.processing.asrProvider}`,
    `**Created At:** ${structured.createdAt}`,
    "",
    "## Slide Headline",
    slide.headline || structured.summary.oneSentence,
    "",
    "## Core Updates",
  ];

  const updates = slide.updates?.slice(0, 3) || [];
  if (updates.length === 0) {
    lines.push("- N/A");
  } else {
    updates.forEach((item) => lines.push(`- ${item}`));
  }

  lines.push("");
  lines.push("## Risks / Watch Items");
  const risks = slide.risks?.slice(0, 2) || [];
  if (risks.length === 0) {
    lines.push("- No major risks captured");
  } else {
    risks.forEach((item) => lines.push(`- ${item}`));
  }

  lines.push("");
  lines.push("## Next Steps");
  const nextSteps = slide.nextSteps?.slice(0, 2) || [];
  if (nextSteps.length === 0) {
    lines.push("- No next steps captured");
  } else {
    nextSteps.forEach((item) => lines.push(`- ${item}`));
  }

  lines.push("");
  lines.push("## Speaker Notes");
  lines.push(structured.summary.executiveSummary || structured.summary.oneSentence);
  lines.push("");

  return lines.join("\n");
}

function renderOiNewsMarkdown(structured) {
  const meta = structured.meeting.meta;
  const draft = structured.oiNewsDraft || buildOiNewsDraft(structured);
  const lines = [
    `# OI News Report | ${meta.company}`,
    "",
    `**Headline:** ${draft.headline}`,
    `**Dateline:** ${draft.dateline}`,
    `**Source:** ${draft.sourceLine}`,
    "",
    "## Key News",
  ];

  (draft.sections[0]?.lines || []).forEach((item) => lines.push(`- ${item}`));
  lines.push("");
  lines.push("## Strategic Relevance");
  (draft.sections[1]?.lines || []).forEach((item) => lines.push(`- ${item}`));
  lines.push("");
  lines.push("## Watch Items");
  (draft.sections[2]?.lines || []).forEach((item) => lines.push(`- ${item}`));
  lines.push("");
  return lines.join("\n");
}

function findSection(sections, keywords) {
  const scored = (sections || [])
    .map((section) => ({
      section,
      score: scoreSection(section, keywords),
    }))
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score);

  return scored[0]?.section;
}

function scoreSection(section, keywords) {
  const title = String(section?.title || "").toLowerCase();
  const bullets = (section?.bullets || []).join(" ").toLowerCase();
  let score = 0;

  keywords.forEach((keyword) => {
    if (keywordMatchesText(keyword, title)) {
      score += 3;
    }
    if (keywordMatchesText(keyword, bullets)) {
      score += 1;
    }
  });

  return score;
}

function collectStructuredLinesByKeywords(structured, keywords, limit = 8) {
  const lines = [
    ...(structured.sections || []).flatMap((section) => (section.bullets || []).map((item) => `${section.title}: ${item}`)),
    ...(structured.fundraisingNotes || []),
    ...(structured.dataPoints || []),
    ...(structured.materialInsights?.roadmap || []),
    ...(structured.materialInsights?.fundraising || []),
    ...(structured.materialInsights?.capTable || []),
    ...(structured.materialInsights?.unitEconomics || []),
  ];

  return dedupeLines(
    lines
      .map(cleanGeneratedText)
      .filter(Boolean)
      .filter((line) => line.length <= 260)
      .filter((line) => keywords.some((keyword) => keywordMatchesText(keyword, line))),
  ).slice(0, limit);
}

function bulletAt(section, index, fallbackLines = []) {
  return section?.bullets?.[index] || fallbackLines[index] || "N/A";
}

function pushLabeledSection(lines, pairs) {
  pairs.forEach(([label, value]) => {
    lines.push(`- **${label}:** ${cleanTranscriptArtifact(value || "N/A")}`);
  });
}

function toBulletLines(text, limit = 3) {
  if (!text) {
    return [];
  }

  return text
    .split(/\r?\n|(?<=[.!?。！？])\s+/)
    .map((item) => cleanGeneratedText(item))
    .filter(Boolean)
    .slice(0, limit);
}

function decodeTextBuffer(buffer) {
  const utf8Text = buffer.toString("utf8");
  const replacementCount = (utf8Text.match(/�/g) || []).length;
  const suspiciousUtf8 = replacementCount > 3 || /[Ã¥Ã¤Ã¶Ã¼]/.test(utf8Text);

  if (!suspiciousUtf8) {
    return utf8Text.trim();
  }

  try {
    return iconv.decode(buffer, "gb18030").trim();
  } catch (_error) {
    return utf8Text.trim();
  }
}

function normalizeTranscriptLikeText(text) {
  if (!text) {
    return "";
  }

  return text
    .replace(/\r\n/g, "\n")
    .replace(/\u0000/g, "")
    .replace(/^\uFEFF/, "")
    .replace(/^\s*\d{4}[/-]\d{1,2}[/-]\d{1,2}.*$/gm, "")
    .replace(/^\s*(关键词|Key words?)\s*[:：]?\s*$/gim, "Keywords:")
    .replace(/^\s*(文字记录|Transcript)\s*[:：]?\s*$/gim, "Transcript:")
    .replace(/^\s*说话人\s*(\d+)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s*$/gm, "Speaker $1 [$2]")
    .replace(/^\s*Speaker\s*(\d+)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s*$/gim, "Speaker $1 [$2]")
    .replace(/^\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*$/gm, "[$1]")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeTranscriptLikeTextV2(text) {
  if (!text) {
    return "";
  }

  return String(text)
    .replace(/\r\n/g, "\n")
    .replace(/\u0000/g, "")
    .replace(/^\uFEFF/, "")
    .replace(/^\s*\d{4}[/-]\d{1,2}[/-]\d{1,2}.*$/gm, "")
    .replace(/^\s*(?:Key words?|关键词)\s*[:：]?\s*$/gim, "Keywords:")
    .replace(/^\s*(?:Transcript|文字记录)\s*[:：]?\s*$/gim, "Transcript:")
    .replace(/^\s*(?:Speaker|说话人)\s*(\d+)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s*$/gim, "Speaker $1 [$2]")
    .replace(/^\s*(?:Speaker|说话人)\s*\d+\s*[:：-]\s*/gim, "")
    .replace(/\b(?:Speaker|SPEAKER|说话人)\s*\d+\s*(?:\[\d{1,2}:\d{2}(?::\d{2})?\])?\s*[:：-]?\s*/gim, "")
    .replace(/\b(?:Speaker|说话人)\s*\d+\b/gim, "")
    .replace(/\[\d{1,2}:\d{2}(?::\d{2})?\]\s*/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

async function extractDocxText(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const mammothText = (result.value || "").trim();
  if (mammothText) {
    return mammothText;
  }

  const zip = await JSZip.loadAsync(buffer);
  const documentXml = zip.files["word/document.xml"];
  if (!documentXml) {
    return "";
  }

  const xml = await documentXml.async("string");
  return [...xml.matchAll(/<w:t[^>]*>(.*?)<\/w:t>/g)]
    .map((match) => decodeXmlEntities(match[1]))
    .map((item) => item.trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

async function extractPptxText(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => extractSlideNumber(a) - extractSlideNumber(b));

  const slides = [];
  const structuredNotes = [];
  for (const slideFile of slideFiles) {
    const xml = await zip.files[slideFile].async("string");
    const textItems = [...xml.matchAll(/<a:t[^>]*>(.*?)<\/a:t>/g)]
      .map((match) => decodeXmlEntities(match[1]))
      .map((item) => item.trim())
      .filter(Boolean);
    const text = textItems.join("\n");

    if (text) {
      slides.push(`Slide ${extractSlideNumber(slideFile)}\n${text}`);
      const derived = derivePptSlideInsights(textItems, extractSlideNumber(slideFile));
      if (derived.length) {
        structuredNotes.push(...derived);
      }
    }
  }

  const blocks = [];
  if (structuredNotes.length) {
    blocks.push("Structured material notes:");
    structuredNotes.forEach((item) => blocks.push(`- ${item}`));
  }
  if (slides.length) {
    blocks.push(slides.join("\n\n"));
  }

  return blocks.join("\n\n").trim();
}

function extractSpreadsheetText(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const structuredNotes = [];
  const sheetChunks = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false });
    const derived = deriveSheetInsights(sheetName, rows);
    if (derived.length) {
      structuredNotes.push(...derived);
    }
    const text = rows
      .map((row) => row.filter((cell) => cell !== undefined && cell !== null && String(cell).trim() !== "").join(" | "))
      .filter(Boolean)
      .join("\n");

    return text ? `Sheet: ${sheetName}\n${text}` : "";
  }).filter(Boolean);

  const blocks = [];
  if (structuredNotes.length) {
    blocks.push("Structured material notes:");
    structuredNotes.forEach((item) => blocks.push(`- ${item}`));
  }
  if (sheetChunks.length) {
    blocks.push(sheetChunks.join("\n\n"));
  }

  return blocks.join("\n\n").trim();
}

function extractSlideNumber(slideFile) {
  const match = slideFile.match(/slide(\d+)\.xml$/);
  return Number(match?.[1] || 0);
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'");
}

function createEmptyMaterialInsights() {
  return {
    roadmap: [],
    fundraising: [],
    capTable: [],
    unitEconomics: [],
  };
}

function mergeMaterialInsightSets(...sets) {
  const merged = createEmptyMaterialInsights();
  sets.filter(Boolean).forEach((set) => {
    mergeMaterialInsights(merged, set);
  });
  return merged;
}

function mergeMaterialInsights(target, incoming) {
  if (!incoming) {
    return target;
  }

  for (const key of Object.keys(createEmptyMaterialInsights())) {
    const nextItems = Array.isArray(incoming[key]) ? incoming[key] : [];
    const existing = new Set(target[key] || []);
    nextItems.forEach((item) => {
      if (item && !existing.has(item)) {
        target[key].push(item);
        existing.add(item);
      }
    });
  }

  return target;
}

function deriveMaterialInsightsFromText(text) {
  const insights = createEmptyMaterialInsights();
  const lines = String(text || "")
    .split(/\r?\n+/)
    .map((line) => cleanGeneratedText(line))
    .filter(Boolean);

  lines.forEach((line) => {
    if (line.length > 180 || looksLikeMostlyTranscriptGarbage(line)) {
      return;
    }
    if (/roadmap|timeline|milestone|pilot launch|mass production|sop|launch/i.test(line)) {
      insights.roadmap.push(line);
    }
    if (looksLikeFundingLine(line)) {
      insights.fundraising.push(line);
    }
    if (/cap table|shareholder|ownership|stake|holding/i.test(line)) {
      insights.capTable.push(line);
    }
    if (/unit economics|payback|cost|price|margin|revenue|gmv/i.test(line)) {
      insights.unitEconomics.push(line);
    }
  });

  return insights;
}

function derivePptSlideInsights(textItems, slideNumber) {
  const joined = textItems.join(" ").toLowerCase();
  const notes = [];

  if (/roadmap|timeline|milestone|202[4-9]/i.test(joined)) {
    const milestones = textItems.filter((item) => /(20\d{2}|roadmap|timeline|milestone|launch|pilot|mass production|sop)/i.test(item));
    if (milestones.length) {
      notes.push(`Roadmap identified on slide ${slideNumber}: ${milestones.slice(0, 6).join(" ; ")}`);
    }
  }

  if (/unit economics|economics|payback|margin|cost|price|revenue/i.test(joined)) {
    const econPoints = textItems.filter((item) => /(\$|rmb|yuan|cost|price|margin|payback|revenue|gmv|unit economics|month|year)/i.test(item));
    if (econPoints.length) {
      notes.push(`Unit economics signals on slide ${slideNumber}: ${econPoints.slice(0, 6).join(" ; ")}`);
    }
  }

  if (/fundraising|series|pre-ipo|valuation|investor|shareholder|cap table/i.test(joined)) {
    const fundingPoints = textItems.filter((item) => /(series|seed|pre-ipo|valuation|investor|shareholder|fundraising|raised)/i.test(item));
    if (fundingPoints.length) {
      notes.push(`Fundraising or cap table signals on slide ${slideNumber}: ${fundingPoints.slice(0, 6).join(" ; ")}`);
    }
  }

  return notes;
}

function deriveSheetInsights(sheetName, rows) {
  const notes = [];
  if (!rows.length) {
    return notes;
  }

  const header = rows[0].map((cell) => String(cell || "").trim().toLowerCase());
  const allRows = rows.map((row) => row.map((cell) => String(cell || "").trim()));

  if (header.some((item) => /round|raised|valuation|investor|shareholder/.test(item))) {
    const entries = allRows.slice(1, 4).map((row) => row.filter(Boolean).join(" | ")).filter(Boolean);
    if (entries.length) {
      notes.push(`Fundraising table detected in sheet ${sheetName}: ${entries.join(" ; ")}`);
    }
  }

  if (header.some((item) => /shareholder|ownership|stake|holding|cap table/.test(item))) {
    const entries = allRows.slice(1, 4).map((row) => row.filter(Boolean).join(" | ")).filter(Boolean);
    if (entries.length) {
      notes.push(`Cap table detected in sheet ${sheetName}: ${entries.join(" ; ")}`);
    }
  }

  if (header.some((item) => /roadmap|milestone|timeline|date/.test(item))) {
    const entries = allRows.slice(1, 5).map((row) => row.filter(Boolean).join(" | ")).filter(Boolean);
    if (entries.length) {
      notes.push(`Roadmap or milestone table detected in sheet ${sheetName}: ${entries.join(" ; ")}`);
    }
  }

  if (header.some((item) => /unit economics|cost|price|margin|payback|revenue/.test(item))) {
    const entries = allRows.slice(1, 5).map((row) => row.filter(Boolean).join(" | ")).filter(Boolean);
    if (entries.length) {
      notes.push(`Unit economics table detected in sheet ${sheetName}: ${entries.join(" ; ")}`);
    }
  }

  return notes;
}

function collectTakeawayLines(structured) {
  const sectionBullets = structured.sections.flatMap((section) => section.bullets || []);
  return [...sectionBullets, ...structured.dataPoints].filter(Boolean).slice(0, 3);
}

function deriveOpenQuestions(structured) {
  const questions = structured.actionItems.filter((item) => /\?/.test(item));
  if (questions.length) {
    return questions.slice(0, 4);
  }
  return [];
}

function deriveFundraisingNotes(structured) {
  return dedupeLines([
    ...(structured.fundraisingNotes || []),
    ...(structured.materialInsights?.fundraising || []),
    ...structured.sections.flatMap((section) =>
      (section.bullets || []).filter((item) => looksLikeFundingLine(item)),
    ),
  ])
    .map(cleanTranscriptArtifact)
    .filter((item) => item && item !== "N/A")
    .filter((item) => item.length <= 240)
    .filter((item) => !looksLikeMostlyTranscriptGarbage(item))
    .slice(0, 10);
}

function deriveUncategorizedLines(structured, usedSectionTitles = []) {
  const uncategorizedSectionBullets = (structured.sections || [])
    .filter((section) => !usedSectionTitles.includes(section.title))
    .flatMap((section) => (section.bullets || []).map((bullet) => `${section.title}: ${bullet}`));

  const categorized = new Set(
    (structured.sections || [])
      .filter((section) => usedSectionTitles.includes(section.title))
      .flatMap((section) => section.bullets || [])
      .map((item) => cleanGeneratedText(item).toLowerCase()),
  );

  return dedupeLines([
    ...(structured.uncategorized || []),
    ...uncategorizedSectionBullets,
  ])
    .map(cleanTranscriptArtifact)
    .filter(Boolean)
    .filter((item) => !categorized.has(cleanGeneratedText(item).toLowerCase()))
    .slice(0, 20);
}

function deriveFullSectionCapture(structured) {
  const blocks = [];

  (structured.sections || []).forEach((section) => {
    const bullets = (section.bullets || [])
      .map(cleanTranscriptArtifact)
      .filter(Boolean);

    if (!section.title || bullets.length === 0) {
      return;
    }

    blocks.push(`### ${cleanTranscriptArtifact(section.title)}`);
    bullets.forEach((item) => blocks.push(`- ${item}`));
    blocks.push("");
  });

  return blocks.filter((item, index, array) => !(item === "" && array[index - 1] === ""));
}

function buildInterviewSourceNotes(structured) {
  const notes = [];
  const userMaterials = structured.meeting.meta.materials?.map((item) => item.originalName).filter(Boolean) || [];
  const provenanceSummary = buildProvenanceSummary(structured.provenance);
  const enhancementHistory = structured.provenance?.enhancementHistory || [];
  notes.push(`**User-provided materials:** ${userMaterials.join(", ") || "None beyond transcript"}`);
  notes.push(
    `**Public-source enrichment:** ${
      structured.research?.searches?.length
        ? structured.research.searches.map((item) => item.query).join("; ")
        : "Not used"
    }`,
  );
  notes.push(`**Provenance summary:** ${provenanceSummary}`);
  notes.push(
    `**Enhancement history:** ${
      enhancementHistory.length
        ? enhancementHistory
            .map((item) => `${item.timestamp.slice(0, 10)} | ${item.fields.join(", ")} | ${item.provider}`)
            .join("; ")
        : "No public-source enhancement applied yet"
    }`,
  );
  notes.push("**Inference / synthesis areas:** Analyst summary, section labeling, prioritization, and risk framing.");
  return notes;
}

function stripMarkdownStrong(value) {
  return String(value || "").replace(/\*\*/g, "").trim();
}

function stripWrappingQuotes(value) {
  return String(value || "").replace(/^["“”']+|["“”']+$/g, "");
}

function cleanTranscriptArtifact(value) {
  return String(value || "")
    .replace(
      /^\s*(?:Speaker|SPEAKER|说话人|發言人|讲话人|璇磋瘽浜[^\s:：-]*)\s*\d*\s*(\[[^\]]+\])?\s*[:：-]?\s*/i,
      "",
    )
    .replace(/\s*\[\d{1,2}:\d{2}(?::\d{2})?\]\s*/g, " ")
    .replace(/^\s*Transcript\s*[:：]\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

function deriveFundraisingRows(structured) {
  const explicitRows = (structured.fundingTable || [])
    .map((row) => [
      row.round || "N/A",
      row.raised || "N/A",
      row.valuation || "N/A",
      row.keyShareholders || "N/A",
    ])
    .filter((row) => row.some((item) => item !== "N/A"));

  if (explicitRows.length) {
    return explicitRows.slice(0, 6);
  }

  const insightRows = (structured.materialInsights?.fundraising || [])
    .slice(0, 3)
    .map((line, index) => toFundraisingRow(line, index))
    .filter(Boolean);

  if (insightRows.length) {
    return insightRows;
  }

  const lines = [
    ...structured.sections.flatMap((section) => section.bullets || []),
    ...structured.dataPoints,
    ...structured.quotes,
    ...splitIntoMemoSentences(structured.meeting?.transcript || ""),
  ].filter(Boolean);

  const matched = lines
    .filter((line) => looksLikeFundingLine(line))
    .filter((line) => line.length <= 180)
    .filter((line) => !looksLikeMostlyTranscriptGarbage(line))
    .slice(0, 8);

  if (matched.length === 0) {
    return [["N/A", "N/A", "N/A", "N/A"]];
  }

  if (matched.length > 1) {
    const combinedRow = toAggregatedFundraisingRow(matched, 0);
    if (combinedRow) {
      return [combinedRow];
    }
  }

  return matched
    .map((line, index) => toFundraisingRow(line, index))
    .filter(Boolean)
    .slice(0, 3);
}

function looksLikeFundingLine(value) {
  return /\b(round|series|seed|pre-ipo|ipo|funding|fundraising|financing|raised|raise|valuation|post-money|investor|investors|shareholder|shareholders|cornerstone)\b/i.test(
    String(value || ""),
  );
}

function toFundraisingRow(line, index) {
  const normalized = cleanTranscriptArtifact(line);
  if (!normalized || isUnsafeFundingCell(normalized)) {
    return null;
  }
  const row = [
    inferRoundLabel(normalized, index),
    extractMoneyValue(normalized) || "N/A",
    extractValuationValue(normalized) || "N/A",
    extractInvestorHint(normalized) || "N/A",
  ];

  const useful = row.some((item, idx) => (idx === 0 ? /seed|angel|series|pre-ipo|ipo|round/i.test(item) : item !== "N/A"));
  return useful ? row : null;
}

function toAggregatedFundraisingRow(lines, index) {
  const normalized = lines.map(cleanTranscriptArtifact).filter(Boolean).join(" ");
  if (!normalized) {
    return null;
  }
  const row = [
    inferRoundLabel(normalized, index),
    extractMoneyValue(normalized) || "N/A",
    extractValuationValue(normalized) || "N/A",
    extractInvestorHint(normalized) || "N/A",
  ];
  const useful = row.some((item, idx) => (idx === 0 ? /seed|angel|series|pre-ipo|ipo|round/i.test(item) : item !== "N/A"));
  return useful ? row : null;
}

function inferRoundLabel(line, index) {
  const match = line.match(/(seed|angel|series\s*[a-d]|pre-ipo|ipo)/i);
  return match?.[1]?.replace(/\s+/g, " ").toUpperCase() || `Round ${index + 1}`;
}

function extractMoneyValue(line) {
  const raisedIndex = String(line || "").toLowerCase().indexOf("raised");
  if (raisedIndex >= 0) {
    const raisedText = String(line).slice(raisedIndex);
    const raisedAmount = raisedText.match(/((?:US\$|\$|RMB)\s*[\d\.,]+\s*(?:m|mn|million|b|bn|billion)?)/i);
    if (raisedAmount?.[1]) {
      return raisedAmount[1].trim();
    }
  }
  const explicitRaised = line.match(/\braised?\s*[:\-]?\s*((?:US\$|\$|RMB)\s*[\d\.,]+\s*(?:m|mn|million|b|bn|billion)?|[\d\.,]+\s*(?:亿美元|亿元|million|billion))/i);
  if (explicitRaised?.[1]) {
    return explicitRaised[1].trim();
  }
  const explicit = line.match(/(?:funding|financing|issue size|融资|募资)\s*[:\-]?\s*([^,;|]+)/i);
  if (explicit?.[1] && /[\d$¥￥]|US\$|RMB/i.test(explicit[1])) {
    const explicitAmount = explicit[1].match(/((?:US\$|\$|RMB)\s*[\d\.,]+\s*(?:m|mn|million|b|bn|billion)?|[\d\.,]+\s*(?:亿美元|亿元|million|billion))/i);
    return explicitAmount?.[1]?.trim() || "";
  }
  const generic = line.match(/((?:US\$|\$|RMB)\s*[\d\.,]+\s*(?:m|mn|million|b|bn|billion)?|[\d\.,]+\s*(?:亿美元|亿元|million|billion))/i);
  return generic?.[1]?.trim() || "";
}

function extractValuationValue(line) {
  const explicit = line.match(/\bvaluation\b\s*[:\-]?\s*([^,;|]+)/i);
  if (explicit?.[1]) {
    return explicit[1]
      .trim()
      .replace(/\s*(?:key shareholders?|investors?|shareholders?)\b.*$/i, "")
      .replace(/[。.!?]\s*$/, "")
      .trim();
  }
  const generic = line.match(/(?:pre-money|post-money|估值)\s*[:\-]?\s*([^,;|]+)/i);
  return generic?.[1]?.trim() || "";
}

function extractInvestorHint(line) {
  const explicit = line.match(/(?:investors?|shareholders?|backed by|key shareholders?|cornerstone)\s*[:\-]?\s*([^;|]+)/i);
  if (explicit?.[1]) {
    return explicit[1]
      .trim()
      .replace(/\s*(?:valuation|raised?|funding|financing)\b.*$/i, "")
      .replace(/[。.!?]\s*$/, "")
      .trim();
  }
  const named = line.match(/(Alibaba|Cainiao|Hyundai|Temasek|SoftBank|Sequoia|腾讯|阿里|菜鸟|国投|上汽|一汽)/i);
  return named?.[1]?.trim() || "";
}

async function readTemplate(templateId) {
  const filePath = path.join(TEMPLATES_DIR, `${templateId}.json`);

  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (_error) {
    throw createError(404, `Template "${templateId}" was not found.`);
  }
}

function slugify(value) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function sanitizeText(value, fallback = "") {
  if (typeof value !== "string") {
    return fallback;
  }

  const cleaned = value.trim();
  return cleaned || fallback;
}

function sanitizeFileName(value) {
  const safe = String(value || "file")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "") || "file";
  const extension = path.extname(safe).slice(0, 12);
  const base = path.basename(safe, extension).slice(0, 70).replace(/[._-]+$/g, "") || "file";
  return `${base}${extension}`;
}

function splitParticipants(value) {
  if (typeof value !== "string" || !value.trim()) {
    return [];
  }

  return value
    .split(/[,，\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function firstMeaningfulSentence(text) {
  return text
    .split(/(?<=[.!?。！？])\s+/)
    .map((sentence) => sentence.trim())
    .find((sentence) => sentence.length > 15);
}

function createError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

module.exports = {
  createDraftJob,
  ensureAppFolders,
  exportJobToObsidian,
  deriveOutputFromJob,
  finalizeJob,
  getArtifactPath,
  getRuntimeProviders,
  listJobs,
  processTranscriptInput,
  processMeetingInput,
  readJob,
};
