const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs/promises");
const fsSync = require("fs");
const {
  ensureAppFolders,
  deriveOutputFromJob,
  exportJobToObsidian,
  finalizeJob,
  getArtifactPath,
  getRuntimeProviders,
  listJobs,
  processMeetingInput,
  processTranscriptInput,
  readJob,
} = require("./pipeline");

const app = express();
const port = process.env.PORT || 3100;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 300 * 1024 * 1024,
  },
});

app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "..", "public")));

app.get("/api/health", async (_req, res) => {
  await ensureAppFolders();
  res.json({ ok: true });
});

app.get("/api/templates", async (_req, res, next) => {
  try {
    const templatesDir = path.join(__dirname, "..", "templates");
    const files = await fs.readdir(templatesDir);
    const templates = await Promise.all(
      files
        .filter((file) => file.endsWith(".json"))
        .map(async (file) => {
          const raw = await fs.readFile(path.join(templatesDir, file), "utf8");
          return JSON.parse(raw);
        }),
    );
    res.json({ templates });
  } catch (error) {
    next(error);
  }
});

app.get("/api/providers", (_req, res) => {
  res.json(getRuntimeProviders());
});

app.get("/api/jobs", async (_req, res, next) => {
  try {
    const jobs = await listJobs();
    res.json({ jobs });
  } catch (error) {
    next(error);
  }
});

app.post("/api/jobs/:id/finalize", async (req, res, next) => {
  try {
    const result = await finalizeJob(req.params.id, {
      selectedEnrichmentFields: Array.isArray(req.body.selectedEnrichmentFields)
        ? req.body.selectedEnrichmentFields
        : [],
      llmProvider: req.body.llmProvider,
      researchProvider: req.body.researchProvider || "none",
    });

    res.status(200).json(result);
  } catch (error) {
    next(error);
  }
});

app.get("/api/jobs/:id/artifacts/:type", async (req, res, next) => {
  try {
    const job = await readJob(req.params.id);
    const artifactPath = getArtifactPath(job, req.params.type);

    if (!artifactPath || !fsSync.existsSync(artifactPath)) {
      res.status(404).json({ error: `${req.params.type} artifact not found.` });
      return;
    }

    res.download(artifactPath);
  } catch (error) {
    next(error);
  }
});

app.post("/api/jobs/:id/export/obsidian", async (req, res, next) => {
  try {
    const exportResult = await exportJobToObsidian(req.params.id);
    const job = await readJob(req.params.id);
    res.status(200).json({
      ok: true,
      export: exportResult,
      job,
    });
  } catch (error) {
    next(error);
  }
});

app.post(
  "/api/jobs/:id/derive",
  upload.any(),
  async (req, res, next) => {
    try {
      const files = Array.isArray(req.files) ? req.files : [];
      const materialFiles = filesByField(files, [
        "materials",
        "materials[]",
        "materialFiles",
        "materialFiles[]",
        "file",
        "files",
        "attachment",
        "attachments",
      ]);
      const result = await deriveOutputFromJob(req.params.id, {
        targetTemplateId: req.body.targetTemplateId || "weekly-report",
        materialFiles,
        notesText: req.body.notesText || "",
        userInstructions: req.body.userInstructions || "",
        llmProvider: req.body.llmProvider,
        researchProvider: req.body.researchProvider || "none",
      });

      res.status(201).json(result);
    } catch (error) {
      next(error);
    }
  },
);

app.post(
  "/api/transcript",
  upload.any(),
  async (req, res, next) => {
    try {
      const files = Array.isArray(req.files) ? req.files : [];
      const audioFile = firstFileByField(files, ["audio", "audioFile", "recording"]);
      const transcriptFile = firstFileByField(files, ["transcriptFile", "transcriptUpload", "uploadedTranscript"]);
      const result = await processTranscriptInput({
        meetingTitle: req.body.meetingTitle,
        company: req.body.company,
        asrProvider: req.body.asrProvider,
        userInstructions: req.body.userInstructions,
        transcriptText: req.body.transcriptText,
        audioFile,
        transcriptFile,
      });

      res.status(201).json(result);
    } catch (error) {
      next(error);
    }
  },
);

app.post(
  "/api/process",
  upload.any(),
  async (req, res, next) => {
  try {
    const files = Array.isArray(req.files) ? req.files : [];
    const audioFile = firstFileByField(files, ["audio", "audioFile", "recording"]);
    const transcriptFile = firstFileByField(files, ["transcriptFile", "transcriptUpload", "uploadedTranscript"]);
    const materialFiles = filesByField(files, [
      "materials",
      "materials[]",
      "materialFiles",
      "materialFiles[]",
      "file",
      "files",
      "attachment",
      "attachments",
    ]);
    const payload = {
      meetingTitle: req.body.meetingTitle,
      company: req.body.company,
      meetingType: req.body.meetingType,
      templateId: req.body.templateId,
      asrProvider: req.body.asrProvider,
      llmProvider: req.body.llmProvider,
      researchProvider: req.body.researchProvider || "none",
      participants: req.body.participants,
      transcriptText: req.body.transcriptText,
      notesText: req.body.notesText,
      userInstructions: req.body.userInstructions,
      audioFile,
      transcriptFile,
      materialFiles,
    };

    const result = await processMeetingInput(payload);
    res.status(201).json(result);
  } catch (error) {
    next(error);
  }
});

app.use((error, _req, res, _next) => {
  if (error?.code === "LIMIT_FILE_SIZE") {
    res.status(400).json({
      error: "Uploaded file is too large. Please keep a single file under 300 MB, or split very long audio before upload.",
    });
    return;
  }

  const status = error.statusCode || 500;
  res.status(status).json({
    error: error.message || "Unexpected error",
  });
});

ensureAppFolders()
  .then(() => {
    app.listen(port, () => {
      console.log(`Meeting automation MVP running on http://localhost:${port}`);
    });
  })
  .catch((error) => {
    console.error("Failed to initialize app folders", error);
    process.exit(1);
  });

function firstFileByField(files, fieldNames) {
  return files.find((file) => fieldNames.includes(file.fieldname));
}

function filesByField(files, fieldNames) {
  return files.filter((file) => fieldNames.includes(file.fieldname));
}
