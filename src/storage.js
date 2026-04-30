const path = require("path");
const fs = require("fs/promises");

const ROOT = path.join(__dirname, "..");
const DATA_DIR = path.join(ROOT, "data");
const TRANSCRIPTS_DIR = path.join(DATA_DIR, "transcripts");
const STRUCTURED_DIR = path.join(DATA_DIR, "structured");
const OUTPUTS_DIR = path.join(DATA_DIR, "outputs");
const INPUT_DIR = path.join(DATA_DIR, "input");
const MATERIALS_DIR = path.join(DATA_DIR, "materials");

function createLocalStorageAdapter() {
  return {
    mode: "local",
    rootDir: DATA_DIR,
    ensureFolders,
    async listStructuredRecords() {
      const files = await fs.readdir(STRUCTURED_DIR);
      return files.filter((file) => file.endsWith(".json"));
    },
    async readStructuredRecord(fileName) {
      return fs.readFile(path.join(STRUCTURED_DIR, fileName), "utf8");
    },
    async writeStructuredRecord(fileName, content) {
      const filePath = path.join(STRUCTURED_DIR, fileName);
      await fs.writeFile(filePath, content, "utf8");
      return filePath;
    },
    async writeTranscript(fileName, content) {
      const filePath = path.join(TRANSCRIPTS_DIR, fileName);
      await fs.writeFile(filePath, content, "utf8");
      return filePath;
    },
    async writeInputBinary(fileName, buffer) {
      const filePath = path.join(INPUT_DIR, fileName);
      await fs.writeFile(filePath, buffer);
      return filePath;
    },
    async writeMaterialBinary(fileName, buffer) {
      const filePath = path.join(MATERIALS_DIR, fileName);
      await fs.writeFile(filePath, buffer);
      return filePath;
    },
    async writeOutputText(fileName, content) {
      const filePath = path.join(OUTPUTS_DIR, fileName);
      await fs.writeFile(filePath, content, "utf8");
      return filePath;
    },
    async writeOutputJson(fileName, payload) {
      const filePath = path.join(OUTPUTS_DIR, fileName);
      await fs.writeFile(filePath, JSON.stringify(payload, null, 2), "utf8");
      return filePath;
    },
    getStructuredPath(fileName) {
      return path.join(STRUCTURED_DIR, fileName);
    },
    getTranscriptPath(fileName) {
      return path.join(TRANSCRIPTS_DIR, fileName);
    },
    getOutputPath(fileName) {
      return path.join(OUTPUTS_DIR, fileName);
    },
    getMaterialPath(fileName) {
      return path.join(MATERIALS_DIR, fileName);
    },
  };
}

async function ensureFolders() {
  await Promise.all([
    fs.mkdir(INPUT_DIR, { recursive: true }),
    fs.mkdir(TRANSCRIPTS_DIR, { recursive: true }),
    fs.mkdir(STRUCTURED_DIR, { recursive: true }),
    fs.mkdir(OUTPUTS_DIR, { recursive: true }),
    fs.mkdir(MATERIALS_DIR, { recursive: true }),
  ]);
}

function getStorageAdapter() {
  const requestedMode = String(process.env.STORAGE_MODE || "local").trim().toLowerCase();
  return createLocalStorageAdapter(requestedMode);
}

module.exports = {
  DATA_DIR,
  INPUT_DIR,
  MATERIALS_DIR,
  OUTPUTS_DIR,
  STRUCTURED_DIR,
  TRANSCRIPTS_DIR,
  ensureFolders,
  getStorageAdapter,
};
