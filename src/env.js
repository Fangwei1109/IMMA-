const fs = require("fs");
const path = require("path");

function loadEnvFile(filePath = path.join(__dirname, "..", ".env")) {
  if (!fs.existsSync(filePath)) {
    return;
  }

  const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }

    const separator = line.indexOf("=");
    if (separator < 1) {
      continue;
    }

    const name = line.slice(0, separator).trim();
    const value = line.slice(separator + 1).trim().replace(/^["']|["']$/g, "");
    if (!name || value.length === 0 || process.env[name]) {
      continue;
    }

    process.env[name] = value;
  }
}

module.exports = {
  loadEnvFile,
};
