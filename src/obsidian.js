const path = require("path");
const fs = require("fs");
const fsp = require("fs/promises");

function getObsidianConfig() {
  const obsidianConfigPath = path.join(process.env.APPDATA || "", "Obsidian", "obsidian.json");

  try {
    const raw = fs.readFileSync(obsidianConfigPath, "utf8");
    const config = JSON.parse(raw);
    const vaultEntries = Object.values(config.vaults || {});
    const activeVault = vaultEntries.find((item) => item.open) || vaultEntries[0];

    return {
      configured: Boolean(activeVault?.path),
      vaultPath: activeVault?.path || "",
    };
  } catch (_error) {
    return {
      configured: false,
      vaultPath: "",
    };
  }
}

async function buildTerminologyHints({ company = "", title = "", userInstructions = "", limit = 20 } = {}) {
  const obsidian = getObsidianConfig();
  if (!obsidian.configured || !obsidian.vaultPath) {
    return [];
  }

  const queryTokens = buildQueryTokens([company, title, userInstructions]);
  const markdownFiles = await collectMarkdownFiles(obsidian.vaultPath, 120);
  const rankedFiles = rankFiles(markdownFiles, queryTokens).slice(0, 14);
  const selectedFiles = rankedFiles.length ? rankedFiles : markdownFiles.slice(0, 8).map((filePath) => ({ filePath, score: 0 }));
  const termScores = new Map();

  for (const { filePath, score } of selectedFiles) {
    const raw = await safeReadFile(filePath);
    if (!raw) {
      continue;
    }

    const head = raw.slice(0, 18000);
    const candidates = extractTermCandidates(head, path.basename(filePath, ".md"));
    for (const candidate of candidates) {
      const normalized = candidate.trim();
      if (!normalized || normalized.length < 2 || normalized.length > 48) {
        continue;
      }
      const nextScore = (termScores.get(normalized) || 0) + 1 + score;
      termScores.set(normalized, nextScore);
    }
  }

  const rankedTerms = [...termScores.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([term]) => term)
    .filter((term) => !looksTooGeneric(term))
    .slice(0, limit * 2);

  return dedupe([
    ...queryTokens,
    ...rankedTerms,
  ]).slice(0, limit);
}

function buildQueryTokens(values) {
  return values
    .flatMap((value) => String(value || "").split(/[^A-Za-z0-9\u4e00-\u9fff]+/))
    .map((token) => token.trim())
    .filter((token) => token.length >= 2)
    .filter((token) => !looksTooGeneric(token))
    .slice(0, 20);
}

async function collectMarkdownFiles(rootDir, limit = 120) {
  const queue = [rootDir];
  const files = [];

  while (queue.length && files.length < limit) {
    const current = queue.shift();
    let entries = [];
    try {
      entries = await fsp.readdir(current, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      const fullPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        queue.push(fullPath);
      } else if (entry.isFile() && entry.name.toLowerCase().endsWith(".md")) {
        files.push(fullPath);
        if (files.length >= limit) {
          break;
        }
      }
    }
  }

  return files;
}

function rankFiles(filePaths, queryTokens) {
  return filePaths
    .map((filePath) => {
      const lower = filePath.toLowerCase();
      const score = queryTokens.reduce((acc, token) => acc + (lower.includes(token.toLowerCase()) ? 3 : 0), 0);
      return { filePath, score };
    })
    .sort((a, b) => b.score - a.score);
}

function extractTermCandidates(content, fileStem) {
  const fileStemTokens = fileStem.split(/[^A-Za-z0-9\u4e00-\u9fff]+/).filter(Boolean);
  const headings = [...content.matchAll(/^\s{0,3}#{1,3}\s+(.+)$/gm)].map((match) => match[1].trim());
  const englishTerms = [...content.matchAll(/\b[A-Z][A-Za-z0-9&+./-]{2,}\b/g)].map((match) => match[0]);
  const allCapsTerms = [...content.matchAll(/\b[A-Z]{2,}(?:[-_][A-Z0-9]{2,})*\b/g)].map((match) => match[0]);
  const chineseTerms = [...content.matchAll(/[\u4e00-\u9fff]{2,12}/g)].map((match) => match[0]);

  return dedupe([
    ...fileStemTokens,
    ...headings,
    ...englishTerms,
    ...allCapsTerms,
    ...chineseTerms,
  ]);
}

function looksTooGeneric(value) {
  const text = String(value || "").trim();
  if (!text) {
    return true;
  }

  return /^(meeting|memo|report|notes|weekly|daily|general|update|summary|interview|automation|oneNote|untitled|company|business|technology|product|what|next|其中|次迭代|exported)$/i.test(text);
}

async function safeReadFile(filePath) {
  try {
    return await fsp.readFile(filePath, "utf8");
  } catch {
    return "";
  }
}

function dedupe(items) {
  return [...new Set(items.filter(Boolean))];
}

module.exports = {
  buildTerminologyHints,
  getObsidianConfig,
};
