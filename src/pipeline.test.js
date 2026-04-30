const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs/promises");
const fsSync = require("fs");
const path = require("path");
const os = require("os");
const { processMeetingInput, processTranscriptInput, deriveOutputFromJob, finalizeJob, exportJobToObsidian, getArtifactPath, getRuntimeProviders } = require("./pipeline");

const sampleTranscript = `
We met the management team to discuss quarterly demand trends and product cadence.
Management said order momentum remained healthy, although pricing remained competitive.
The team expects margin pressure to ease in the second half as utilization improves.
We need follow-up on channel inventory, export sustainability, and the timing of the next launch.
`;

test("generates interview memo markdown output", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Management diligence",
    company: "SampleCo",
    meetingType: "management-interview",
    participants: "CEO, CFO",
    templateId: "interview-knowledge-base",
    asrProvider: "auto",
    transcriptText: sampleTranscript,
  });

  assert.equal(result.job.templateId, "interview-knowledge-base");
  assert.equal(result.job.asrProvider, "manual");
  assert.equal(result.job.sourceType, "provided-transcript");
  assert.match(result.markdown, /# SampleCo \| Interview Memo/);
  assert.match(result.markdown, /## Investment Take/);
  assert.match(result.markdown, /## Open Questions \/ Follow-ups/);
  assert.doesNotMatch(result.markdown, /## Full Section Capture/);
  assert.doesNotMatch(result.markdown, /## Key Evidence \/ Data Points/);
  await assertFileExists(result.job.markdownPath);
  await assertFileExists(getArtifactPath(result, "transcript"));
});

test("uses uploaded transcript file as transcript source", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Uploaded transcript",
    company: "UploadCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    asrProvider: "auto",
    llmProvider: "none",
    transcriptFile: {
      originalname: "call-transcript.txt",
      mimetype: "text/plain",
      buffer: Buffer.from(
        "Speaker 1 [00:01:02]: Management said revenue comes from annual software subscriptions and two OEM pilot customers are active.",
        "utf8",
      ),
    },
    materialFiles: [
      {
        originalname: "brief.txt",
        mimetype: "text/plain",
        buffer: Buffer.from("Supporting note: expand to Singapore next year.", "utf8"),
      },
    ],
  });

  assert.equal(result.job.sourceType, "uploaded-transcript");
  assert.equal(result.job.asrProvider, "manual");
  assert.equal(result.structured.meeting.meta.materials.length, 1);
  assert.doesNotMatch(result.inputs.transcriptText, /Speaker 1|\[00:01:02\]/);
  assert.match(result.inputs.transcriptText, /annual software subscriptions/i);
});

test("creates clean transcript before memo generation", async () => {
  const result = await processTranscriptInput({
    meetingTitle: "Transcript stage",
    company: "StageCo",
    asrProvider: "auto",
    transcriptFile: {
      originalname: "stage-transcript.txt",
      mimetype: "text/plain",
      buffer: Buffer.from(
        "Speaker 2 [00:02:10]: Management said the world model architecture is the core technical breakthrough and customer pilots are ongoing.",
        "utf8",
      ),
    },
  });

  assert.equal(result.sourceType, "uploaded-transcript");
  assert.equal(result.asrProvider, "manual");
  assert.doesNotMatch(result.transcriptText, /Speaker 2|\[00:02:10\]/);
  assert.match(result.transcriptText, /world model architecture/i);
  await assertFileExists(result.transcriptPath);
});

test("cleans speaker labels from provided transcript in interview memo", async () => {
  const noisyTranscript = `
Speaker 1: We expect a pilot launch in June and a broader rollout in Q4.
说话人2：The financing target is around US$30m and management is discussing cornerstone investors.
[00:12:32] The team said product readiness improved after the latest validation cycle.
`;

  const result = await processMeetingInput({
    meetingTitle: "Noisy transcript",
    company: "CleanCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    asrProvider: "auto",
    llmProvider: "none",
    transcriptText: noisyTranscript,
  });

  assert.doesNotMatch(result.markdown, /Speaker 1/i);
  assert.doesNotMatch(result.markdown, /说话人2|说话人 2/i);
  const cleanedTranscript = await fs.readFile(result.job.transcriptPath, "utf8");
  assert.doesNotMatch(cleanedTranscript, /Speaker 1/i);
  assert.doesNotMatch(cleanedTranscript, /\[00:12:32\]/);
});

test("keeps fundraising snapshot structured instead of dumping transcript chunks", async () => {
  const noisyFundingTranscript = `
Speaker 1: The company completed a Series A financing last year.
Speaker 2: Raised: US$30m. Valuation: US$200m post-money. Key shareholders: Alibaba, Hyundai.
[00:12:32] Management said the next step is a Pre-IPO process.
`;

  const result = await processMeetingInput({
    meetingTitle: "Funding cleanup test",
    company: "FundingCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    asrProvider: "auto",
    llmProvider: "none",
    transcriptText: noisyFundingTranscript,
  });

  assert.doesNotMatch(result.markdown, /\|[^|\n]*Speaker/i);
  assert.doesNotMatch(result.markdown, /\|[^|\n]*00:12:32/i);
});

test("generates weekly report markdown output", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Weekly portfolio sync",
    company: "SampleCo",
    meetingType: "weekly-sync",
    participants: "PM, Analyst",
    templateId: "weekly-report",
    asrProvider: "whisper",
    transcriptText: sampleTranscript,
  });

  assert.equal(result.job.templateId, "weekly-report");
  assert.equal(result.job.asrProvider, "manual");
  assert.match(result.markdown, /## Core Updates/);
  assert.match(result.markdown, /## Next Steps/);
  await assertFileExists(result.job.markdownPath);
});

test("finalizes interview memo without generating ppt", async () => {
  const draft = await processMeetingInput({
    meetingTitle: "Management diligence",
    company: "SampleCo",
    meetingType: "management-interview",
    participants: "CEO, CFO",
    templateId: "interview-knowledge-base",
    asrProvider: "auto",
    researchProvider: "none",
    transcriptText: sampleTranscript,
  });

  const finalized = await finalizeJob(draft.job.id, {
    selectedEnrichmentFields: ["team", "funding"],
    llmProvider: "none",
    researchProvider: "none",
  });

  assert.equal(finalized.job.templateId, "interview-knowledge-base");
  assert.equal(finalized.job.pptPath, null);
  assert.match(finalized.markdown, /## Open Questions \/ Follow-ups/);
  assert.match(finalized.markdown, /## Analyst View/);
});

test("generates free-style memo without forcing fixed company-wide buckets", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Tech deep dive",
    company: "FocusCo",
    meetingType: "technical-diligence",
    templateId: "interview-free-style",
    llmProvider: "none",
    transcriptText:
      "The interview focused on sensor fusion, perception stack, and deployment architecture. Management spent most of the call on technical roadmap, model training loop, and hardware constraints. Follow-up is needed on commercialization timing and validation metrics.",
  });

  assert.equal(result.job.templateId, "interview-free-style");
  assert.match(result.markdown, /# FocusCo \| Free Style Memo/);
  assert.match(result.markdown, /### Technical \/ Product Discussion|### Primary Discussion/);
  assert.doesNotMatch(result.markdown, /## Business \/ Strategy/);
  assert.doesNotMatch(result.markdown, /## Team/);
  assert.match(result.markdown, /\*\*Detected Focus:\*\* Tech Deep Dive/);
  assert.match(result.markdown, /<details>\s*<summary>Source Notes<\/summary>/);
  assert.equal(result.review.recommendation.recommendedTemplateId, "interview-free-style");
});

test("recommends interview memo for broader company coverage", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Broad diligence",
    company: "BreadthCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    transcriptText:
      "The founder discussed company history, management hiring, product roadmap, customer pipeline, margin profile, fundraising history, and IPO planning. The team also covered business model, manufacturing readiness, and geographic expansion.",
  });

  assert.equal(result.review.recommendation.recommendedTemplateId, "interview-knowledge-base");
  assert.match(result.markdown, /<details>\s*<summary>Source Notes<\/summary>/);
});

test("interview memo preserves data-focused technical subtopics and avoids wrong buckets", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Data stack interview",
    company: "DataStackCo",
    meetingType: "technical-diligence",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    transcriptText:
      "Management said the core technical discussion was about ego-centric data strategy and data collection devices. The team described sensor rigs, in-vehicle collection hardware, labeling infrastructure, simulation data, and model evaluation tooling. The go-to-market plan is to start with two OEM pilot customers. Revenue model will be software subscription plus data service fees. We need follow-up on data ownership and validation metrics.",
  });

  assert.match(result.markdown, /## Product \/ Technology/);
  assert.match(result.markdown, /ego-centric data strategy|data collection devices|sensor rigs/i);
  assert.match(result.markdown, /> \[!QUOTE\] "/);
  assert.match(result.markdown, /## Business \/ Strategy[\s\S]*go-to-market/i);
  assert.match(result.markdown, /## Commercial \/ Financial Signals[\s\S]*Revenue model/i);
  assert.doesNotMatch(result.markdown, /## Product \/ Technology[\s\S]*Revenue model will be software subscription plus data service fees[\s\S]*## Commercial \/ Financial Signals/);
});

test("interview memo routes quotes into relevant tech subtopics", async () => {
  const result = await processMeetingInput({
    meetingTitle: "World model interview",
    company: "WorldModelCo",
    meetingType: "technical-diligence",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    transcriptText:
      "Management said the world model architecture is the core technical breakthrough and it connects perception, planning, and simulation. The team also said the ego-centric data engine collects in-vehicle sensor data, labeling outputs, and model evaluation results. Revenue model will be software subscription fees. The GTM plan is OEM pilots.",
  });

  assert.match(result.markdown, /World Model/);
  assert.match(result.markdown, /Data Strategy|Data Collection/);
  assert.match(result.markdown, /World Model[\s\S]*> \[!QUOTE\] ".*world model architecture/i);
  assert.doesNotMatch(result.markdown, /Other \/ Unclassified Themes[\s\S]*> \[!QUOTE\]/);
  assert.match(result.markdown, /Commercial \/ Financial Signals[\s\S]*Revenue model/i);
});

test("interview memo excludes interviewer HMG intro from company evidence", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Role aware interview",
    company: "RoleCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    transcriptText:
      "We are Hyundai Motor Group and Cradle, and we invest in mobility startups from our side. Can you explain your revenue model? Management said the company charges annual software subscription fees and has two OEM pilot customers.",
  });

  assert.match(result.markdown, /annual software subscription fees/i);
  assert.match(result.markdown, /OEM pilot customers/i);
  assert.doesNotMatch(result.markdown, /Hyundai Motor Group|Cradle|we invest in mobility startups/i);
  assert.ok(result.structured.roleAnalysis.excludedContext.some((line) => /Hyundai Motor Group|Cradle/i.test(line)));
  await assertFileExists(getArtifactPath(result, "role-transcript"));
  await assertFileExists(getArtifactPath(result, "evidence-bank"));
});

test("derives weekly report ppt from an interview memo plus additional material", async () => {
  const memo = await processMeetingInput({
    meetingTitle: "Derived weekly source",
    company: "DerivedCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    transcriptText:
      "Management said the company has two OEM pilot customers and annual software subscription revenue. The product roadmap includes a 2026 launch and model evaluation milestones.",
  });

  const weekly = await deriveOutputFromJob(memo.job.id, {
    targetTemplateId: "weekly-report",
    llmProvider: "none",
    userInstructions: "Emphasize HMG collaboration relevance.",
    materialFiles: [
      {
        originalname: "weekly-brief.txt",
        mimetype: "text/plain",
        buffer: Buffer.from("Additional weekly report material: Series A fundraising, Singapore GTM, and OEM integration plan.", "utf8"),
      },
    ],
  });

  assert.equal(weekly.job.templateId, "weekly-report");
  assert.equal(weekly.job.derivedFromJobId, memo.job.id);
  assert.ok(weekly.job.pptPath);
  assert.ok(weekly.structured.weeklyReportDraft.sections.length >= 1);
  assert.doesNotMatch(JSON.stringify(weekly.structured.weeklyReportDraft), /Evidence Bank|Role-labeled|Source Notes|companyEvidence|interviewerContext/i);
  const productSection = weekly.structured.weeklyReportDraft.sections.find((section) => section.label === "Product / Technology");
  assert.ok(productSection.lines.length > 0);
  assert.doesNotMatch(productSection.lines.join("\n"), /^# Product|^# Technology/im);
  assert.match(weekly.markdown, /Weekly Report|Core Updates/i);
  await assertFileExists(weekly.job.pptPath);
});

test("ingests supporting materials and notes for weekly report generation", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Weekly portfolio sync",
    company: "SampleCo",
    meetingType: "weekly-sync",
    participants: "PM, Analyst",
    templateId: "weekly-report",
    asrProvider: "auto",
    researchProvider: "none",
    notesText: "Founder previously led APAC expansion and recently announced a strategic partnership.",
    materialFiles: [
      {
        originalname: "brief.txt",
        mimetype: "text/plain",
        buffer: Buffer.from("SampleCo plans to expand to Singapore and signed a pilot with a logistics partner.", "utf8"),
      },
    ],
  });

  assert.equal(result.job.researchProvider, "none");
  assert.equal(result.structured.meeting.meta.materials.length, 2);
  assert.match(result.structured.meeting.materialsText, /expand to Singapore/);
  assert.match(result.markdown, /## Core Updates/);
});

test("accepts image materials without configured vision model", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Image-backed diligence",
    company: "SampleCo",
    meetingType: "management-interview",
    participants: "CEO",
    templateId: "interview-knowledge-base",
    asrProvider: "auto",
    llmProvider: "none",
    transcriptText:
      "The company shared a product roadmap image and noted it expects a new launch in 2026 with pilot deployments ongoing.",
    materialFiles: [
      {
        originalname: "roadmap.png",
        mimetype: "image/png",
        buffer: Buffer.from("89504e470d0a1a0a", "hex"),
      },
    ],
  });

  assert.equal(result.structured.meeting.meta.materials[0].extractionMethod, "image-unprocessed");
  assert.match(result.markdown, /Interview Memo/);
});

test("extracts docx supporting material text", async () => {
  const JSZip = require("jszip");
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p><w:r><w:t>Demo roadmap and Series A timeline with commercialization milestones, customer pipeline, product focus, and fundraising preparation for the next financing round.</w:t></w:r></w:p>
      </w:body>
    </w:document>`,
  );

  const result = await processMeetingInput({
    meetingTitle: "Docx input",
    company: "DocxCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    materialFiles: [
      {
        originalname: "memo.docx",
        mimetype: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        buffer: await zip.generateAsync({ type: "nodebuffer" }),
      },
    ],
  });

  assert.equal(result.structured.meeting.meta.materials[0].extractionMethod, "docx-mammoth");
  assert.match(result.structured.meeting.materialsText, /Series A timeline/);
});

test("extracts pptx supporting material text", async () => {
  const JSZip = require("jszip");
  const zip = new JSZip();
  zip.file("ppt/slides/slide1.xml", '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>2026 Roadmap</a:t><a:t>Pilot launch</a:t><a:t>Mass production SOP</a:t></p:sld>');
  zip.file("ppt/slides/slide2.xml", '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>Fundraising plan</a:t><a:t>Pre-IPO</a:t><a:t>Valuation $800m</a:t><a:t>Unit economics</a:t><a:t>Payback 18 months</a:t></p:sld>');

  const result = await processMeetingInput({
    meetingTitle: "Pptx input",
    company: "DeckCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    materialFiles: [
      {
        originalname: "deck.pptx",
        mimetype: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        buffer: await zip.generateAsync({ type: "nodebuffer" }),
      },
    ],
  });

  assert.equal(result.structured.meeting.meta.materials[0].extractionMethod, "pptx-zip");
  assert.match(result.structured.meeting.materialsText, /Slide 1/);
  assert.match(result.structured.meeting.materialsText, /Fundraising plan/);
  assert.match(result.structured.meeting.materialsText, /Roadmap identified on slide 1/);
  assert.match(result.structured.meeting.materialsText, /Unit economics signals on slide 2/);
  assert.ok(result.structured.materialInsights.roadmap.length > 0);
  assert.ok(result.structured.materialInsights.fundraising.length > 0);
  assert.ok(result.structured.materialInsights.unitEconomics.length > 0);
});

test("extracts xlsx supporting material text", async () => {
  const XLSX = require("xlsx");
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([
    ["Round", "Raised", "Valuation"],
    ["Series A", "$10m", "$80m"],
  ]);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Funding");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  const result = await processMeetingInput({
    meetingTitle: "Xlsx input",
    company: "SheetCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    materialFiles: [
      {
        originalname: "funding.xlsx",
        mimetype: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        buffer,
      },
    ],
  });

  assert.equal(result.structured.meeting.meta.materials[0].extractionMethod, "spreadsheet-xlsx");
  assert.match(result.structured.meeting.materialsText, /Series A \| \$10m \| \$80m/);
  assert.match(result.structured.meeting.materialsText, /Fundraising table detected in sheet Funding/);
  assert.ok(result.structured.materialInsights.fundraising.length > 0);
});

test("runtime providers expose research configuration", () => {
  const previousAli = process.env.ALI_API_KEY;
  process.env.ALI_API_KEY = "test-key";
  const providers = getRuntimeProviders();
  assert.ok(providers.research);
  assert.ok(Array.isArray(providers.research.providers));
  assert.ok(providers.llm);
  assert.ok(Array.isArray(providers.llm.providers));
  assert.equal(providers.storage.mode, "local");
  assert.ok(providers.asr.some((provider) => provider.id === "ali" && provider.configured));
  process.env.ALI_API_KEY = previousAli;
});

test("draft review exposes provenance summary and render checks", async () => {
  const result = await processMeetingInput({
    meetingTitle: "Weekly review",
    company: "TraceCo",
    meetingType: "weekly-sync",
    templateId: "weekly-report",
    llmProvider: "none",
    transcriptText:
      "The founder said the product rollout spans three cities. The company is preparing a Series A fundraising and expects to file for IPO readiness later.",
  });

  assert.ok(result.review.provenanceSummary);
  assert.ok(Array.isArray(result.review.warnings));
  assert.ok(result.structured.provenance);
  assert.ok(result.structured.renderChecks);
  assert.equal(result.structured.provenance.fieldSources.product.source, "user-provided");
});

test("finalize appends enhancement history", async () => {
  const draft = await processMeetingInput({
    meetingTitle: "Enhancement trace test",
    company: "TraceCo",
    meetingType: "management-interview",
    templateId: "interview-knowledge-base",
    llmProvider: "none",
    researchProvider: "none",
    transcriptText: sampleTranscript,
  });

  const finalized = await finalizeJob(draft.job.id, {
    selectedEnrichmentFields: ["team", "funding"],
    llmProvider: "none",
    researchProvider: "none",
  });

  assert.ok(finalized.structured.provenance.enhancementHistory.length >= 1);
  assert.deepEqual(finalized.structured.provenance.enhancementHistory.at(-1).fields, ["team", "funding"]);
});

test("exports interview memo markdown to detected obsidian vault", async () => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "obsidian-test-"));
  const appDataDir = path.join(tempRoot, "AppData");
  const obsidianDir = path.join(appDataDir, "Obsidian");
  const vaultDir = path.join(tempRoot, "Vault");
  await fs.mkdir(obsidianDir, { recursive: true });
  await fs.mkdir(vaultDir, { recursive: true });
  await fs.writeFile(
    path.join(obsidianDir, "obsidian.json"),
    JSON.stringify({
      vaults: {
        test: {
          path: vaultDir,
          ts: Date.now(),
          open: true,
        },
      },
    }),
    "utf8",
  );

  const previousAppData = process.env.APPDATA;
  process.env.APPDATA = appDataDir;

  try {
    const draft = await processMeetingInput({
      meetingTitle: "Obsidian export test",
      company: "VaultCo",
      meetingType: "management-interview",
      templateId: "interview-knowledge-base",
      llmProvider: "none",
      transcriptText: sampleTranscript,
    });

    const exported = await exportJobToObsidian(draft.job.id);
    assert.match(exported.path, /Meeting Automation/);
    assert.equal(fsSync.existsSync(exported.path), true);
  } finally {
    process.env.APPDATA = previousAppData;
  }
});

async function assertFileExists(filePath) {
  const resolved = path.resolve(filePath);
  await fs.access(resolved);
}
