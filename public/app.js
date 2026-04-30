const templateCards = Array.from(document.querySelectorAll(".choice-card"));
const form = document.getElementById("process-form");
const transcribeButton = document.getElementById("transcribe-button");
const transcriptStepStatus = document.getElementById("transcript-step-status");
const cleanTranscriptText = document.getElementById("clean-transcript-text");
const memoStep = document.getElementById("memo-step");
const asrProviderSelect = document.getElementById("asr-provider-select");
const llmProviderSelect = document.getElementById("llm-provider-select");
const resultMarkdown = document.getElementById("result-markdown");
const reviewPanel = document.getElementById("review-panel");
const reviewSummary = document.getElementById("review-summary");
const reviewItems = document.getElementById("review-items");
const deriveWeeklyPanel = document.getElementById("derive-weekly-panel");
const deriveWeeklyButton = document.getElementById("derive-weekly-button");
const deriveWeeklyMaterials = document.getElementById("derive-weekly-materials");
const deriveWeeklyInstructions = document.getElementById("derive-weekly-instructions");
const deriveWeeklyCopy = document.getElementById("derive-weekly-copy");
const deriveWeeklyStatus = document.getElementById("derive-weekly-status");
const weeklyReportDownloadLink = document.getElementById("weekly-report-download-link");
const generateFinalButton = document.getElementById("generate-final-button");
const toggleEnhancementButton = document.getElementById("toggle-enhancement-button");
const transcriptDownloadLink = document.getElementById("transcript-download-link");
const roleTranscriptDownloadLink = document.getElementById("role-transcript-download-link");
const evidenceBankDownloadLink = document.getElementById("evidence-bank-download-link");
const markdownDownloadLink = document.getElementById("markdown-download-link");
const pptDownloadLink = document.getElementById("ppt-download-link");
const resultTranscript = document.getElementById("result-transcript");
const obsidianExportButton = document.getElementById("obsidian-export-button");
const advancedArtifacts = document.getElementById("advanced-artifacts");
const jobsList = document.getElementById("jobs-list");
const statusPill = document.getElementById("status-pill");
const refreshButton = document.getElementById("refresh-button");
const submitButton = document.getElementById("submit-button");
const progressBar = document.getElementById("progress-bar");
const progressLabel = document.getElementById("progress-label");
const obsidianVault = document.getElementById("obsidian-vault");
const llmStatus = document.getElementById("llm-status");
const templateRecommendation = document.getElementById("template-recommendation");

let latestJob = null;
let runtimeProviders = null;
let enhancementOpen = false;

boot();

templateCards.forEach((card) => {
  card.addEventListener("click", () => {
    const radio = card.querySelector('input[type="radio"]');
    radio.checked = true;
    syncTemplateCards();
    updateActionLabels();
  });
});

form.addEventListener("submit", handleDraftGeneration);
transcribeButton.addEventListener("click", handleTranscriptGeneration);
form.querySelector('input[name="audio"]').addEventListener("change", clearCleanTranscript);
form.querySelector('input[name="transcriptFile"]').addEventListener("change", clearCleanTranscript);
generateFinalButton.addEventListener("click", finalizeJob);
deriveWeeklyButton.addEventListener("click", deriveWeeklyReport);
toggleEnhancementButton.addEventListener("click", toggleEnhancementPanel);
obsidianExportButton.addEventListener("click", exportToObsidian);
refreshButton.addEventListener("click", loadJobs);

async function boot() {
  await Promise.all([loadProviders(), loadJobs()]);
  syncTemplateCards();
  updateActionLabels();
  updateDerivedWorkflowState(null);
  setMemoStepReady(false);
  resetProgress();
}

async function handleTranscriptGeneration() {
  transcribeButton.disabled = true;
  setMemoStepReady(false);
  updateStatus("Cleaning transcript");
  setProgress("parse");
  transcriptStepStatus.textContent = "Processing source...";

  try {
    const formData = new FormData(form);
    const audioFile = formData.get("audio");
    const transcriptFile = formData.get("transcriptFile");

    if ((!audioFile || !audioFile.size) && (!transcriptFile || !transcriptFile.size)) {
      throw new Error("Add an audio file or transcript file before creating the clean transcript.");
    }

    formData.set("transcriptText", cleanTranscriptText.value || "");
    const response = await fetch("/api/transcript", {
      method: "POST",
      body: formData,
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Failed to create clean transcript.");
    }

    cleanTranscriptText.value = data.transcriptText || "";
    resultTranscript.textContent = data.transcriptText || "No transcript generated.";
    transcriptStepStatus.textContent = `Ready | ${data.sourceType} | ${data.asrProvider}`;
    updateStatus("Transcript ready");
    progressLabel.textContent = "Clean transcript ready. Choose a memo style and generate the memo draft.";
    setProgress("review");
    setMemoStepReady(true);
  } catch (error) {
    updateStatus("Error");
    progressLabel.textContent = error.message;
    transcriptStepStatus.textContent = error.message;
    resultTranscript.textContent = `Transcript generation failed.\n\n${error.message}`;
    cleanTranscriptText.value = "";
    setMemoStepReady(false);
    resetProgress();
  } finally {
    transcribeButton.disabled = false;
  }
}

async function loadProviders() {
  const response = await fetch("/api/providers");
  const data = await response.json();
  runtimeProviders = data;

  asrProviderSelect.innerHTML = data.asr
    .map((provider) => {
      const suffix = provider.configured ? "" : " (not configured)";
      return `<option value="${provider.id}">${provider.name}${suffix}</option>`;
    })
    .join("");

  llmProviderSelect.innerHTML = data.llm.providers
    .filter((provider) => provider.id !== "openai" || provider.configured)
    .map((provider) => {
      const suffix = provider.configured ? "" : " (not configured)";
      const selected = provider.id === data.llm.defaultProvider ? " selected" : "";
      return `<option value="${provider.id}"${selected}>${provider.label}${suffix}</option>`;
    })
    .join("");

  obsidianVault.textContent = data.integrations?.obsidian?.configured
    ? data.integrations.obsidian.vaultPath
    : "Not detected";

  llmStatus.textContent =
    data.llm.defaultProvider === "none"
      ? "LLM is not configured. The app will fall back to lighter extraction."
      : `Default LLM: ${data.llm.defaultProvider}`;

  updateObsidianButtonState();
}

async function loadJobs() {
  const response = await fetch("/api/jobs");
  const data = await response.json();

  if (!data.jobs.length) {
    jobsList.innerHTML = `<p class="muted">No jobs yet. Generate one from the new workflow.</p>`;
    return;
  }

  jobsList.innerHTML = data.jobs
    .slice(0, 8)
    .map(
      (job) => `
        <article class="job-card">
          <h3>${escapeHtml(job.job.templateName)}</h3>
          <p>${escapeHtml(job.job.title || "Untitled")} | ${escapeHtml(job.job.company || "Unknown company")}</p>
          <p>${escapeHtml(job.job.modelMode)} | ${escapeHtml(job.job.createdAt)}</p>
          <p>${escapeHtml(job.job.summary || "")}</p>
        </article>
      `,
    )
    .join("");
}

async function handleDraftGeneration(event) {
  event.preventDefault();
  submitButton.disabled = true;
  latestJob = null;
  enhancementOpen = false;
  updateStatus("Generating draft");
  setProgress("upload");

  try {
    const formData = new FormData(form);
    ensureOptionalDefaults(formData);
    const transcriptText = String(formData.get("transcriptText") || "").trim();

    if (!transcriptText) {
      throw new Error("Create a clean transcript first, then generate the memo.");
    }

    setProgress("parse");
    const response = await fetch("/api/process", {
      method: "POST",
      body: formData,
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Failed to generate draft.");
    }

    latestJob = data;
    setProgress("draft");
    renderJob(data, { finalized: false });
    if (selectedTemplateId() === "oi-news-report") {
      updateStatus("Creating OI News PPT");
      progressLabel.textContent = "Draft ready. Rendering portrait PPT now.";
      await finalizeLatestJob({ selectedEnrichmentFields: [] });
    } else {
      setProgress("review");
    }
    await loadJobs();
  } catch (error) {
    updateStatus("Error");
    progressLabel.textContent = error.message;
    resultMarkdown.textContent = `Generation failed.\n\n${error.message}`;
    renderCoverageReview(null);
    renderTemplateRecommendation(null);
    hideActionLinks();
    resetProgress();
  } finally {
    submitButton.disabled = false;
  }
}

async function finalizeJob() {
  if (!latestJob?.job?.id) {
    return;
  }

  generateFinalButton.disabled = true;
  try {
    const selectedEnrichmentFields = Array.from(
      reviewItems.querySelectorAll('input[type="checkbox"]:checked'),
    ).map((input) => input.dataset.gapId);
    await finalizeLatestJob({ selectedEnrichmentFields });
  } catch (error) {
    updateStatus("Error");
    progressLabel.textContent = error.message;
  } finally {
    generateFinalButton.disabled = false;
  }
}

async function finalizeLatestJob({ selectedEnrichmentFields = [] }) {
  updateStatus("Creating final file");
  setProgress("final");

  const response = await fetch(`/api/jobs/${latestJob.job.id}/finalize`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      selectedEnrichmentFields,
      llmProvider: llmProviderSelect.value,
    }),
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "Failed to create final file.");
  }

  latestJob = data;
  renderJob(data, { finalized: true });
}

async function exportToObsidian() {
  if (!latestJob?.job?.id) {
    return;
  }

  obsidianExportButton.disabled = true;

  try {
    const response = await fetch(`/api/jobs/${latestJob.job.id}/export/obsidian`, {
      method: "POST",
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Failed to export to Obsidian.");
    }

    latestJob = data.job;
    updateStatus("Saved to Obsidian");
    progressLabel.textContent = `Saved to Obsidian: ${data.export.path}`;
    obsidianExportButton.textContent = "Saved";
    obsidianExportButton.classList.add("saved-state");
  } catch (error) {
    updateStatus("Error");
    progressLabel.textContent = error.message;
    obsidianExportButton.textContent = "Save Failed";
  } finally {
    updateObsidianButtonState();
  }
}

async function deriveWeeklyReport() {
  if (!latestJob?.job?.id) {
    return;
  }

  deriveWeeklyButton.disabled = true;
  updateStatus("Creating weekly report");
  setProgress("final");
  deriveWeeklyPanel.classList.add("working");
  deriveWeeklyStatus.textContent = "Rendering weekly report PPT...";
  progressLabel.textContent = "Rendering weekly report PPT. This can take a little while because PowerPoint is being generated.";
  let failed = false;

  try {
    const formData = new FormData();
    formData.set("targetTemplateId", "weekly-report");
    formData.set("llmProvider", llmProviderSelect.value);
    formData.set("userInstructions", deriveWeeklyInstructions.value || "");
    Array.from(deriveWeeklyMaterials.files || []).forEach((file) => {
      formData.append("materials", file);
    });

    const response = await fetch(`/api/jobs/${latestJob.job.id}/derive`, {
      method: "POST",
      body: formData,
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to generate weekly report.");
    }

    latestJob = data;
    renderJob(data, { finalized: true });
    deriveWeeklyStatus.textContent = "Weekly Report PPT ready. Download it below.";
    progressLabel.textContent = "Weekly Report PPT ready. Use the download button in Step 03.";
    await loadJobs();
  } catch (error) {
    failed = true;
    updateStatus("Error");
    progressLabel.textContent = error.message;
    deriveWeeklyStatus.textContent = error.message;
  } finally {
    deriveWeeklyPanel.classList.remove("working");
    if (failed) {
      deriveWeeklyButton.disabled = false;
      deriveWeeklyMaterials.disabled = false;
      deriveWeeklyInstructions.disabled = false;
    } else {
      updateDerivedWorkflowState(latestJob);
    }
  }
}

function renderJob(jobRecord, { finalized }) {
  const providerLabel =
    jobRecord.job.modelMode === "llm"
      ? `LLM (${jobRecord.job.llmProvider})`
      : "Fallback";
  updateStatus(finalized ? `${providerLabel} | Final ready` : `${providerLabel} | Draft ready`);
  progressLabel.textContent = finalized
    ? "Final file ready for download or import. Clean transcript remains available above."
    : jobRecord.job.templateId === "oi-news-report"
      ? "Draft ready. OI News Report will auto-render into PPT."
      : "Draft ready. Clean transcript and output are both available for preview/download.";
  resultTranscript.textContent = jobRecord.structured?.meeting?.transcript || jobRecord.inputs?.transcriptText || "No transcript generated.";
  resultMarkdown.textContent = jobRecord.markdown || "No markdown generated.";
  const canEnhanceWeekly = jobRecord.job.templateId === "weekly-report";
  renderCoverageReview(canEnhanceWeekly ? jobRecord.review : null, enhancementOpen && canEnhanceWeekly);
  renderTemplateRecommendation(jobRecord.review);
  updateActionLabels();
  updateDownloadLinks(jobRecord);
  updateDerivedWorkflowState(jobRecord);

  updateObsidianButtonState();

  if (finalized) {
    setProgress("done");
  }
}

function updateDerivedWorkflowState(jobRecord) {
  const canDeriveWeekly = ["interview-knowledge-base", "interview-free-style"].includes(jobRecord?.job?.templateId);
  const hasWeeklyOutput = jobRecord?.job?.templateId === "weekly-report" && Boolean(jobRecord?.job?.pptPath);

  deriveWeeklyPanel.classList.remove("hidden");
  deriveWeeklyPanel.classList.toggle("ready", canDeriveWeekly || hasWeeklyOutput);
  deriveWeeklyPanel.classList.toggle("disabled-state", !canDeriveWeekly && !hasWeeklyOutput);
  deriveWeeklyButton.disabled = !canDeriveWeekly;
  deriveWeeklyMaterials.disabled = !canDeriveWeekly;
  deriveWeeklyInstructions.disabled = !canDeriveWeekly;

  if (hasWeeklyOutput) {
    deriveWeeklyCopy.textContent = "Weekly Report PPT is ready.";
    deriveWeeklyStatus.textContent = "Weekly Report PPT ready. Use the large download button above.";
    weeklyReportDownloadLink.href = `/api/jobs/${jobRecord.job.id}/artifacts/ppt`;
    weeklyReportDownloadLink.classList.remove("hidden");
    deriveWeeklyButton.classList.add("hidden");
  } else if (canDeriveWeekly) {
    deriveWeeklyCopy.textContent =
      "Ready. Use the current memo and optional new materials to generate the weekly report deck.";
    deriveWeeklyStatus.textContent = "Ready to generate weekly report.";
    weeklyReportDownloadLink.classList.add("hidden");
    weeklyReportDownloadLink.removeAttribute("href");
    deriveWeeklyButton.classList.remove("hidden");
    toggleEnhancementButton.classList.add("hidden");
    reviewPanel.classList.add("hidden");
  } else {
    deriveWeeklyCopy.textContent =
      "Available after an Interview Memo or Free Style Memo is generated.";
    deriveWeeklyStatus.textContent = "Waiting for a memo draft.";
    weeklyReportDownloadLink.classList.add("hidden");
    weeklyReportDownloadLink.removeAttribute("href");
    deriveWeeklyButton.classList.remove("hidden");
    toggleEnhancementButton.classList.add("hidden");
    reviewPanel.classList.add("hidden");
  }
}

function renderCoverageReview(review, isOpen = false) {
  if (!review?.items?.length) {
    reviewPanel.classList.add("hidden");
    reviewSummary.innerHTML = "";
    reviewItems.innerHTML = "";
    toggleEnhancementButton.classList.add("hidden");
    return;
  }

  reviewPanel.classList.toggle("hidden", !isOpen);
  reviewSummary.innerHTML = buildReviewSummaryHtml(review);
  reviewItems.innerHTML = `
    <div class="enhancement-table-wrap">
      <table class="enhancement-table">
        <thead>
          <tr>
            <th>Use</th>
            <th>Area</th>
            <th>Status</th>
            <th>Current Evidence</th>
            <th>Public / Prepared Addition</th>
          </tr>
        </thead>
        <tbody>
          ${review.items.map(renderReviewTableRow).join("")}
        </tbody>
      </table>
    </div>
  `;

  toggleEnhancementButton.classList.remove("hidden");
  toggleEnhancementButton.textContent = isOpen ? "Hide Enhancement" : "Enhancement";
}

function renderReviewTableRow(item) {
  const shouldCheck = (item.status === "missing" || item.status === "thin") && item.enrichable;
  const evidence = item.evidencePreview || {};
  const currentEvidence = [
    ...formatEvidenceList("User", evidence.userProvided),
    ...formatEvidenceList("Draft", evidence.currentDraft),
  ];
  const preparedEvidence = [
    ...formatEvidenceList("Public", evidence.publicSource),
    item.proposedAddition ? `<span>${escapeHtml(item.proposedAddition)}</span>` : "",
  ].filter(Boolean);

  return `
    <tr>
      <td>
        <input
          type="checkbox"
          data-gap-id="${escapeHtml(item.id)}"
          ${shouldCheck ? "checked" : ""}
          ${item.enrichable ? "" : "disabled"}
          aria-label="Enhance ${escapeHtml(item.label)}"
        />
      </td>
      <td>
        <strong>${escapeHtml(item.label)}</strong>
        <small>${escapeHtml(item.recommendation || "")}</small>
      </td>
      <td>
        <span class="status-chip status-${escapeHtml(item.status || "unknown")}">${escapeHtml(item.status || "unknown")}</span>
        <small>${escapeHtml(item.source || "unknown")}</small>
      </td>
      <td>${currentEvidence.length ? currentEvidence.join("") : "<span class=\"muted\">No grounded evidence yet.</span>"}</td>
      <td>${preparedEvidence.length ? preparedEvidence.join("") : "<span class=\"muted\">No automatic addition planned.</span>"}</td>
    </tr>
  `;
}

function formatEvidenceList(label, items = []) {
  return (items || [])
    .filter(Boolean)
    .slice(0, 2)
    .map((item) => `<span><strong>${escapeHtml(label)}:</strong> ${escapeHtml(item)}</span>`);
}

function buildReviewSummaryHtml(review) {
  const parts = [];
  if (review.summary) {
    parts.push(`<strong>${escapeHtml(review.summary)}</strong>`);
  }
  if (review.focusProfile?.primaryLabel) {
    parts.push(`<span>Detected focus: ${escapeHtml(review.focusProfile.primaryLabel)}</span>`);
  }
  if (review.recommendation?.recommendedTemplateName) {
    parts.push(`<span>Suggested memo shape: ${escapeHtml(review.recommendation.recommendedTemplateName)}</span>`);
  }
  if (review.provenanceSummary) {
    parts.push(`<span>${escapeHtml(review.provenanceSummary)}</span>`);
  }
  if (Array.isArray(review.warnings) && review.warnings.length) {
    parts.push(
      `<span>Render warnings: ${escapeHtml(review.warnings.join(" | "))}</span>`,
    );
  }
  return parts.join("<br />");
}

function renderTemplateRecommendation(review) {
  if (!templateRecommendation) {
    return;
  }

  if (!review?.recommendation?.recommendedTemplateName) {
    templateRecommendation.textContent = "The app can recommend a memo structure after it reads your transcript.";
    return;
  }

  const selected = selectedTemplateId();
  const suggestion = review.recommendation.recommendedTemplateId;
  const focus = review.focusProfile?.primaryLabel || "Mixed Diligence";
  const alignment =
    selected === suggestion
      ? "Current selection matches the detected interview shape."
      : "You can keep your current choice, but the suggested structure may read more cleanly.";

  templateRecommendation.textContent = `Detected focus: ${focus}. Suggested output: ${review.recommendation.recommendedTemplateName}. ${alignment}`;
}

function updateDownloadLinks(jobRecord) {
  const transcriptHref = `/api/jobs/${jobRecord.job.id}/artifacts/transcript`;
  transcriptDownloadLink.href = transcriptHref;
  transcriptDownloadLink.classList.remove("hidden");
  transcriptDownloadLink.textContent = "Download Transcript";

  roleTranscriptDownloadLink.href = `/api/jobs/${jobRecord.job.id}/artifacts/role-transcript`;
  roleTranscriptDownloadLink.classList.remove("hidden");
  roleTranscriptDownloadLink.textContent = "Role-labeled Transcript";

  evidenceBankDownloadLink.href = `/api/jobs/${jobRecord.job.id}/artifacts/evidence-bank`;
  evidenceBankDownloadLink.classList.remove("hidden");
  evidenceBankDownloadLink.textContent = "Evidence Bank JSON";
  advancedArtifacts.classList.remove("hidden");

  const markdownHref = `/api/jobs/${jobRecord.job.id}/artifacts/markdown`;
  markdownDownloadLink.href = markdownHref;
  markdownDownloadLink.classList.remove("hidden");
  markdownDownloadLink.textContent =
    ["interview-knowledge-base", "interview-free-style"].includes(jobRecord.job.templateId)
      ? "Download Markdown"
      : "Download Draft Markdown";

  if (jobRecord.job.pptPath) {
    pptDownloadLink.href = `/api/jobs/${jobRecord.job.id}/artifacts/ppt`;
    pptDownloadLink.classList.remove("hidden");
    pptDownloadLink.textContent =
      jobRecord.job.templateId === "oi-news-report"
        ? "Download OI News PPT"
        : jobRecord.job.templateId === "weekly-report"
          ? "Download Weekly Report PPT"
          : "Download PPT";
    pptDownloadLink.classList.toggle("primary-download", jobRecord.job.templateId === "weekly-report");
  } else {
    pptDownloadLink.classList.add("hidden");
    pptDownloadLink.classList.remove("primary-download");
    pptDownloadLink.removeAttribute("href");
  }
}

function updateActionLabels() {
  const templateId = selectedTemplateId();
  generateFinalButton.textContent =
    ["interview-knowledge-base", "interview-free-style"].includes(templateId)
      ? "Apply Enhancement And Refresh Memo"
      : templateId === "oi-news-report"
        ? "Apply Enhancement And Refresh OI News PPT"
        : "Apply Enhancement And Refresh PPT";
}

function syncTemplateCards() {
  templateCards.forEach((card) => {
    const radio = card.querySelector('input[type="radio"]');
    card.classList.toggle("active", radio.checked);
  });
}

function selectedTemplateId() {
  return form.querySelector('input[name="templateId"]:checked')?.value || "interview-knowledge-base";
}

function ensureOptionalDefaults(formData) {
  const transcriptText = String(cleanTranscriptText?.value || formData.get("transcriptText") || "").trim();
  const notesText = String(formData.get("notesText") || "").trim();
  const instructions = String(formData.get("userInstructions") || "").trim();

  if (!formData.get("meetingTitle")) {
    formData.set("meetingTitle", "");
  }

  if (!formData.get("company")) {
    formData.set("company", "");
  }

  if (!notesText && instructions) {
    formData.set("notesText", `Instructions:\n${instructions}`);
  } else if (instructions) {
    formData.set("notesText", `${notesText}\n\nInstructions:\n${instructions}`);
  }

  formData.set("transcriptText", transcriptText);
}

function setMemoStepReady(isReady) {
  memoStep.classList.toggle("locked-step", !isReady);
  submitButton.disabled = !isReady;
}

function clearCleanTranscript() {
  cleanTranscriptText.value = "";
  resultTranscript.textContent = "Cleaned transcript will appear here after audio transcription or transcript upload.";
  transcriptStepStatus.textContent = "Source changed. Create a new clean transcript.";
  setMemoStepReady(false);
}

function updateStatus(text) {
  statusPill.textContent = text;
}

function resetProgress() {
  progressBar.style.width = "0%";
  progressLabel.textContent = "Waiting to start.";
}

function setProgress(currentStep) {
  const widths = {
    upload: 12,
    parse: 34,
    draft: 60,
    review: 80,
    final: 92,
    done: 100,
  };
  const labels = {
    upload: "Uploading source materials...",
    parse: "Reading transcript and supporting files...",
    draft: "Generating first draft...",
    review: "Preparing review step...",
    final: "Creating final output...",
    done: "Completed.",
  };

  progressBar.style.width = `${widths[currentStep] || 0}%`;
  progressLabel.textContent = labels[currentStep] || "Processing...";
}

function hideActionLinks() {
  transcriptDownloadLink.classList.add("hidden");
  roleTranscriptDownloadLink.classList.add("hidden");
  evidenceBankDownloadLink.classList.add("hidden");
  advancedArtifacts.classList.add("hidden");
  markdownDownloadLink.classList.add("hidden");
  pptDownloadLink.classList.add("hidden");
  weeklyReportDownloadLink.classList.add("hidden");
  weeklyReportDownloadLink.removeAttribute("href");
  toggleEnhancementButton.classList.add("hidden");
  updateDerivedWorkflowState(null);
  updateObsidianButtonState();
  reviewPanel.classList.add("hidden");
}

function updateObsidianButtonState() {
  const obsidianReady = Boolean(runtimeProviders?.integrations?.obsidian?.configured);
  const memoReady = ["interview-knowledge-base", "interview-free-style"].includes(latestJob?.job?.templateId);
  const savedPath = latestJob?.artifacts?.obsidianPath || latestJob?.job?.obsidianPath;
  obsidianExportButton.classList.toggle("saved-state", Boolean(savedPath));
  obsidianExportButton.disabled = !(obsidianReady && memoReady) || Boolean(savedPath);
  obsidianExportButton.textContent = !obsidianReady
    ? "Obsidian Not Detected"
    : savedPath
      ? "Saved"
      : "Save to Obsidian";
}

function toggleEnhancementPanel() {
  enhancementOpen = !enhancementOpen;
  renderCoverageReview(latestJob?.review, enhancementOpen);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
