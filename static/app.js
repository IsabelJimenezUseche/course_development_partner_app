const $ = (selector) => document.querySelector(selector);

const elements = {
  modelStatus: $("#modelStatus"), projectForm: $("#projectForm"), projectName: $("#projectName"),
  courseName: $("#courseName"), courseLevel: $("#courseLevel"), classTime: $("#classTime"),
  learningOutcome: $("#learningOutcome"), outcomeHint: $("#outcomeHint"), sourceNotes: $("#sourceNotes"),
  messageInput: $("#messageInput"), sendButton: $("#sendButton"), sendButtonLabel: $("#sendButtonLabel"),
  composerHelp: $("#composerHelp"), conversation: $("#conversation"), sessionLabel: $("#sessionLabel"),
  skillRuntimeLabel: $("#skillRuntimeLabel"), sourceFiles: $("#sourceFiles"),
  dataAcknowledgment: $("#dataAcknowledgment"), dropZone: $("#dropZone"), uploadStatus: $("#uploadStatus"),
  sourceList: $("#sourceList"), sourceCount: $("#sourceCount"), saveStatus: $("#saveStatus"),
  currentProjectName: $("#currentProjectName"), projectSwitcher: $("#projectSwitcher"),
  projectsDialog: $("#projectsDialog"), projectList: $("#projectList"), newProjectButton: $("#newProjectButton"),
  decisionDialog: $("#decisionDialog"), decisionTitle: $("#decisionTitle"), decisionQuestion: $("#decisionQuestion"), decisionStatus: $("#decisionStatus"),
  decisionOptions: $("#decisionOptions"), customDecisionButton: $("#customDecisionButton"),
  artifactsButton: $("#artifactsButton"), artifactCount: $("#artifactCount"),
  artifactsDialog: $("#artifactsDialog"), artifactList: $("#artifactList"),
  stateButton: $("#stateButton"), stateCount: $("#stateCount"), stateDialog: $("#stateDialog"),
  stateList: $("#stateList"), stateProfile: $("#stateProfile"), runValidatorsButton: $("#runValidatorsButton"),
  validationResults: $("#validationResults"), stateTemplatePicker: $("#stateTemplatePicker"),
  stateEditorForm: $("#stateEditorForm"), stateEditorLabel: $("#stateEditorLabel"),
  stateEditorContent: $("#stateEditorContent"), stateEditorCancel: $("#stateEditorCancel"),
  artifactToolButton: $("#artifactToolButton"), artifactToolDialog: $("#artifactToolDialog"),
  artifactToolForm: $("#artifactToolForm"), artifactToolKind: $("#artifactToolKind"),
  artifactToolTitle: $("#artifactToolTitle"), artifactToolSubtitle: $("#artifactToolSubtitle"),
  artifactToolOutline: $("#artifactToolOutline"), artifactToolStatus: $("#artifactToolStatus"),
  generateArtifactButton: $("#generateArtifactButton"),
  artifactPreviewDialog: $("#artifactPreviewDialog"), artifactPreviewTitle: $("#artifactPreviewTitle"),
  artifactPreview: $("#artifactPreview"), artifactPreviewActions: $("#artifactPreviewActions"),
  sourcePreviewDialog: $("#sourcePreviewDialog"), sourcePreviewTitle: $("#sourcePreviewTitle"),
  sourcePreviewMeta: $("#sourcePreviewMeta"), sourceVectorPreview: $("#sourceVectorPreview"),
  sourcePreviewDownload: $("#sourcePreviewDownload"),
  exportButton: $("#exportButton"), exportDialog: $("#exportDialog"),
  exportEyebrow: $("#exportEyebrow"), exportDialogTitle: $("#exportDialogTitle"),
  exportDescription: $("#exportDescription"), exportMarkdownDescription: $("#exportMarkdownDescription"),
  exportHtmlDescription: $("#exportHtmlDescription"), exportJsonDescription: $("#exportJsonDescription"),
  exportZipDescription: $("#exportZipDescription"), exportZip: $("#exportZip"),
  exportMarkdown: $("#exportMarkdown"), exportHtml: $("#exportHtml"), exportJson: $("#exportJson"),
  traceButton: $("#traceButton"), traceDialog: $("#traceDialog"),
  traceSummary: $("#traceSummary"), traceTimeline: $("#traceTimeline"),
  settingsButton: $("#settingsButton"), settingsDialog: $("#settingsDialog"), settingsForm: $("#settingsForm"),
  settingsTarget: $("#settingsTarget"), settingsPath: $("#settingsPath"), settingsFields: $("#settingsFields"),
  settingsStatus: $("#settingsStatus"), saveSettingsButton: $("#saveSettingsButton"),
  emptyState: $("#emptyState"), projectPanel: $("#projectPanel"),
  mobileBriefToggle: $("#mobileBriefToggle"), mobileBriefClose: $("#mobileBriefClose"),
  footerYear: $("#footerYear"),
};

const state = {
  configured: false,
  busy: false,
  loadingProject: false,
  started: false,
  projectId: null,
  project: null,
  projects: [],
  messages: [],
  artifacts: [],
  sources: [],
  skillRuntime: null,
  uploadMaxBytes: 20 * 1024 * 1024,
  environment: null,
  saveTimer: null,
  pendingDecision: null,
  artifactToolNotice: null,
  stateFiles: [],
  stateAssets: [],
  editingStateFile: null,
};

function openDialog(dialog) {
  if (!dialog.open) dialog.showModal();
}

function closeDialog(dialog) {
  if (dialog.open) dialog.close();
}

function setMobileBriefOpen(open) {
  elements.projectPanel.classList.toggle("mobile-open", open);
  elements.mobileBriefToggle.setAttribute("aria-expanded", String(open));
}

function createButtonIcon(name) {
  const paths = {
    artifacts: '<path d="M3 6.5h6l2 2h10v10.5H3z"/><path d="M3 6.5V5h7l2 2h9v1.5"/>',
    check: '<path d="m5 12 4 4L19 6"/>',
    choices: '<circle cx="12" cy="12" r="9"/><path d="M9.8 9a2.4 2.4 0 1 1 3.5 2.15c-.8.4-1.3.9-1.3 1.85M12 17h.01"/>',
    download: '<path d="M12 3v12m0 0 4-4m-4 4-4-4M5 21h14"/>',
    preview: '<path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6z"/><circle cx="12" cy="12" r="2.5"/>',
    trash: '<path d="M4 7h16M9 7V4h6v3M7 7l1 13h8l1-13M10 11v5M14 11v5"/>',
  };
  const iconSpan = document.createElement("span"); iconSpan.className = "button-icon"; iconSpan.setAttribute("aria-hidden", "true");
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg"); svg.setAttribute("viewBox", "0 0 24 24"); svg.innerHTML = paths[name] || paths.choices; iconSpan.appendChild(svg);
  return iconSpan;
}

function setIconLabel(control, icon, label, iconOnly = false) {
  const iconSpan = createButtonIcon(icon);
  control.replaceChildren(iconSpan);
  if (!iconOnly) { const labelSpan = document.createElement("span"); labelSpan.className = "button-label"; labelSpan.textContent = label; control.appendChild(labelSpan); }
  control.classList.toggle("icon-only", iconOnly);
}

function setModelStatus(text) {
  elements.modelStatus.dataset.tooltip = text;
  elements.modelStatus.setAttribute("aria-label", `Model status: ${text}`);
  elements.modelStatus.querySelector(".status-label").textContent = text;
}

const DEFAULT_MODE = "Co-design";
// The skill renamed this mode; projects saved before the rename still report "Studio".
const LEGACY_MODES = { Studio: DEFAULT_MODE };

function normalizeMode(mode) {
  return LEGACY_MODES[mode] || mode || DEFAULT_MODE;
}

function selectedMode() {
  return document.querySelector('input[name="mode"]:checked')?.value || DEFAULT_MODE;
}

function projectPayload() {
  return {
    name: elements.projectName.value.trim() || "Untitled course project",
    course_name: elements.courseName.value.trim(),
    level: elements.courseLevel.value,
    class_time: elements.classTime.value.trim() || "Not specified",
    outcome: elements.learningOutcome.value.trim(),
    mode: selectedMode(),
    notes: elements.sourceNotes.value.trim(),
  };
}

function formatDate(value) {
  if (!value) return "";
  return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }).format(new Date(value));
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function setSaveStatus(text, kind = "") {
  elements.saveStatus.textContent = text;
  elements.saveStatus.className = `save-status${kind ? ` ${kind}` : ""}`;
}

function fillProjectForm(project) {
  state.loadingProject = true;
  elements.projectName.value = project.name || "";
  elements.courseName.value = project.course_name || "";
  elements.courseLevel.value = project.level || "Undergraduate";
  elements.classTime.value = project.class_time || "50 minutes";
  elements.learningOutcome.value = project.outcome || "";
  elements.sourceNotes.value = project.notes || "";
  const projectMode = normalizeMode(project.mode);
  const mode = document.querySelector(`input[name="mode"][value="${CSS.escape(projectMode)}"]`);
  if (mode) mode.checked = true;
  elements.currentProjectName.textContent = project.name;
  elements.sessionLabel.textContent = `${projectMode} project · ${state.messages.length} saved messages`;
  state.loadingProject = false;
}

function updateControls() {
  const hasOutcome = Boolean(elements.learningOutcome.value.trim());
  const canSend = state.configured && state.projectId && !state.busy && (state.started ? Boolean(elements.messageInput.value.trim()) : hasOutcome);
  elements.sendButton.disabled = !canSend;
  elements.messageInput.disabled = !state.started || state.busy;
  elements.sendButtonLabel.textContent = state.busy ? "Working…" : state.started ? "Send" : "Start design session";
  if (!state.configured) elements.composerHelp.textContent = "Add the Purdue model ID and API key in app/.env.";
  else if (!state.started) elements.composerHelp.textContent = hasOutcome ? "Ready to prepare the first design decision." : "Complete the learning outcome to start.";
  else elements.composerHelp.textContent = state.busy ? "Preparing a source-grounded response…" : "Use ⌘/Ctrl + Enter to send. Your work saves automatically.";
  if (!state.busy && state.artifactToolNotice) elements.composerHelp.textContent = state.artifactToolNotice;
}

function createLoadingMessage() {
  const article = document.createElement("article");
  article.className = "message assistant-message";
  article.innerHTML = '<div class="message-avatar" aria-hidden="true">CD</div><div class="message-content"><div class="message-label">Design partner</div><div class="loading-dots" aria-label="Preparing the next design decision"><span></span><span></span><span></span></div></div>';
  elements.conversation.appendChild(article);
  elements.conversation.scrollTop = elements.conversation.scrollHeight;
  return article;
}

function artifactsForMessage(messageId, fallbackArtifact = null) {
  const linked = state.artifacts.filter((artifact) => artifact.message_id === messageId);
  if (!linked.length && fallbackArtifact) linked.push(fallbackArtifact);
  return linked;
}

function selectionForDecision(originMessageId) {
  return state.messages.find((message) => message.decision_trace?.origin_message_id === originMessageId)?.decision_trace || null;
}

function enhanceMarkdownTables(root) {
  root.querySelectorAll("table").forEach((table) => {
    if (table.closest(".table-scroll")) return;
    const headers = [...table.querySelectorAll("thead th")].map((cell) => cell.textContent.trim());
    if (headers.length >= 7) table.classList.add("wide-table");
    table.querySelectorAll("tbody tr").forEach((row) => {
      [...row.children].forEach((cell, index) => {
        if (headers[index]) cell.dataset.label = headers[index];
      });
    });
    const wrapper = document.createElement("div");
    wrapper.className = "table-scroll";
    wrapper.setAttribute("role", "region");
    wrapper.setAttribute("aria-label", headers.length ? `Table: ${headers.join(", ")}` : "Data table");
    wrapper.tabIndex = 0;
    table.parentNode.insertBefore(wrapper, table);
    wrapper.appendChild(table);
  });
}

function createMessage(message, options = {}) {
  const article = document.createElement("article");
  article.className = `message ${message.role === "user" ? "user-message" : "assistant-message"}${options.error ? " error-message" : ""}`;
  if (message.role !== "user") {
    const avatar = document.createElement("div");
    avatar.className = "message-avatar";
    avatar.setAttribute("aria-hidden", "true");
    avatar.textContent = options.error ? "!" : "CD";
    article.appendChild(avatar);
  }
  const content = document.createElement("div");
  content.className = "message-content";
  const label = document.createElement("div");
  label.className = "message-label";
  label.textContent = options.error ? "Connection issue" : message.role === "user" ? "You" : "Design partner";
  const body = document.createElement("div");
  if (message.role === "assistant" && message.html) {
    body.className = "message-text markdown-body";
    body.innerHTML = message.html;
    enhanceMarkdownTables(body);
  } else {
    body.className = "message-text";
    body.textContent = message.content;
  }
  content.append(label, body);

  if (!options.error) {
    const linkedArtifacts = artifactsForMessage(message.id, options.artifact);
    const actions = document.createElement("div");
    actions.className = "message-actions";
    if (message.role === "assistant" && linkedArtifacts.length) {
      const view = document.createElement("button");
      view.type = "button";
      view.className = "message-action";
      setIconLabel(view, "artifacts", linkedArtifacts.length === 1 ? linkedArtifacts[0].kind : `${linkedArtifacts.length} artifacts`);
      view.setAttribute("aria-label", linkedArtifacts.length === 1 ? `Open artifact ${linkedArtifacts[0].title}` : `Open ${linkedArtifacts.length} linked artifacts`);
      view.addEventListener("click", () => {
        if (linkedArtifacts.length === 1) previewArtifact(linkedArtifacts[0]);
        else { renderArtifacts(); openDialog(elements.artifactsDialog); }
      });
      actions.appendChild(view);
    }
    if (message.role === "assistant" && message.auto_decision) {
      const autoPill = document.createElement("span"); autoPill.className = "artifact-pill auto-decision-pill"; autoPill.textContent = `⚡ Auto · ${message.auto_decision.selected_label}`; actions.appendChild(autoPill);
    }
    if (message.role === "assistant" && message.decision && state.project?.mode !== "Auto") {
      const savedSelection = selectionForDecision(message.id);
      const choose = document.createElement("button");
      choose.type = "button";
      choose.className = `message-action${savedSelection ? " decision-recorded" : ""}`;
      setIconLabel(choose, savedSelection ? "check" : "choices", savedSelection ? "Selected" : "Choose");
      choose.setAttribute("aria-label", savedSelection ? `Review selected decision: ${savedSelection.selected_label}` : `Review choices for ${message.decision.question}`);
      choose.addEventListener("click", () => showDecision(message.decision, message.id, savedSelection));
      actions.appendChild(choose);
    }
    if (message.id) {
      const download = document.createElement("button");
      download.type = "button";
      download.className = "message-action";
      setIconLabel(download, "download", "", true);
      download.title = "Download message";
      download.setAttribute("aria-label", `Download ${message.role === "user" ? "instructor" : "design partner"} message from ${formatDate(message.created_at)}`);
      download.addEventListener("click", () => openMessageExport(message));
      actions.appendChild(download);
    }
    if (actions.childElementCount) content.appendChild(actions);
  }

  if (message.skill_runtime || message.created_at) {
    const meta = document.createElement("div");
    meta.className = "message-meta";
    const runtime = message.skill_runtime;
    const sources = message.sources_used?.length ? ` · Sources ${message.sources_used.join(", ")}` : "";
    meta.textContent = `${formatDate(message.created_at)}${runtime ? ` · Skill ${runtime.profile} · ${runtime.fingerprint}` : ""}${sources}`;
    content.appendChild(meta);
  }
  article.appendChild(content);
  elements.conversation.appendChild(article);
  if (!options.noScroll) elements.conversation.scrollTop = elements.conversation.scrollHeight;
  return article;
}

function renderConversation() {
  elements.conversation.replaceChildren();
  state.messages.forEach((message) => createMessage(message, { noScroll: true }));
  state.started = state.messages.length > 0;
  elements.emptyState.classList.toggle("compact", state.started);
  elements.sessionLabel.textContent = `${normalizeMode(state.project?.mode)} project · ${state.messages.length} saved messages`;
  elements.conversation.scrollTop = elements.conversation.scrollHeight;
  updateControls();
  const answeredDecisionIds = new Set(state.messages.map((message) => message.decision_trace?.origin_message_id).filter(Boolean));
  const unresolvedDecision = [...state.messages].reverse().find((message) => message.role === "assistant" && message.decision && !answeredDecisionIds.has(message.id));
  if (unresolvedDecision && state.project?.mode !== "Auto") showDecision(unresolvedDecision.decision, unresolvedDecision.id);
}

function renderSources() {
  elements.sourceList.replaceChildren();
  elements.sourceCount.textContent = String(state.sources.length);
  state.sources.forEach((source) => {
    const item = document.createElement("div"); item.className = "source-item";
    const details = document.createElement("div");
    const name = document.createElement("button"); name.type = "button"; name.className = "source-item-name source-name-button"; name.title = `Preview ${source.filename}`; name.textContent = source.filename; name.addEventListener("click", () => previewSource(source));
    const vectorSummary = source.vector_index ? ` · ${source.vector_index.chunks} vector chunks` : "";
    const meta = document.createElement("div"); meta.className = "source-item-meta"; meta.textContent = `${source.source_id} · ${formatBytes(source.size_bytes)} · ${source.character_count.toLocaleString()} chars${vectorSummary}`;
    details.append(name, meta);
    const actions = document.createElement("div"); actions.className = "source-item-actions";
    const preview = document.createElement("button"); preview.type = "button"; preview.className = "source-action"; setIconLabel(preview, "preview", "", true); preview.title = "Preview extracted text"; preview.setAttribute("aria-label", `Preview ${source.filename}`); preview.addEventListener("click", () => previewSource(source));
    const download = document.createElement("a"); download.className = "source-action"; setIconLabel(download, "download", "", true); download.href = sourceDownloadUrl(source); download.setAttribute("download", ""); download.title = "Download original"; download.setAttribute("aria-label", `Download original ${source.filename}`);
    const remove = document.createElement("button"); remove.className = "source-action danger"; remove.type = "button"; setIconLabel(remove, "trash", "", true); remove.title = "Remove from project"; remove.setAttribute("aria-label", `Remove ${source.filename}`); remove.addEventListener("click", () => removeSource(source));
    actions.append(preview, download, remove); item.append(details, actions); elements.sourceList.appendChild(item);
  });
}

function sourceDownloadUrl(source) {
  return `/api/projects/${encodeURIComponent(state.projectId)}/sources/${encodeURIComponent(source.source_id)}/download`;
}

async function previewSource(source) {
  elements.sourcePreviewTitle.textContent = source.filename;
  elements.sourcePreviewMeta.textContent = `${source.source_id} · ${formatBytes(source.size_bytes)} · Vector overview`;
  const loading = document.createElement("p"); loading.className = "source-preview-loading"; loading.textContent = "Loading vector overview…"; elements.sourceVectorPreview.replaceChildren(loading);
  elements.sourcePreviewDownload.href = sourceDownloadUrl(source);
  elements.sourcePreviewDownload.setAttribute("aria-label", `Download original ${source.filename}`);
  openDialog(elements.sourcePreviewDialog);
  try {
    const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/sources/${encodeURIComponent(source.source_id)}/preview`);
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || "Preview unavailable");
    const vector = result.source.vector_index ? ` · ${result.source.vector_index.chunks} vector chunks` : "";
    elements.sourcePreviewMeta.textContent = `${result.source.source_id} · ${formatBytes(result.source.size_bytes)} · ${result.source.character_count.toLocaleString()} extracted characters${vector} · ${result.vector_metadata.dimensions} dimensions`;
    const cards = result.vector_chunks.map((chunk) => {
      const card = document.createElement("article"); card.className = "vector-chunk-card";
      const header = document.createElement("div"); header.className = "vector-chunk-header";
      const locator = document.createElement("strong"); locator.textContent = chunk.locator;
      const size = document.createElement("span"); size.textContent = `Chunk ${chunk.index + 1} · ${chunk.character_count.toLocaleString()} chars`;
      header.append(locator, size);
      const idea = document.createElement("p"); idea.className = "vector-main-idea"; idea.textContent = chunk.main_idea;
      const keywords = document.createElement("div"); keywords.className = "vector-keywords";
      chunk.keywords.forEach((keyword) => { const tag = document.createElement("span"); tag.textContent = keyword; keywords.appendChild(tag); });
      card.append(header, idea, keywords); return card;
    });
    elements.sourceVectorPreview.replaceChildren(...cards);
  } catch (error) {
    const failure = document.createElement("p"); failure.className = "source-preview-loading"; failure.textContent = error.message; elements.sourceVectorPreview.replaceChildren(failure);
  }
}

const VALIDATION_LABELS = { pass: "Passed", fail: "Errors", incomplete: "Gaps", timeout: "Timed out", error: "Tool error", unavailable: "No validator", empty: "No state yet" };

function renderStateFiles() {
  elements.stateCount.textContent = String(state.stateFiles.length);
  elements.stateList.replaceChildren();
  if (!state.stateFiles.length) {
    const empty = document.createElement("p");
    empty.className = "empty-note";
    empty.textContent = "No portable state files yet. The design partner creates them as the work progresses, or start one from a template below.";
    elements.stateList.append(empty);
    return;
  }
  state.stateFiles.forEach((file) => {
    const row = document.createElement("div");
    row.className = "state-row";
    const label = document.createElement("div");
    const name = document.createElement("strong");
    name.textContent = file.file;
    const meta = document.createElement("small");
    meta.textContent = `${formatBytes(file.bytes)} · updated ${formatDate(file.updated_at)}${file.has_validator ? "" : " · no dedicated validator"}`;
    label.append(name, meta);
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "secondary-button";
    edit.textContent = "Edit";
    edit.addEventListener("click", () => openStateEditor(file.file));
    row.append(label, edit);
    elements.stateList.append(row);
  });
}

function renderValidationResults(report) {
  elements.validationResults.replaceChildren();
  if (!report) return;
  const header = document.createElement("p");
  header.className = `validation-summary ${report.status}`;
  header.textContent = `${VALIDATION_LABELS[report.status] || report.status} · ${report.checks.length} validator${report.checks.length === 1 ? "" : "s"} · profile ${report.design_profile}`;
  elements.validationResults.append(header);
  report.checks.forEach((check) => {
    const block = document.createElement("details");
    block.className = `validation-check ${check.status}`;
    if (check.status !== "pass") block.open = true;
    const summary = document.createElement("summary");
    summary.textContent = `${check.script} — ${VALIDATION_LABELS[check.status] || check.status}`;
    block.append(summary);
    const list = document.createElement("ul");
    check.findings.forEach((finding) => {
      const item = document.createElement("li");
      item.className = `finding ${finding.level}`;
      item.textContent = `${finding.level}: ${finding.message}`;
      list.append(item);
    });
    block.append(list);
    elements.validationResults.append(block);
  });
  const note = document.createElement("p");
  note.className = "empty-note";
  note.textContent = report.scope_note || "";
  elements.validationResults.append(note);
}

async function refreshStateFiles() {
  if (!state.projectId) return;
  const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/state`);
  if (!response.ok) return;
  const result = await response.json();
  state.stateFiles = result.state_files || [];
  renderStateFiles();
}

async function loadStateTemplates() {
  const response = await fetch("/api/skill/assets");
  if (!response.ok) return;
  const result = await response.json();
  state.stateAssets = result.assets || [];
  elements.stateTemplatePicker.replaceChildren();
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Choose a template…";
  elements.stateTemplatePicker.append(placeholder);
  state.stateAssets.forEach((asset) => {
    const option = document.createElement("option");
    option.value = asset.file;
    option.textContent = asset.validator ? `${asset.file} (validated)` : asset.file;
    elements.stateTemplatePicker.append(option);
  });
}

async function openStateEditor(filename, seedFromTemplate = false) {
  const url = seedFromTemplate
    ? `/api/skill/assets/${encodeURIComponent(filename)}`
    : `/api/projects/${encodeURIComponent(state.projectId)}/state/${encodeURIComponent(filename)}`;
  const response = await fetch(url);
  if (!response.ok) return;
  const result = await response.json();
  state.editingStateFile = filename;
  elements.stateEditorLabel.textContent = `${filename}${seedFromTemplate ? " (from skill template)" : ""}`;
  elements.stateEditorContent.value = result.content || "";
  elements.stateEditorForm.hidden = false;
  elements.stateEditorContent.focus();
}

function closeStateEditor() {
  state.editingStateFile = null;
  elements.stateEditorForm.hidden = true;
  elements.stateEditorContent.value = "";
  elements.stateTemplatePicker.value = "";
}

async function saveStateFile(event) {
  event.preventDefault();
  if (!state.editingStateFile || !state.projectId) return;
  const filename = state.editingStateFile;
  const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/state/${encodeURIComponent(filename)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ file: filename, content: elements.stateEditorContent.value, design_profile: elements.stateProfile.value }),
  });
  if (!response.ok) {
    const error = await response.json().catch(() => ({ detail: "Could not save state file" }));
    elements.validationResults.textContent = error.detail;
    return;
  }
  const result = await response.json();
  closeStateEditor();
  await refreshStateFiles();
  if (result.state_file?.validation) {
    renderValidationResults({ status: result.state_file.validation.status, design_profile: elements.stateProfile.value, checks: [result.state_file.validation], scope_note: "Structural check for this file only." });
  }
}

async function runValidators() {
  if (!state.projectId) return;
  elements.runValidatorsButton.disabled = true;
  elements.validationResults.textContent = "Running the skill's validators…";
  try {
    const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/validate?design_profile=${encodeURIComponent(elements.stateProfile.value)}`, { method: "POST" });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || "Validation failed");
    renderValidationResults(result);
  } catch (error) {
    elements.validationResults.textContent = error.message;
  } finally {
    elements.runValidatorsButton.disabled = false;
  }
}

function renderArtifacts() {
  elements.artifactCount.textContent = String(state.artifacts.length);
  elements.artifactList.replaceChildren();
  if (!state.artifacts.length) {
    const empty = document.createElement("div"); empty.className = "empty-list"; empty.textContent = "Artifacts such as worksheets, rubrics, and validation reports will be saved here automatically."; elements.artifactList.appendChild(empty); return;
  }
  state.artifacts.forEach((artifact) => {
    const item = document.createElement("div"); item.className = "artifact-list-item";
    const details = document.createElement("div");
    const kind = document.createElement("span"); kind.className = "artifact-kind"; kind.textContent = artifact.kind;
    const title = document.createElement("strong"); title.textContent = artifact.title;
    const date = document.createElement("small"); date.textContent = `Created ${formatDate(artifact.created_at)}`;
    details.append(kind, title, date);
    const actions = document.createElement("div"); actions.className = "artifact-actions";
    const preview = document.createElement("button"); preview.type = "button"; setIconLabel(preview, "preview", "Preview"); preview.addEventListener("click", () => previewArtifact(artifact));
    const markdown = document.createElement("a"); markdown.textContent = "Markdown"; markdown.href = artifactDownloadUrl(artifact, "markdown"); markdown.setAttribute("download", "");
    const html = document.createElement("a"); html.textContent = "HTML"; html.href = artifactDownloadUrl(artifact, "html"); html.setAttribute("download", "");
    preview.setAttribute("aria-label", `Preview ${artifact.title}`);
    markdown.setAttribute("aria-label", `Download ${artifact.title} as Markdown`);
    html.setAttribute("aria-label", `Download ${artifact.title} as HTML`);
    actions.append(preview);
    if (artifact.has_file) {
      const office = document.createElement("a"); office.textContent = artifact.file_format === "pptx" ? "PowerPoint" : "Word"; office.href = artifactDownloadUrl(artifact, "office"); office.setAttribute("download", ""); office.setAttribute("aria-label", `Download ${artifact.title} as ${office.textContent}`); actions.appendChild(office);
    }
    actions.append(markdown, html); item.append(details, actions); elements.artifactList.appendChild(item);
  });
}

function artifactDownloadUrl(artifact, format) {
  return `/api/projects/${encodeURIComponent(state.projectId)}/artifacts/${encodeURIComponent(artifact.id)}/download?format=${format}`;
}

function previewArtifact(artifact) {
  elements.artifactPreviewTitle.textContent = artifact.title;
  elements.artifactPreview.innerHTML = artifact.html;
  enhanceMarkdownTables(elements.artifactPreview);
  elements.artifactPreviewActions.replaceChildren();
  if (artifact.has_file) {
    const office = document.createElement("a"); office.href = artifactDownloadUrl(artifact, "office"); office.setAttribute("download", ""); office.textContent = `Download ${artifact.file_format === "pptx" ? "PowerPoint" : "Word"}`; elements.artifactPreviewActions.appendChild(office);
  }
  ["markdown", "html"].forEach((format) => {
    const link = document.createElement("a"); link.href = artifactDownloadUrl(artifact, format); link.setAttribute("download", ""); link.textContent = `Download ${format === "html" ? "HTML" : "Markdown"}`; elements.artifactPreviewActions.appendChild(link);
  });
  closeDialog(elements.artifactsDialog);
  openDialog(elements.artifactPreviewDialog);
}

function parseArtifactOutline(outline, kind) {
  return outline.trim().split(/\n\s*\n+/).map((block, index) => {
    const lines = block.split("\n").map((line) => line.trim()).filter(Boolean);
    const heading = (lines.shift() || `Section ${index + 1}`).replace(/^#{1,3}\s*/, "").slice(0, 120);
    const bullets = [];
    const prompts = [];
    const body = [];
    lines.forEach((line) => {
      if (/^[-*]\s+/.test(line)) bullets.push(line.replace(/^[-*]\s+/, ""));
      else if (kind === "worksheet" && line.endsWith("?")) prompts.push(line);
      else body.push(line);
    });
    return { heading, body: body.join(" "), bullets, prompts, checklist: [], response_lines: 3, table: null };
  }).slice(0, 12);
}

async function generateTeachingFile(event) {
  event.preventDefault();
  const kind = elements.artifactToolKind.value;
  const title = elements.artifactToolTitle.value.trim();
  const sections = parseArtifactOutline(elements.artifactToolOutline.value, kind);
  if (!title || !sections.length) return;
  elements.artifactToolStatus.classList.remove("error");
  elements.artifactToolStatus.textContent = "Building the Office file locally…";
  elements.generateArtifactButton.disabled = true;
  try {
    const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/artifact-tools/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ kind, title, subtitle: elements.artifactToolSubtitle.value.trim(), sections, source_ids: [] }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || "Could not create the teaching file");
    state.artifacts.unshift(result.artifact);
    renderArtifacts();
    closeDialog(elements.artifactToolDialog);
    openDialog(elements.artifactsDialog);
    elements.artifactToolForm.reset();
  } catch (error) {
    elements.artifactToolStatus.classList.add("error");
    elements.artifactToolStatus.textContent = error.message;
  } finally {
    elements.generateArtifactButton.disabled = false;
  }
}

function showDecision(decision, originMessageId = null, savedSelection = null) {
  const recordedSelection = savedSelection || selectionForDecision(originMessageId);
  elements.decisionTitle.textContent = recordedSelection ? "Review decision" : "Choose the next direction";
  elements.decisionQuestion.textContent = decision.question;
  elements.decisionStatus.replaceChildren();
  elements.decisionStatus.hidden = !recordedSelection;
  if (recordedSelection) {
    const mark = document.createElement("span"); mark.setAttribute("aria-hidden", "true"); mark.textContent = "✓";
    const copy = document.createElement("span"); const label = document.createElement("small"); label.textContent = "Recorded choice"; const value = document.createElement("strong"); value.textContent = recordedSelection.selected_label; copy.append(label, value); elements.decisionStatus.append(mark, copy);
  }
  elements.decisionOptions.replaceChildren();
  decision.options.forEach((option, index) => {
    const isSelected = Boolean(recordedSelection && (recordedSelection.selected_value === option.value || recordedSelection.selected_label === option.label));
    const button = document.createElement("button"); button.type = "button"; button.className = `decision-option${isSelected ? " selected" : ""}`; button.setAttribute("aria-pressed", String(isSelected));
    const number = document.createElement("span"); number.className = "decision-option-index"; number.textContent = isSelected ? "✓" : String(index + 1);
    const copy = document.createElement("span"); const label = document.createElement("strong"); label.textContent = option.label; const description = document.createElement("small"); description.textContent = option.description; copy.append(label, description);
    const arrow = document.createElement("span"); arrow.className = "decision-option-arrow"; arrow.setAttribute("aria-hidden", "true"); arrow.textContent = "→";
    button.append(number, copy, arrow);
    button.addEventListener("click", async () => {
      if (isSelected) { closeDialog(elements.decisionDialog); return; }
      closeDialog(elements.decisionDialog);
      state.pendingDecision = null;
      await requestDesignPartner(
        option.value,
        `Selected: ${option.label}\n${option.description}`,
        decision.skill_profile || "auto",
        {
          origin_message_id: originMessageId,
          question: decision.question,
          selected_label: option.label,
          selected_value: option.value,
        },
      );
    });
    elements.decisionOptions.appendChild(button);
  });
  elements.customDecisionButton.onclick = () => {
    state.pendingDecision = { origin_message_id: originMessageId, question: decision.question };
    closeDialog(elements.decisionDialog);
    elements.messageInput.placeholder = "Describe the direction you want to take…";
    elements.messageInput.focus();
  };
  elements.customDecisionButton.textContent = recordedSelection ? "Suggest a different choice" : "I want to suggest something else";
  openDialog(elements.decisionDialog);
}

function renderProjectList() {
  elements.projectList.replaceChildren();
  state.projects.forEach((project) => {
    const item = document.createElement("div"); item.className = `project-list-item${project.id === state.projectId ? " current" : ""}`;
    const open = document.createElement("button"); open.type = "button"; open.className = "project-open";
    const name = document.createElement("strong"); name.textContent = project.name;
    const meta = document.createElement("small"); meta.textContent = `${project.message_count || 0} messages · ${project.artifact_count || 0} artifacts · ${formatDate(project.updated_at)}`;
    open.append(name, meta); open.addEventListener("click", async () => { closeDialog(elements.projectsDialog); await loadProject(project.id); });
    open.setAttribute("aria-label", `Open project ${project.name}`);
    const remove = document.createElement("button"); remove.type = "button"; remove.className = "project-delete"; remove.textContent = "Delete"; remove.setAttribute("aria-label", `Delete project ${project.name}`); remove.addEventListener("click", () => deleteProject(project));
    item.append(open, remove); elements.projectList.appendChild(item);
  });
}

async function refreshProjects() {
  const response = await fetch("/api/projects");
  if (!response.ok) throw new Error("Could not load projects");
  state.projects = (await response.json()).projects || [];
  renderProjectList();
}

async function createProject() {
  const response = await fetch("/api/projects", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: "Untitled course project" }) });
  const result = await response.json();
  if (!response.ok) throw new Error(result.detail || "Could not create project");
  await refreshProjects();
  await loadProject(result.project.id);
  closeDialog(elements.projectsDialog);
  elements.projectName.focus(); elements.projectName.select();
}

async function deleteProject(project) {
  if (!window.confirm(`Delete “${project.name}” and its local conversation, artifacts, and uploaded sources?`)) return;
  const response = await fetch(`/api/projects/${encodeURIComponent(project.id)}`, { method: "DELETE" });
  if (!response.ok) { const result = await response.json(); throw new Error(result.detail || "Could not delete project"); }
  if (project.id === state.projectId) {
    localStorage.removeItem("courseDesignProjectId");
    await refreshProjects();
    if (state.projects.length) await loadProject(state.projects[0].id); else await createProject();
  } else await refreshProjects();
}

async function loadProject(projectId) {
  const response = await fetch(`/api/projects/${encodeURIComponent(projectId)}`);
  const workspace = await response.json();
  if (!response.ok) throw new Error(workspace.detail || "Could not load project");
  state.projectId = projectId; state.project = workspace.project; state.messages = workspace.messages || []; state.artifacts = workspace.artifacts || []; state.sources = workspace.sources || []; state.stateFiles = workspace.state_files || [];
  localStorage.setItem("courseDesignProjectId", projectId);
  state.pendingDecision = null;
  elements.messageInput.placeholder = "Add a direction, question, or correction…";
  fillProjectForm(state.project); renderSources(); renderArtifacts(); renderStateFiles(); renderConversation(); updateExportLinks(); renderProjectList(); setSaveStatus("Saved locally");
  if (window.matchMedia("(max-width: 820px)").matches) setMobileBriefOpen(!state.messages.length && !state.project.outcome);
  else setMobileBriefOpen(false);
}

function setExportLinks(base) {
  if (!state.projectId) return;
  elements.exportMarkdown.href = `${base}?format=markdown`;
  elements.exportHtml.href = `${base}?format=html`;
  elements.exportJson.href = `${base}?format=json`;
  [elements.exportMarkdown, elements.exportHtml, elements.exportJson].forEach((link) => link.setAttribute("download", ""));
}

function updateExportLinks() {
  if (!state.projectId) return;
  const base = `/api/projects/${encodeURIComponent(state.projectId)}/export`;
  setExportLinks(base);
  elements.exportZip.href = `${base}?format=zip`;
  elements.exportZip.setAttribute("download", "");
}

function openFullExport() {
  elements.exportEyebrow.textContent = "Portable project";
  elements.exportDialogTitle.textContent = "Download project";
  elements.exportDescription.textContent = "Export the project context, conversation, SKILL decision trace, and generated teaching files.";
  elements.exportZip.hidden = false;
  elements.exportMarkdownDescription.textContent = "Editable conversation, trace, and artifact index";
  elements.exportHtmlDescription.textContent = "Rendered project for viewing or printing";
  elements.exportJsonDescription.textContent = "Project backup with history and trace metadata";
  updateExportLinks();
  openDialog(elements.exportDialog);
}

function openMessageExport(message) {
  elements.exportEyebrow.textContent = message.role === "user" ? "Instructor message" : "Design partner response";
  elements.exportDialogTitle.textContent = "Download selected message";
  elements.exportDescription.textContent = "Export only this message. Assistant-message exports include the SKILL profile, loaded files, and runtime fingerprint.";
  elements.exportZip.hidden = true;
  elements.exportMarkdownDescription.textContent = "Editable copy of this message";
  elements.exportHtmlDescription.textContent = "Rendered message for viewing or printing";
  elements.exportJsonDescription.textContent = "Message content and trace metadata";
  const base = `/api/projects/${encodeURIComponent(state.projectId)}/messages/${encodeURIComponent(message.id)}/download`;
  setExportLinks(base);
  openDialog(elements.exportDialog);
}

function appendTraceDetail(parent, labelText, value) {
  if (!value) return;
  const row = document.createElement("div"); row.className = "trace-detail";
  const label = document.createElement("strong"); label.textContent = labelText;
  const content = document.createElement("span"); content.textContent = value;
  row.append(label, content); parent.appendChild(row);
}

function renderTrace(trace) {
  elements.traceSummary.replaceChildren();
  const labels = [
    ["Events", trace.summary.events], ["SKILL routes", trace.summary.routed_responses],
    ["Questions", trace.summary.decisions], ["Answers", trace.summary.answers],
    ["Artifacts", trace.summary.artifacts],
  ];
  labels.forEach(([labelText, value]) => {
    const card = document.createElement("div"); card.className = "trace-stat";
    const count = document.createElement("strong"); count.textContent = String(value);
    const label = document.createElement("span"); label.textContent = labelText;
    card.append(count, label); elements.traceSummary.appendChild(card);
  });

  elements.traceTimeline.replaceChildren();
  if (!trace.events.length) {
    const empty = document.createElement("div"); empty.className = "empty-list"; empty.textContent = "The trace will appear after the first project message."; elements.traceTimeline.appendChild(empty); return;
  }
  trace.events.forEach((event) => {
    const article = document.createElement("article"); article.className = `trace-event trace-${event.type}`;
    const marker = document.createElement("div"); marker.className = "trace-marker"; marker.setAttribute("aria-hidden", "true"); marker.textContent = String(event.sequence);
    const card = document.createElement("div"); card.className = "trace-event-card";
    const header = document.createElement("div"); header.className = "trace-event-header";
    const titleGroup = document.createElement("div");
    const type = document.createElement("span"); type.className = "trace-type"; type.textContent = event.type;
    const title = document.createElement("h3"); title.textContent = event.title;
    titleGroup.append(type, title);
    const date = document.createElement("time"); date.textContent = formatDate(event.created_at); header.append(titleGroup, date);
    card.appendChild(header);

    if (event.decision) {
      const question = document.createElement("p"); question.className = "trace-question"; question.textContent = event.decision.question; card.appendChild(question);
      const options = document.createElement("div"); options.className = "trace-options";
      event.decision.options?.forEach((option) => { const chip = document.createElement("span"); chip.textContent = option.label; options.appendChild(chip); });
      card.appendChild(options);
    } else if (event.auto_decision) {
      const answer = document.createElement("div"); answer.className = "trace-answer auto";
      const prompt = document.createElement("span"); prompt.textContent = event.auto_decision.question;
      const selected = document.createElement("strong"); selected.textContent = `Auto selected: ${event.auto_decision.selected_label}`;
      const rationale = document.createElement("p"); rationale.textContent = event.auto_decision.rationale;
      answer.append(prompt, selected, rationale); card.appendChild(answer);
    } else if (event.decision_trace) {
      const answer = document.createElement("div"); answer.className = "trace-answer";
      const prompt = document.createElement("span"); prompt.textContent = event.decision_trace.question;
      const selected = document.createElement("strong"); selected.textContent = `Selected: ${event.decision_trace.selected_label}`;
      answer.append(prompt, selected); card.appendChild(answer);
      appendTraceDetail(card, "Linked question", event.decision_trace.origin_message_id || "Legacy decision");
    } else {
      const preview = document.createElement("p"); preview.className = "trace-preview"; preview.textContent = event.content_preview; card.appendChild(preview);
    }

    if (event.skill_runtime) {
      const runtime = document.createElement("section"); runtime.className = "trace-runtime";
      const runtimeTitle = document.createElement("strong"); runtimeTitle.textContent = "SKILL runtime route"; runtime.appendChild(runtimeTitle);
      appendTraceDetail(runtime, "Profile", event.skill_runtime.profile);
      appendTraceDetail(runtime, "Mode", event.skill_runtime.mode);
      appendTraceDetail(runtime, "Loaded files", event.skill_runtime.loaded_files?.join(" → "));
      appendTraceDetail(runtime, "Fingerprint", event.skill_runtime.fingerprint);
      card.appendChild(runtime);
    }
    appendTraceDetail(card, "Sources", event.sources_used?.join(", "));
    const linkedArtifacts = event.artifacts?.length ? event.artifacts : event.artifact ? [event.artifact] : [];
    linkedArtifacts.forEach((artifact) => appendTraceDetail(card, "Artifact", `${artifact.title} · ${artifact.kind}`));
    article.append(marker, card); elements.traceTimeline.appendChild(article);
  });
}

async function openSkillTrace() {
  elements.traceSummary.replaceChildren();
  elements.traceTimeline.textContent = "Loading the persistent project trace…";
  openDialog(elements.traceDialog);
  try {
    const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/trace`);
    const trace = await response.json();
    if (!response.ok) throw new Error(trace.detail || "Could not load the SKILL trace");
    renderTrace(trace);
  } catch (error) {
    elements.traceTimeline.textContent = error.message;
  }
}

function scheduleProjectSave() {
  if (state.loadingProject || !state.projectId) return;
  clearTimeout(state.saveTimer); setSaveStatus("Saving…", "saving");
  state.saveTimer = setTimeout(saveProject, 450);
}

async function saveProject() {
  try {
    const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(projectPayload()) });
    const result = await response.json(); if (!response.ok) throw new Error(result.detail || "Could not save project");
    state.project = result.project; elements.currentProjectName.textContent = state.project.name; setSaveStatus("Saved locally"); await refreshProjects(); updateControls();
  } catch (error) { setSaveStatus("Save failed", "error"); }
}

async function uploadProjectFiles(files) {
  const selectedFiles = Array.from(files || []); if (!selectedFiles.length || !state.projectId) return;
  elements.uploadStatus.classList.remove("error");
  if (!elements.dataAcknowledgment.checked) { elements.uploadStatus.classList.add("error"); elements.uploadStatus.textContent = "Confirm the data statement before uploading."; elements.dataAcknowledgment.focus(); return; }
  const oversized = selectedFiles.find((file) => file.size > state.uploadMaxBytes);
  if (oversized) { elements.uploadStatus.classList.add("error"); elements.uploadStatus.textContent = `${oversized.name} exceeds the ${formatBytes(state.uploadMaxBytes)} limit.`; return; }
  const data = new FormData(); selectedFiles.forEach((file) => data.append("files", file)); data.append("data_classification_ack", "true");
  elements.uploadStatus.textContent = `Processing ${selectedFiles.length} ${selectedFiles.length === 1 ? "file" : "files"} locally…`;
  try {
    const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/sources`, { method: "POST", body: data });
    const result = await response.json(); if (!response.ok) throw new Error(result.detail || "Upload failed");
    const workspace = await (await fetch(`/api/projects/${encodeURIComponent(state.projectId)}`)).json(); state.sources = workspace.sources || []; renderSources();
    elements.uploadStatus.textContent = result.errors?.length ? `${result.sources.length} uploaded; ${result.errors.length} could not be processed.` : `${result.sources.length} ${result.sources.length === 1 ? "source is" : "sources are"} ready.`;
    elements.uploadStatus.classList.toggle("error", Boolean(result.errors?.length));
  } catch (error) { elements.uploadStatus.classList.add("error"); elements.uploadStatus.textContent = error.message; }
  finally { elements.sourceFiles.value = ""; }
}

async function removeSource(source) {
  if (!window.confirm(`Remove ${source.filename} from this project?`)) return;
  const response = await fetch(`/api/projects/${encodeURIComponent(state.projectId)}/sources/${source.source_id}`, { method: "DELETE" });
  const result = await response.json(); if (!response.ok) { elements.uploadStatus.classList.add("error"); elements.uploadStatus.textContent = result.detail || "Could not remove source"; return; }
  state.sources = state.sources.filter((item) => item.source_id !== source.source_id); renderSources(); elements.uploadStatus.textContent = `${source.filename} was removed.`;
}

function updateSettingsTargetPath() {
  const target = state.environment?.targets?.find((item) => item.id === elements.settingsTarget.value);
  elements.settingsPath.textContent = target ? `${target.path}${target.exists ? " · existing file" : " · will be created"}` : "";
}

function renderEnvironmentSettings() {
  const configuration = state.environment;
  elements.settingsTarget.replaceChildren();
  configuration.targets.forEach((target) => {
    const option = document.createElement("option"); option.value = target.id; option.textContent = target.label; if (target.id === "current") option.selected = true; elements.settingsTarget.appendChild(option);
  });
  updateSettingsTargetPath();
  elements.settingsFields.replaceChildren();
  const groups = new Map();
  configuration.fields.forEach((field) => {
    if (!groups.has(field.group)) groups.set(field.group, []);
    groups.get(field.group).push(field);
  });
  groups.forEach((fields, groupName) => {
    const fieldset = document.createElement("fieldset"); fieldset.className = "settings-group";
    const legend = document.createElement("legend"); legend.textContent = groupName;
    const grid = document.createElement("div"); grid.className = "settings-grid";
    fields.forEach((field) => {
      const label = document.createElement("label"); label.className = "setting-field";
      const title = document.createElement("span"); const name = document.createElement("span"); name.textContent = field.label;
      const source = document.createElement("span"); source.className = "setting-source"; source.textContent = field.source || "default"; title.append(name, source);
      const input = document.createElement("input"); input.className = "setting-input"; input.dataset.envName = field.name; input.dataset.secret = String(field.secret); input.type = field.secret ? "password" : "text"; input.autocomplete = "off";
      if (field.secret) input.placeholder = field.configured ? "Configured — enter a new value to replace" : "Enter API key"; else input.value = field.value || "";
      const hint = document.createElement("small"); hint.className = "setting-hint"; hint.textContent = `${field.name}${field.restart ? " · restart required after changing" : ""}`;
      label.append(title, input, hint);
      if (field.secret && field.configured) {
        const secretActions = document.createElement("label"); secretActions.className = "secret-actions";
        const clear = document.createElement("input"); clear.type = "checkbox"; clear.className = "setting-clear"; clear.dataset.envName = field.name;
        const clearText = document.createElement("span"); clearText.textContent = "Remove this key from the selected file"; secretActions.append(clear, clearText); label.appendChild(secretActions);
      }
      grid.appendChild(label);
    });
    fieldset.append(legend, grid); elements.settingsFields.appendChild(fieldset);
  });
}

async function loadEnvironmentSettings() {
  const response = await fetch("/api/settings"); const result = await response.json();
  if (!response.ok) throw new Error(result.detail || "Could not load environment settings");
  state.environment = result; renderEnvironmentSettings();
}

async function saveEnvironmentSettings(event) {
  event.preventDefault();
  elements.settingsStatus.classList.remove("error"); elements.settingsStatus.textContent = "Saving…"; elements.saveSettingsButton.disabled = true;
  const values = {};
  const inputs = elements.settingsFields.querySelectorAll(".setting-input");
  inputs.forEach((input) => {
    const isSecret = input.dataset.secret === "true";
    const clear = elements.settingsFields.querySelector(`.setting-clear[data-env-name="${CSS.escape(input.dataset.envName)}"]`);
    if (clear?.checked) values[input.dataset.envName] = null;
    else if (!isSecret || input.value.trim()) values[input.dataset.envName] = input.value;
  });
  try {
    const response = await fetch("/api/settings", { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ target: elements.settingsTarget.value, values }) });
    const result = await response.json(); if (!response.ok) throw new Error(result.detail || "Could not save configuration");
    state.environment = result.configuration; renderEnvironmentSettings(); await loadConfiguration();
    elements.settingsStatus.textContent = result.restart_required ? "Saved. Restart the app to apply server or port changes." : `Saved to ${result.path}`;
  } catch (error) { elements.settingsStatus.classList.add("error"); elements.settingsStatus.textContent = error.message; }
  finally { elements.saveSettingsButton.disabled = false; }
}

async function loadConfiguration() {
  try {
    const response = await fetch("/api/config"); if (!response.ok) throw new Error("Configuration unavailable"); const config = await response.json();
    state.skillRuntime = config.skill_runtime; state.uploadMaxBytes = config.upload_max_bytes || state.uploadMaxBytes;
    state.configured = Boolean(config.api_key_configured && config.model_id && config.skill_runtime?.fingerprint);
    elements.modelStatus.classList.toggle("ready", state.configured); elements.modelStatus.classList.toggle("error", !state.configured);
    setModelStatus(state.configured ? `${config.model_id} ready` : "Model configuration needed");
    elements.skillRuntimeLabel.textContent = config.skill_runtime ? `Skill active · ${config.skill_runtime.loaded_files.join(" + ")} · gpt-oss checks on · ${config.skill_runtime.fingerprint}` : "Skill runtime unavailable";
  } catch (error) { state.configured = false; elements.modelStatus.classList.add("error"); setModelStatus("Service unavailable"); elements.skillRuntimeLabel.textContent = "Skill runtime unavailable"; }
  updateControls();
}

async function requestDesignPartner(userText, displayText, skillProfile = "auto", decisionTrace = null) {
  state.busy = true; state.artifactToolNotice = null; updateControls();
  const optimisticUser = createMessage({ role: "user", content: displayText, created_at: new Date().toISOString() });
  const loading = createLoadingMessage();
  try {
    const response = await fetch("/api/chat", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ messages: [{ role: "user", content: userText }], display_content: displayText, project_id: state.projectId, skill_profile: skillProfile, decision_trace: decisionTrace }) });
    const result = await response.json(); if (!response.ok) throw new Error(typeof result.detail === "string" ? result.detail : "The model request was not successful."); if (!result.content) throw new Error("The model returned an empty response.");
    loading.remove(); optimisticUser.remove();
    if (result.artifact) state.artifacts.unshift(result.artifact);
    const assistant = result.assistant_message || { role: "assistant", content: result.content, html: result.html, decision: result.decision, skill_runtime: result.skill_runtime, sources_used: result.sources_used, created_at: new Date().toISOString() };
    const persistedUser = result.user_message || { role: "user", content: displayText, created_at: new Date().toISOString(), decision_trace: decisionTrace };
    state.messages.push(persistedUser, assistant);
    createMessage(persistedUser);
    createMessage(assistant, { artifact: result.artifact }); renderArtifacts();
    if (result.artifact_tool_error) state.artifactToolNotice = `The response was saved, but the Office file was not created: ${result.artifact_tool_error}`;
    if (result.state_file) await refreshStateFiles();
    if (result.skill_runtime) { state.skillRuntime = result.skill_runtime; elements.skillRuntimeLabel.textContent = `Skill active · ${result.skill_runtime.profile} · gpt-oss checks on · ${result.skill_runtime.fingerprint}`; }
    elements.emptyState.classList.add("compact"); state.started = true; elements.sessionLabel.textContent = `${state.project.mode} project · ${state.messages.length} saved messages`; await refreshProjects();
    if (result.decision && state.project?.mode !== "Auto") showDecision(result.decision, assistant.id);
  } catch (error) { loading.remove(); createMessage({ role: "assistant", content: error.message }, { error: true }); }
  finally { state.busy = false; elements.messageInput.value = ""; updateControls(); if (state.started && !elements.decisionDialog.open) elements.messageInput.focus(); }
}

async function startSession() {
  const context = projectPayload();
  if (!context.outcome) { elements.learningOutcome.closest(".field").classList.add("invalid"); elements.outcomeHint.textContent = "A learning outcome is required before starting."; elements.learningOutcome.focus(); return; }
  await saveProject();
  elements.learningOutcome.closest(".field").classList.remove("invalid"); elements.outcomeHint.textContent = "Use an observable action such as analyze, compare, design, or justify."; state.started = true;
  const prompt = `Begin a ${context.mode} course-design collaboration.\nCourse or module: ${context.course_name || "Not yet named"}\nLevel: ${context.level}\nClass time: ${context.class_time}\nLearning outcome: ${context.outcome}\nInstructor-provided constraints or source notes: ${context.notes || "No additional constraints or source notes provided"}`;
  const display = `${context.course_name || context.name}\nOutcome: ${context.outcome}\nMode: ${context.mode}`;
  await requestDesignPartner(prompt, display, "establish");
}

async function initialize() {
  elements.footerYear.textContent = String(new Date().getFullYear());
  await loadConfiguration();
  await refreshProjects();
  const requestedId = new URLSearchParams(window.location.search).get("project");
  const savedId = localStorage.getItem("courseDesignProjectId");
  const initial = state.projects.find((project) => project.id === requestedId) || state.projects.find((project) => project.id === savedId) || state.projects[0];
  if (initial) await loadProject(initial.id); else await createProject();
}

elements.sendButton.addEventListener("click", () => {
  if (!state.started) { startSession(); return; }
  const message = elements.messageInput.value.trim();
  const pending = state.pendingDecision;
  state.pendingDecision = null;
  elements.messageInput.placeholder = "Add a direction, question, or correction…";
  const trace = pending ? {
    origin_message_id: pending.origin_message_id,
    question: pending.question,
    selected_label: "Custom response",
    selected_value: message,
  } : null;
  requestDesignPartner(message, message, "auto", trace);
});
elements.messageInput.addEventListener("input", updateControls);
elements.messageInput.addEventListener("keydown", (event) => { if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) { event.preventDefault(); if (!elements.sendButton.disabled) elements.sendButton.click(); } });
elements.projectForm.addEventListener("submit", (event) => event.preventDefault());
elements.projectForm.addEventListener("input", (event) => { if (event.target === elements.learningOutcome) { elements.learningOutcome.closest(".field").classList.remove("invalid"); elements.outcomeHint.textContent = "Use an observable action such as analyze, compare, design, or justify."; updateControls(); } scheduleProjectSave(); });
elements.projectForm.addEventListener("change", scheduleProjectSave);
elements.projectSwitcher.addEventListener("click", async () => { await refreshProjects(); openDialog(elements.projectsDialog); });
elements.mobileBriefToggle.addEventListener("click", () => setMobileBriefOpen(true));
elements.mobileBriefClose.addEventListener("click", () => setMobileBriefOpen(false));
window.matchMedia("(max-width: 820px)").addEventListener("change", (event) => { if (!event.matches) setMobileBriefOpen(false); });
document.addEventListener("keydown", (event) => { if (event.key === "Escape" && elements.projectPanel.classList.contains("mobile-open")) setMobileBriefOpen(false); });
elements.newProjectButton.addEventListener("click", createProject);
elements.artifactsButton.addEventListener("click", () => { renderArtifacts(); openDialog(elements.artifactsDialog); });
elements.artifactToolButton.addEventListener("click", () => { closeDialog(elements.artifactsDialog); elements.artifactToolStatus.textContent = ""; openDialog(elements.artifactToolDialog); elements.artifactToolTitle.focus(); });
elements.artifactToolForm.addEventListener("submit", generateTeachingFile);
elements.stateButton.addEventListener("click", async () => {
  elements.validationResults.replaceChildren();
  closeStateEditor();
  await Promise.all([refreshStateFiles(), loadStateTemplates()]);
  openDialog(elements.stateDialog);
});
elements.runValidatorsButton.addEventListener("click", runValidators);
elements.stateTemplatePicker.addEventListener("change", (event) => { if (event.target.value) openStateEditor(event.target.value, true); });
elements.stateEditorCancel.addEventListener("click", closeStateEditor);
elements.stateEditorForm.addEventListener("submit", saveStateFile);
elements.traceButton.addEventListener("click", openSkillTrace);
elements.exportButton.addEventListener("click", openFullExport);
elements.settingsButton.addEventListener("click", async () => {
  elements.settingsStatus.textContent = "";
  try { await loadEnvironmentSettings(); openDialog(elements.settingsDialog); }
  catch (error) { elements.composerHelp.textContent = error.message; }
});
elements.settingsTarget.addEventListener("change", updateSettingsTargetPath);
elements.settingsForm.addEventListener("submit", saveEnvironmentSettings);
elements.sourceFiles.addEventListener("change", () => uploadProjectFiles(elements.sourceFiles.files));
elements.dataAcknowledgment.addEventListener("change", () => { if (elements.dataAcknowledgment.checked) { elements.uploadStatus.classList.remove("error"); elements.uploadStatus.textContent = "Data statement confirmed for this browser session."; } });
elements.dropZone.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); elements.sourceFiles.click(); } });
["dragenter", "dragover"].forEach((name) => elements.dropZone.addEventListener(name, (event) => { event.preventDefault(); elements.dropZone.classList.add("dragging"); }));
["dragleave", "drop"].forEach((name) => elements.dropZone.addEventListener(name, (event) => { event.preventDefault(); elements.dropZone.classList.remove("dragging"); }));
elements.dropZone.addEventListener("drop", (event) => uploadProjectFiles(event.dataTransfer.files));
document.querySelectorAll("[data-close]").forEach((button) => button.addEventListener("click", () => closeDialog(document.getElementById(button.dataset.close))));
document.querySelectorAll("dialog").forEach((dialog) => dialog.addEventListener("click", (event) => { if (event.target === dialog) closeDialog(dialog); }));

initialize().catch((error) => { elements.modelStatus.classList.add("error"); setModelStatus("Workspace unavailable"); elements.composerHelp.textContent = error.message; });
