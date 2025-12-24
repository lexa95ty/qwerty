// prompts.js
(function () {
  const PC = window.PromptConfig;
  const $ = (id) => document.getElementById(id);

  // -----------------------------
  // DOM: Presets
  // -----------------------------
  const presetSelect = $("presetSelect");
  const presetName = $("presetName");
  const presetDesc = $("presetDesc");
  const activeBadge = $("activeBadge");
  const makeActiveBtn = $("makeActiveBtn");
  const deletePresetBtn = $("deletePresetBtn");
  const createName = $("createName");
  const createBase = $("createBase");
  const createBtn = $("createBtn");

  // -----------------------------
  // DOM: Global AI settings
  // -----------------------------
  const autoAddContext = $("autoAddContext");
  const autoAddSummary = $("autoAddSummary");
  const autoAddTail = $("autoAddTail");
  const tailChars = $("tailChars");
  const updateSummaryAfterStage = $("updateSummaryAfterStage");
  const updateSummaryAfterSection = $("updateSummaryAfterSection");

  // -----------------------------
  // DOM: Import/Export
  // -----------------------------
  const exportBtn = $("exportBtn");
  const importBtn = $("importBtn");
  const importFile = $("importFile");

  // -----------------------------
  // DOM: Variables
  // -----------------------------
  const varSearch = $("varSearch");
  const copyHintBtn = $("copyHintBtn");
  const varButtons = $("varButtons");
  const newVarKey = $("newVarKey");
  const addVarBtn = $("addVarBtn");
  const userVars = $("userVars");

  // -----------------------------
  // DOM: Tabs + header
  // -----------------------------
  const contextTitle = $("contextTitle");
  const contextSubtitle = $("contextSubtitle");
  const saveBtn = $("saveBtn");
  const testBtn = $("testBtn");

  const tabStages = $("tabStages");
  const tabPrompts = $("tabPrompts");
  const tabStyle = $("tabStyle"); // may be absent in old HTML

  const panelStages = $("panelStages");
  const panelPrompts = $("panelPrompts");
  const panelStyle = $("panelStyle"); // may be absent in old HTML

  // -----------------------------
  // DOM: Stages
  // -----------------------------
  const chaptersCount = $("chaptersCount");
  const sectionsPerChapter = $("sectionsPerChapter");
  const syncChaptersBtn = $("syncChaptersBtn");
  const stagesList = $("stagesList");
  const addStageBtn = $("addStageBtn");
  const addStagePanel = $("addStagePanel");
  const newStageType = $("newStageType");
  const newStageLabel = $("newStageLabel");
  const newStageChapterWrap = $("newStageChapterWrap");
  const newStageChapter = $("newStageChapter");
  const newStagePromptWrap = $("newStagePromptWrap");
  const newStagePrompt = $("newStagePrompt");
  const confirmAddStageBtn = $("confirmAddStageBtn");
  const cancelAddStageBtn = $("cancelAddStageBtn");

  // -----------------------------
  // DOM: Prompts
  // -----------------------------
  const promptSelect = $("promptSelect");
  const duplicatePromptBtn = $("duplicatePromptBtn");
  const renamePromptBtn = $("renamePromptBtn");
  const deletePromptBtn = $("deletePromptBtn");

  const newPromptId = $("newPromptId");
  const newPromptTitle = $("newPromptTitle");
  const newPromptBase = $("newPromptBase");
  const createPromptBtn = $("createPromptBtn");

  const modelEl = $("model");
  const modeEl = $("mode");
  const temperatureEl = $("temperature");
  const maxTokensEl = $("maxTokens");
  const systemEl = $("system");
  const userEl = $("user");
  const testOut = $("testOut");

  // -----------------------------
  // DOM: DOCX Style (optional; only if prompts.html includes it)
  // -----------------------------
  const docxStyleResetBtn = $("docxStyleResetBtn");
  const docxFontName = $("docxFontName");
  const docxFontSizePt = $("docxFontSizePt");
  const docxMarginTop = $("docxMarginTop");
  const docxMarginRight = $("docxMarginRight");
  const docxMarginBottom = $("docxMarginBottom");
  const docxMarginLeft = $("docxMarginLeft");
  const docxLineSpacing = $("docxLineSpacing");
  const docxCustomLine = $("docxCustomLine");
  const docxFirstLineIndent = $("docxFirstLineIndent");
  const docxTocEnabled = $("docxTocEnabled");

  const docxPageNumbersEnabled = $("docxPageNumbersEnabled");
  const docxHideOnTitle = $("docxHideOnTitle");
  const docxPageNumbersPosition = $("docxPageNumbersPosition");
  const docxPageNumbersStartAt = $("docxPageNumbersStartAt");

  const docxStyleJson = $("docxStyleJson");
  const docxStyleFormatBtn = $("docxStyleFormatBtn");
  const docxStyleApplyBtn = $("docxStyleApplyBtn");
  const docxStyleImportBtn = $("docxStyleImportBtn");
  const docxStyleImportFile = $("docxStyleImportFile");
  const docxStyleCopyBtn = $("docxStyleCopyBtn");

  const hasStyleUI = !!(tabStyle && panelStyle);
  const hasDocxStyleControls = !!(docxStyleJson && docxFontName && docxFontSizePt);

  // -----------------------------
  // State
  // -----------------------------
  let allPresets = null;
  let currentPresetId = null;
  let activeTextarea = null;

  let styleSyncLock = false; // prevents feedback loops while syncing UI <-> JSON

  // -----------------------------
  // Helpers
  // -----------------------------
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
    }[c]));
  }

  function clampInt(v, min, max, def) {
    const n = parseInt(String(v || ""), 10);
    if (!Number.isFinite(n)) return def;
    return Math.max(min, Math.min(max, n));
  }

  function clampNum(v, def) {
    const n = Number(v);
    return Number.isFinite(n) ? n : def;
  }

  function isValidId(id) {
    return /^[a-zA-Z0-9_]+$/.test(id || "");
  }

  function deepClone(obj) {
    try { return PC.deepClone ? PC.deepClone(obj) : JSON.parse(JSON.stringify(obj)); }
    catch (_) { return JSON.parse(JSON.stringify(obj)); }
  }

  function safeJsonParse(s) {
    try { return { ok: true, val: JSON.parse(s) }; }
    catch (e) { return { ok: false, error: e.message }; }
  }

  function prettyJson(v) {
    return JSON.stringify(v, null, 2);
  }

  function builtinVarsList() {
    // Supports:
    // - PC.BUILTIN_VARS as array [{key,label},...]
    // - PC.BUILTIN_VARS as object {topic:"", ...}
    const bv = PC.BUILTIN_VARS;
    if (Array.isArray(bv)) return bv;
    if (bv && typeof bv === "object") {
      return Object.keys(bv).map((k) => ({ key: k, label: k }));
    }
    return [];
  }

  function persist() {
    PC.saveAllPresets(allPresets);
    refreshPresetDropdown();
    updateActiveBadge();
  }

  function getPreset() {
    return allPresets[currentPresetId];
  }

  function getGen(p) {
    p.settings = p.settings || {};
    p.settings.generation = p.settings.generation || { chaptersCount: 2, sectionsPerChapter: 4, defaultPromptIds: {}, stages: [] };
    const gen = p.settings.generation;

    gen.defaultPromptIds = gen.defaultPromptIds || {
      plan: "plan",
      section: "section",
      intro: "intro",
      conclusion: "conclusion",
      bibliography: "bibliography",
      summary: "summary",
      appendix: "appendix"
    };

    gen.chaptersCount = clampInt(gen.chaptersCount, 1, 10, 2);
    gen.sectionsPerChapter = clampInt(gen.sectionsPerChapter, 1, 12, 4);

    gen.stages = Array.isArray(gen.stages) ? gen.stages : [];
    if (!gen.stages.length) gen.stages = PC.defaultPipeline(gen.chaptersCount, gen.defaultPromptIds);

    return gen;
  }

  function getPromptIds(p) {
    p.prompts = p.prompts || {};
    return Object.keys(p.prompts).sort((a, b) => a.localeCompare(b));
  }

  // -----------------------------
  // DOCX Style helpers (per preset)
  // -----------------------------
  function fallbackDefaultDocxStyle() {
    const now = new Date().toISOString();
    return {
      meta: {
        schemaVersion: (PC.DOCX_STYLE_SCHEMA_VERSION || "projectai-docx-style/v1"),
        name: "Default",
        author: "ProjectAI",
        description: "Default docx style",
        createdAt: now,
        updatedAt: now
      },
      page: {
        size: "A4",
        orientation: "portrait",
        marginsCm: { top: 2, right: 1.5, bottom: 2, left: 3 },
        headerDistanceCm: 1.25,
        footerDistanceCm: 1.25,
        gutterCm: 0
      },
      base: {
        font: { name: "Times New Roman", sizePt: 14, color: "#000000", lang: "ru-RU" },
        paragraph: { alignment: "justify", lineSpacing: "1.5", customLine: 1.35, firstLineIndentCm: 1.25, spacingBeforePt: 0, spacingAfterPt: 0 }
      },
      toc: { enabled: true, showPageNumbers: true, tabLeader: "dots" },
      headersFooters: {
        differentFirstPage: true,
        differentOddEven: false,
        pageNumbers: { enabled: true, position: "footer-center", startAt: 1, hideOnTitle: true }
      },
      bibliography: { numbering: { format: "1." }, item: { alignment: "justify", hangingIndentCm: 0, spacingBeforePt: 0, spacingAfterPt: 0 } },
      headings: {
        h1: { font: { sizePt: 14, bold: true, allCaps: true, color: "#000000" }, paragraph: { alignment: "center", spacingBeforePt: 24, spacingAfterPt: 24, pageBreakBefore: true, keepWithNext: true } },
        h2: { font: { sizePt: 14, bold: true, allCaps: false, color: "#000000" }, paragraph: { alignment: "center", spacingBeforePt: 18, spacingAfterPt: 12, pageBreakBefore: false, keepWithNext: true } }
      },
      advanced: { compatibilityJson: "", raw: "" }
    };
  }

  function normalizeDocxStyle(style) {
    // Prefer PromptConfig implementation if present
    if (typeof PC.normalizeDocxStyle === "function") return PC.normalizeDocxStyle(style);
    // Fallback: very light normalization
    const s = deepClone(style || {});
    const def = fallbackDefaultDocxStyle();

    function merge(target, src) {
      if (!src || typeof src !== "object") return target;
      for (const k of Object.keys(src)) {
        const v = src[k];
        if (v && typeof v === "object" && !Array.isArray(v)) {
          target[k] = merge(target[k] && typeof target[k] === "object" ? target[k] : {}, v);
        } else {
          target[k] = v;
        }
      }
      return target;
    }

    const out = merge(def, s);
    out.meta = out.meta || {};
    out.meta.schemaVersion = out.meta.schemaVersion || (PC.DOCX_STYLE_SCHEMA_VERSION || "projectai-docx-style/v1");
    out.meta.updatedAt = new Date().toISOString();
    out.meta.createdAt = out.meta.createdAt || out.meta.updatedAt;

    out.page = out.page || def.page;
    out.page.marginsCm = out.page.marginsCm || def.page.marginsCm;

    out.base = out.base || def.base;
    out.base.font = out.base.font || def.base.font;
    out.base.paragraph = out.base.paragraph || def.base.paragraph;

    out.headersFooters = out.headersFooters || def.headersFooters;
    out.headersFooters.pageNumbers = out.headersFooters.pageNumbers || def.headersFooters.pageNumbers;

    out.toc = out.toc || def.toc;

    return out;
  }

  function defaultDocxStyle() {
    if (typeof PC.defaultDocxStyle === "function") return PC.defaultDocxStyle();
    return fallbackDefaultDocxStyle();
  }

  function ensureDocxStyle(p) {
    p.settings = p.settings || {};
    if (!p.settings.docxStyle) p.settings.docxStyle = defaultDocxStyle();
    p.settings.docxStyle = normalizeDocxStyle(p.settings.docxStyle);
    return p.settings.docxStyle;
  }

  function refreshDocxStyleJson() {
    if (!hasDocxStyleControls) return;
    const p = getPreset();
    const style = ensureDocxStyle(p);
    styleSyncLock = true;
    docxStyleJson.value = prettyJson(style);
    styleSyncLock = false;
  }

  function syncDocxQuickFromStyle() {
    if (!hasDocxStyleControls) return;
    const p = getPreset();
    const style = ensureDocxStyle(p);

    styleSyncLock = true;

    // Base font
    docxFontName.value = style.base?.font?.name || "Times New Roman";
    docxFontSizePt.value = String(style.base?.font?.sizePt ?? 14);

    // Margins
    const m = style.page?.marginsCm || { top: 2, right: 1.5, bottom: 2, left: 3 };
    docxMarginTop.value = String(m.top ?? 2);
    docxMarginRight.value = String(m.right ?? 1.5);
    docxMarginBottom.value = String(m.bottom ?? 2);
    docxMarginLeft.value = String(m.left ?? 3);

    // Paragraph line spacing
    docxLineSpacing.value = String(style.base?.paragraph?.lineSpacing ?? "1.5");
    docxCustomLine.value = String(style.base?.paragraph?.customLine ?? 1.35);

    // First line indent
    docxFirstLineIndent.value = String(style.base?.paragraph?.firstLineIndentCm ?? 1.25);

    // TOC enabled
    const tocEnabled = (style.toc?.enabled !== false);
    docxTocEnabled.checked = !!tocEnabled;

    // Page numbers
    const pn = style.headersFooters?.pageNumbers || {};
    docxPageNumbersEnabled.checked = (pn.enabled !== false);
    docxHideOnTitle.checked = !!pn.hideOnTitle;
    docxPageNumbersPosition.value = String(pn.position || "footer-center");
    docxPageNumbersStartAt.value = String(pn.startAt ?? 1);

    styleSyncLock = false;
  }

  function applyDocxQuickToStyle() {
    if (!hasDocxStyleControls) return;
    if (styleSyncLock) return;

    const p = getPreset();
    const style = ensureDocxStyle(p);

    // Base font
    style.base = style.base || {};
    style.base.font = style.base.font || {};
    style.base.font.name = String(docxFontName.value || "Times New Roman").trim() || "Times New Roman";
    style.base.font.sizePt = clampInt(docxFontSizePt.value, 8, 28, 14);

    // Page margins
    style.page = style.page || {};
    style.page.marginsCm = style.page.marginsCm || { top: 2, right: 1.5, bottom: 2, left: 3 };
    style.page.marginsCm.top = clampNum(docxMarginTop.value, 2);
    style.page.marginsCm.right = clampNum(docxMarginRight.value, 1.5);
    style.page.marginsCm.bottom = clampNum(docxMarginBottom.value, 2);
    style.page.marginsCm.left = clampNum(docxMarginLeft.value, 3);

    // Paragraph settings
    style.base.paragraph = style.base.paragraph || {};
    style.base.paragraph.lineSpacing = String(docxLineSpacing.value || "1.5");
    style.base.paragraph.customLine = clampNum(docxCustomLine.value, 1.35);
    style.base.paragraph.firstLineIndentCm = clampNum(docxFirstLineIndent.value, 1.25);

    // TOC
    style.toc = style.toc || {};
    style.toc.enabled = !!docxTocEnabled.checked;

    // Headers/Footers + Page numbers
    style.headersFooters = style.headersFooters || {};
    style.headersFooters.pageNumbers = style.headersFooters.pageNumbers || {};
    style.headersFooters.pageNumbers.enabled = !!docxPageNumbersEnabled.checked;
    style.headersFooters.pageNumbers.hideOnTitle = !!docxHideOnTitle.checked;
    style.headersFooters.pageNumbers.position = String(docxPageNumbersPosition.value || "footer-center");
    style.headersFooters.pageNumbers.startAt = clampInt(docxPageNumbersStartAt.value, 1, 999, 1);

    // Normalize + save
    p.settings.docxStyle = normalizeDocxStyle(style);
    persist();

    // Update Advanced JSON view
    refreshDocxStyleJson();
  }

  function applyDocxJsonToStyle() {
    if (!hasDocxStyleControls) return;
    const parsed = safeJsonParse(docxStyleJson.value);
    if (!parsed.ok) {
      alert("❌ Ошибка JSON: " + parsed.error);
      return;
    }
    const p = getPreset();
    p.settings = p.settings || {};
    p.settings.docxStyle = normalizeDocxStyle(parsed.val);
    persist();
    syncDocxQuickFromStyle();
    refreshDocxStyleJson();
    alert("✅ DOCX-оформление применено.");
  }

  function formatDocxJson() {
    if (!hasDocxStyleControls) return;
    const parsed = safeJsonParse(docxStyleJson.value);
    if (!parsed.ok) {
      alert("❌ Ошибка JSON: " + parsed.error);
      return;
    }
    docxStyleJson.value = prettyJson(parsed.val);
  }

  async function copyDocxJson() {
    if (!hasDocxStyleControls) return;
    try {
      await navigator.clipboard.writeText(docxStyleJson.value || "");
      alert("✅ JSON скопирован.");
    } catch (e) {
      alert("❌ Не удалось скопировать (браузер запретил).");
    }
  }

  async function importDocxJson(file) {
    if (!hasDocxStyleControls) return;
    const txt = await file.text();
    const parsed = safeJsonParse(txt);
    if (!parsed.ok) {
      alert("❌ Ошибка JSON: " + parsed.error);
      return;
    }
    const p = getPreset();
    p.settings = p.settings || {};
    p.settings.docxStyle = normalizeDocxStyle(parsed.val);
    persist();
    syncDocxQuickFromStyle();
    refreshDocxStyleJson();
    alert("✅ Профиль оформления импортирован.");
  }

  function resetDocxStyle() {
    if (!hasDocxStyleControls) return;
    if (!confirm("Сбросить оформление DOCX к дефолтному для пресета?")) return;
    const p = getPreset();
    p.settings = p.settings || {};
    p.settings.docxStyle = normalizeDocxStyle(defaultDocxStyle());
    persist();
    syncDocxQuickFromStyle();
    refreshDocxStyleJson();
  }

  // -----------------------------
  // UI: Presets dropdown + badge
  // -----------------------------
  function refreshPresetDropdown() {
    const activeId = PC.getActivePresetId();
    presetSelect.innerHTML = "";
    for (const [id, p] of Object.entries(allPresets)) {
      const opt = document.createElement("option");
      opt.value = id;
      opt.textContent = `${p.name || id}${id === activeId ? " (активный)" : ""}`;
      presetSelect.appendChild(opt);
    }
    presetSelect.value = currentPresetId;
  }

  function updateActiveBadge() {
    const activeId = PC.getActivePresetId();
    activeBadge.textContent = (currentPresetId === activeId) ? "активный" : "не активен";
    activeBadge.className = "text-xs px-2.5 py-1 rounded-full border " + (currentPresetId === activeId
      ? "bg-emerald-500/20 border-emerald-300/20 text-emerald-100"
      : "bg-white/5 border-white/10 text-white/70");
  }

  // -----------------------------
  // Tabs
  // -----------------------------
  function showTab(which) {
    // which: "stages" | "prompts" | "style"
    const isStages = which === "stages";
    const isPrompts = which === "prompts";
    const isStyle = which === "style";

    panelStages.classList.toggle("hidden", !isStages);
    panelPrompts.classList.toggle("hidden", !isPrompts);

    if (hasStyleUI) panelStyle.classList.toggle("hidden", !isStyle);

    tabStages.classList.toggle("tab-active", isStages);
    tabStages.classList.toggle("tab-idle", !isStages);

    tabPrompts.classList.toggle("tab-active", isPrompts);
    tabPrompts.classList.toggle("tab-idle", !isPrompts);

    if (hasStyleUI) {
      tabStyle.classList.toggle("tab-active", isStyle);
      tabStyle.classList.toggle("tab-idle", !isStyle);
    }
  }

  // -----------------------------
  // Set preset
  // -----------------------------
  function setPreset(id) {
    currentPresetId = id;
    refreshPresetDropdown();
    updateActiveBadge();

    let p = getPreset();
    let gen = getGen(p);

    presetName.value = p.name || id;
    presetDesc.value = p.description || "";
    contextTitle.textContent = p.name || id;
    contextSubtitle.textContent = p.description || "";

    autoAddContext.checked = !!p.settings.autoAddContext;
    autoAddSummary.checked = !!p.settings.autoAddSummary;
    autoAddTail.checked = !!p.settings.autoAddTail;
    tailChars.value = p.settings.tailChars ?? 12000;
    updateSummaryAfterStage.checked = !!p.settings.updateSummaryAfterEachStage;
    updateSummaryAfterSection.checked = !!p.settings.updateSummaryAfterEachSection;

    chaptersCount.value = gen.chaptersCount;
    sectionsPerChapter.value = gen.sectionsPerChapter;

    // Ensure prompts exist and migrated
    p.prompts = p.prompts || (PC.defaultPrompts ? PC.defaultPrompts() : {});
    if (PC.migratePreset) p = PC.migratePreset(p);
    allPresets[currentPresetId] = p;

    // Ensure DOCX style profile exists (per preset)
    if (hasDocxStyleControls) ensureDocxStyle(p);

    renderStages();
    renderPromptSelectors();
    renderUserVars();
    renderVarButtons();
    loadPromptIntoEditor(promptSelect.value);

    // Style panel
    if (hasDocxStyleControls) {
      syncDocxQuickFromStyle();
      refreshDocxStyleJson();
    }
  }

  function savePresetMeta() {
    const p = getPreset();
    p.name = (presetName.value || currentPresetId).trim();
    p.description = (presetDesc.value || "").trim();
  }

  function savePresetSettings() {
    const p = getPreset();
    p.settings = p.settings || {};
    p.settings.autoAddContext = !!autoAddContext.checked;
    p.settings.autoAddSummary = !!autoAddSummary.checked;
    p.settings.autoAddTail = !!autoAddTail.checked;
    p.settings.tailChars = parseInt(tailChars.value || "12000", 10);
    p.settings.updateSummaryAfterEachStage = !!updateSummaryAfterStage.checked;
    p.settings.updateSummaryAfterEachSection = !!updateSummaryAfterSection.checked;

    const gen = getGen(p);
    gen.chaptersCount = clampInt(chaptersCount.value, 1, 10, 2);
    gen.sectionsPerChapter = clampInt(sectionsPerChapter.value, 1, 12, 4);

    // disable chapter stages outside range
    for (const st of gen.stages) {
      if (st.type === "write_chapter" && typeof st.chapterIndex === "number" && st.chapterIndex >= gen.chaptersCount) {
        st.enabled = false;
        st._outOfRange = true;
      } else if (st && st._outOfRange) {
        delete st._outOfRange;
      }
    }
  }

  // -----------------------------
  // Stages
  // -----------------------------
  function getStages() {
    const p = getPreset();
    const gen = getGen(p);
    return gen.stages;
  }

  function stageTypeLabel(st) {
    switch (st.type) {
      case "plan": return "План";
      case "write_chapter": return "Глава (параграфы)";
      case "intro": return "Введение";
      case "conclusion": return "Заключение";
      case "bibliography": return "Источники";
      case "appendix": return "Приложение";
      case "summary": return "Summary";
      case "prompt": return "Пользовательский промпт";
      default: return st.type || "Этап";
    }
  }

  function stageSubLabel(st) {
    if (st.type === "write_chapter") {
      const oor = st._outOfRange ? " (вне диапазона — отключено)" : "";
      return `глава №${(st.chapterIndex ?? 0) + 1}, prompt: ${st.promptId || "—"}${oor}`;
    }
    if (st.type === "prompt") return `prompt: ${st.promptId || "—"}`;
    if (st.promptId) return `prompt: ${st.promptId}`;
    return "";
  }

  function renderStages() {
    const p = getPreset();
    getGen(p);
    const stages = getStages();
    stagesList.innerHTML = "";

    if (!stages.length) {
      stagesList.innerHTML = `<div class="text-xs text-white/45">Этапов пока нет. Нажмите “+ Добавить этап”.</div>`;
      return;
    }

    const promptIds = getPromptIds(p);

    stages.forEach((st, idx) => {
      const row = document.createElement("div");
      row.className = "rounded-2xl bg-white/5 border border-white/10 p-4";

      const typeLabel = stageTypeLabel(st);
      const sub = stageSubLabel(st);

      const promptSelectHtml = (selected) => promptIds.map(o =>
        `<option value="${o}" ${o === selected ? "selected" : ""}>${o}</option>`
      ).join("");

      row.innerHTML = `
        <div class="flex items-start justify-between gap-4">
          <div class="flex-1">
            <div class="flex items-center gap-3">
              <input type="checkbox" data-act="enabled" ${st.enabled ? "checked" : ""} class="accent-indigo-400 mt-0.5"/>
              <input data-act="label" value="${escapeHtml(st.label || typeLabel)}"
                class="flex-1 rounded-xl bg-white/5 border border-white/10 px-3 py-2 text-sm
                       focus:outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-500/20"/>
            </div>

            <div class="mt-2 text-xs text-white/45">
              <span class="text-white/60">${escapeHtml(typeLabel)}</span>
              <span class="ml-2">${escapeHtml(sub)}</span>
            </div>

            ${(st.type === "write_chapter") ? `
              <div class="mt-3 grid grid-cols-2 gap-3">
                <div>
                  <label class="text-xs text-white/60">Глава (номер)</label>
                  <input data-act="chapterIndex" type="number" min="1" step="1" value="${(st.chapterIndex ?? 0) + 1}"
                    class="mt-1 w-full rounded-xl bg-white/5 border border-white/10 px-3 py-2 text-sm
                           focus:outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-500/20"/>
                </div>
                <div>
                  <label class="text-xs text-white/60">Промпт</label>
                  <select data-act="promptId"
                    class="mt-1 w-full rounded-xl bg-white/5 border border-white/10 px-3 py-2 text-sm
                           focus:outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-500/20">
                    ${promptSelectHtml(st.promptId)}
                  </select>
                </div>
              </div>
            ` : ""}

            ${(st.type === "prompt") ? `
              <div class="mt-3">
                <label class="text-xs text-white/60">Промпт</label>
                <select data-act="promptId"
                  class="mt-1 w-full rounded-xl bg-white/5 border border-white/10 px-3 py-2 text-sm
                         focus:outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-500/20">
                  ${promptSelectHtml(st.promptId)}
                </select>
              </div>
            ` : ""}

            ${(st.type !== "write_chapter" && st.type !== "prompt") ? `
              <div class="mt-3">
                <label class="text-xs text-white/60">Промпт</label>
                <select data-act="promptId"
                  class="mt-1 w-full rounded-xl bg-white/5 border border-white/10 px-3 py-2 text-sm
                         focus:outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-500/20">
                  ${promptSelectHtml(st.promptId)}
                </select>
                <div class="text-[11px] text-white/40 mt-1">Можно переопределить промпт для конкретного этапа.</div>
              </div>
            ` : ""}
          </div>

          <div class="flex flex-col gap-2">
            <button data-act="up" class="rounded-xl px-3 py-2 bg-white/5 border border-white/10 hover:bg-white/10 transition text-xs">↑</button>
            <button data-act="down" class="rounded-xl px-3 py-2 bg-white/5 border border-white/10 hover:bg-white/10 transition text-xs">↓</button>
            <button data-act="delete" class="rounded-xl px-3 py-2 bg-white/5 border border-white/10 hover:bg-white/10 transition text-xs">Удалить</button>
          </div>
        </div>
      `;

      row.querySelectorAll("[data-act]").forEach(el => {
        const act = el.getAttribute("data-act");
        if (act === "up" || act === "down" || act === "delete") {
          el.addEventListener("click", (e) => {
            e.preventDefault();
            onStageAction(idx, act);
          });
        } else {
          el.addEventListener("change", () => onStageChange(idx, act, el));
          if (act === "label") el.addEventListener("input", () => onStageChange(idx, act, el));
        }
      });

      stagesList.appendChild(row);
    });
  }

  function onStageChange(idx, act, el) {
    const p = getPreset();
    const gen = getGen(p);
    const stages = getStages();
    const st = stages[idx];
    if (!st) return;

    if (act === "enabled") st.enabled = !!el.checked;
    if (act === "label") st.label = String(el.value || "").trim() || stageTypeLabel(st);
    if (act === "promptId") st.promptId = String(el.value || "").trim() || st.promptId;

    if (act === "chapterIndex") {
      const maxCh = gen.chaptersCount || 2;
      const chNum = clampInt(el.value, 1, maxCh, 1);
      st.chapterIndex = chNum - 1;
      if (!st.label || /^Глава\s+\d+$/i.test(st.label)) st.label = `Глава ${chNum}`;
    }

    persist();
    renderStages();
  }

  function onStageAction(idx, act) {
    const stages = getStages();
    if (act === "delete") {
      stages.splice(idx, 1);
      persist();
      renderStages();
      return;
    }
    if (act === "up" && idx > 0) {
      [stages[idx - 1], stages[idx]] = [stages[idx], stages[idx - 1]];
      persist();
      renderStages();
      return;
    }
    if (act === "down" && idx < stages.length - 1) {
      [stages[idx + 1], stages[idx]] = [stages[idx], stages[idx + 1]];
      persist();
      renderStages();
      return;
    }
  }

  function syncChapterStages() {
    savePresetSettings();
    const p = getPreset();
    const gen = getGen(p);
    const n = gen.chaptersCount;
    const dpi = gen.defaultPromptIds;

    const existing = gen.stages || [];
    const rest = existing.filter(s => s.type !== "write_chapter");

    const chapterStages = [];
    for (let i = 0; i < n; i++) {
      chapterStages.push({
        id: `ch_${i + 1}_${Date.now()}_${i}`,
        type: "write_chapter",
        enabled: true,
        label: `Глава ${i + 1}`,
        chapterIndex: i,
        promptId: dpi.section
      });
    }

    const planPos = rest.findIndex(s => s.type === "plan");
    if (planPos === -1) gen.stages = [...chapterStages, ...rest];
    else gen.stages = [...rest.slice(0, planPos + 1), ...chapterStages, ...rest.slice(planPos + 1)];

    persist();
    renderStages();
    alert("✅ Этапы глав синхронизированы.");
  }

  function toggleAddStage(open) {
    addStagePanel.classList.toggle("hidden", !open);
    if (open) {
      newStageLabel.value = "";
      newStageType.value = "plan";
      updateNewStageControls();
    }
  }

  function updateNewStageControls() {
    const t = newStageType.value;
    newStageChapterWrap.classList.toggle("hidden", t !== "write_chapter");
    newStagePromptWrap.classList.toggle("hidden", t !== "prompt");
    const p = getPreset();
    const gen = getGen(p);
    if (t === "write_chapter") {
      newStageChapter.value = 1;
      newStageChapter.max = String(gen.chaptersCount || 2);
    }
    if (t === "prompt") {
      fillPromptSelect(newStagePrompt, getPromptIds(p), getGen(p).defaultPromptIds.section);
    }
  }

  function addStage() {
    savePresetSettings();
    const p = getPreset();
    const gen = getGen(p);
    const dpi = gen.defaultPromptIds;

    const id = `st_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
    const label = (newStageLabel.value || "").trim() || stageTypeLabel({ type: newStageType.value });
    const t = newStageType.value;

    const stage = { id, type: t, enabled: true, label };

    if (t === "plan") stage.promptId = dpi.plan;
    if (t === "intro") stage.promptId = dpi.intro;
    if (t === "conclusion") stage.promptId = dpi.conclusion;
    if (t === "bibliography") stage.promptId = dpi.bibliography;
    if (t === "appendix") stage.promptId = dpi.appendix;
    if (t === "summary") stage.promptId = dpi.summary;

    if (t === "write_chapter") {
      const chNum = clampInt(newStageChapter.value, 1, gen.chaptersCount, 1);
      stage.chapterIndex = chNum - 1;
      stage.promptId = dpi.section;
      if (!(newStageLabel.value || "").trim()) stage.label = `Глава ${chNum}`;
    }

    if (t === "prompt") {
      stage.promptId = newStagePrompt.value || dpi.section;
    }

    gen.stages.push(stage);
    persist();
    renderStages();
    toggleAddStage(false);
  }

  // -----------------------------
  // Prompts list + management
  // -----------------------------
  function protectedPromptIds(p) {
    const gen = getGen(p);
    const dpi = gen.defaultPromptIds || {};
    return new Set(Object.values(dpi || {}).filter(Boolean));
  }

  function renderPromptSelectors() {
    const p = getPreset();
    const promptIds = getPromptIds(p);

    promptSelect.innerHTML = "";
    for (const id of promptIds) {
      const opt = document.createElement("option");
      opt.value = id;
      const title = p.prompts[id]?.title ? ` — ${p.prompts[id].title}` : "";
      opt.textContent = `${id}${title}`;
      promptSelect.appendChild(opt);
    }
    if (!promptSelect.value) promptSelect.value = promptIds[0] || "";

    fillPromptSelect(newPromptBase, promptIds, promptSelect.value || promptIds[0] || "");
    fillPromptSelect(newStagePrompt, promptIds, getGen(p).defaultPromptIds.section);
  }

  function fillPromptSelect(selectEl, promptIds, selected) {
    if (!selectEl) return;
    selectEl.innerHTML = "";
    for (const id of promptIds) {
      const opt = document.createElement("option");
      opt.value = id;
      opt.textContent = id;
      if (id === selected) opt.selected = true;
      selectEl.appendChild(opt);
    }
  }

  function loadPromptIntoEditor(promptId) {
    const p = getPreset();
    p.prompts = p.prompts || {};
    const pr = p.prompts[promptId];
    if (!pr) return;

    modelEl.value = pr.model ?? "deepseek-chat";
    modeEl.value = pr.mode ?? "text";
    temperatureEl.value = pr.temperature ?? 0.7;
    maxTokensEl.value = pr.max_tokens ?? 600;
    systemEl.value = pr.system ?? "";
    userEl.value = pr.user ?? "";
  }

  function saveEditorToPrompt() {
    const p = getPreset();
    const pid = promptSelect.value;
    p.prompts = p.prompts || {};
    p.prompts[pid] = p.prompts[pid] || { id: pid, title: pid };

    const pr = p.prompts[pid];
    pr.id = pid;
    pr.title = pr.title || pid;

    pr.model = (modelEl.value || "").trim() || "deepseek-chat";
    pr.mode = modeEl.value;
    pr.temperature = parseFloat(temperatureEl.value || "0.7");
    pr.max_tokens = parseInt(maxTokensEl.value || "600", 10);
    pr.system = systemEl.value;
    pr.user = userEl.value;
  }

  function createPrompt() {
    const p = getPreset();
    const id = (newPromptId.value || "").trim();
    const title = (newPromptTitle.value || "").trim();
    const baseId = (newPromptBase.value || "").trim();

    if (!id) { alert("Введите promptId."); return; }
    if (!isValidId(id)) { alert("promptId: только латиница, цифры и _."); return; }
    if (p.prompts[id]) { alert("Такой promptId уже существует."); return; }

    const base = p.prompts[baseId]
      ? deepClone(p.prompts[baseId])
      : { id, title: id, mode: "text", model: "deepseek-chat", temperature: 0.7, max_tokens: 600, system: "", user: "" };

    base.id = id;
    base.title = title || base.title || id;

    p.prompts[id] = base;

    newPromptId.value = "";
    newPromptTitle.value = "";

    persist();
    renderPromptSelectors();
    promptSelect.value = id;
    loadPromptIntoEditor(id);
    alert("✅ Промпт создан.");
  }

  function duplicatePrompt() {
    const p = getPreset();
    const src = promptSelect.value;
    if (!src || !p.prompts[src]) return;

    const newId = prompt(`Новый promptId (латиница, цифры, _):`, `${src}_copy_${Date.now()}`);
    if (newId === null) return;
    const id = String(newId).trim();
    if (!id) return;
    if (!isValidId(id)) { alert("Некорректный promptId."); return; }
    if (p.prompts[id]) { alert("Такой promptId уже существует."); return; }

    const copy = deepClone(p.prompts[src]);
    copy.id = id;
    copy.title = (copy.title ? copy.title + " (копия)" : id);

    p.prompts[id] = copy;

    persist();
    renderPromptSelectors();
    promptSelect.value = id;
    loadPromptIntoEditor(id);
    alert("✅ Промпт продублирован.");
  }

  function renamePromptIdSafe() {
    const p = getPreset();
    const gen = getGen(p);

    const oldId = promptSelect.value;
    if (!oldId) return;
    if (!p.prompts[oldId]) return;

    const newIdRaw = prompt(`Новый promptId для "${oldId}":`, oldId);
    if (newIdRaw === null) return;
    const newId = String(newIdRaw).trim();

    if (!newId || newId === oldId) return;
    if (!isValidId(newId)) { alert("Новый promptId некорректный (только латиница, цифры, _)."); return; }
    if (p.prompts[newId]) { alert("Такой promptId уже существует."); return; }

    const moved = p.prompts[oldId];
    delete p.prompts[oldId];
    moved.id = newId;
    p.prompts[newId] = moved;

    for (const st of gen.stages) {
      if (st && st.promptId === oldId) st.promptId = newId;
    }

    for (const k of Object.keys(gen.defaultPromptIds || {})) {
      if (gen.defaultPromptIds[k] === oldId) gen.defaultPromptIds[k] = newId;
    }

    persist();
    renderPromptSelectors();
    renderStages();

    promptSelect.value = newId;
    loadPromptIntoEditor(newId);

    alert("✅ promptId переименован. Этапы обновлены.");
  }

  function deletePromptSafe() {
    const p = getPreset();
    const gen = getGen(p);
    const id = promptSelect.value;
    if (!id) return;

    const protectedIds = protectedPromptIds(p);
    if (protectedIds.has(id)) {
      alert("Этот промпт используется как 'промпт по умолчанию' для ключевого этапа (plan/section/intro/...).\nСначала переименуйте/назначьте другой, затем удаляйте.");
      return;
    }

    if (!confirm(`Удалить промпт "${id}"?`)) return;
    delete p.prompts[id];

    const dpi = gen.defaultPromptIds || {};
    const fallbackAny = getPromptIds(p)[0] || "section";

    for (const st of gen.stages) {
      if (!st || st.promptId !== id) continue;
      st.promptId = fallbackForStage(st.type, dpi, fallbackAny);
    }

    persist();
    renderPromptSelectors();
    renderStages();

    const remaining = getPromptIds(p);
    promptSelect.value = remaining[0] || "";
    loadPromptIntoEditor(promptSelect.value);

    alert("✅ Промпт удалён.");
  }

  function fallbackForStage(type, dpi, any) {
    if (type === "plan") return dpi.plan || any;
    if (type === "write_chapter") return dpi.section || any;
    if (type === "intro") return dpi.intro || any;
    if (type === "conclusion") return dpi.conclusion || any;
    if (type === "bibliography") return dpi.bibliography || any;
    if (type === "appendix") return dpi.appendix || any;
    if (type === "summary") return dpi.summary || any;
    return any;
  }

  // -----------------------------
  // Variables
  // -----------------------------
  function renderVarButtons() {
    const p = getPreset();
    p.userVariables = p.userVariables || {};

    varButtons.innerHTML = "";
    const q = (varSearch.value || "").trim().toLowerCase();

    const builtins = builtinVarsList().map(v => ({ key: v.key, label: v.label || v.key }));
    const users = Object.keys(p.userVariables).map(k => ({ key: k, label: `user: ${k}` }));
    const allVars = [...builtins, ...users];

    allVars
      .filter(v => !q || v.key.toLowerCase().includes(q) || (v.label || "").toLowerCase().includes(q))
      .forEach(v => addVarButton(v.key, v.label));
  }

  function addVarButton(key, label) {
    const b = document.createElement("button");
    b.className = "rounded-full px-3 py-1.5 text-xs bg-white/5 border border-white/10 hover:bg-white/10 transition";
    b.type = "button";
    b.textContent = label;
    b.title = `Вставить {{${key}}}`;
    b.addEventListener("click", () => insertVar(key));
    varButtons.appendChild(b);
  }

  function insertVar(key) {
    const target = activeTextarea || userEl;
    const token = `{{${key}}}`;
    const start = target.selectionStart ?? target.value.length;
    const end = target.selectionEnd ?? target.value.length;
    target.value = target.value.slice(0, start) + token + target.value.slice(end);
    target.focus();
    target.selectionStart = target.selectionEnd = start + token.length;
  }

  function renderUserVars() {
    const p = getPreset();
    p.userVariables = p.userVariables || {};
    userVars.innerHTML = "";

    const entries = Object.entries(p.userVariables);
    if (!entries.length) {
      userVars.innerHTML = `<div class="text-xs text-white/45">Нет пользовательских переменных.</div>`;
      return;
    }

    for (const [k, v] of entries) {
      const row = document.createElement("div");
      row.className = "flex gap-2";
      row.innerHTML = `
        <input data-k="${k}" value="${escapeHtml(String(v ?? ""))}"
          class="flex-1 rounded-2xl bg-white/5 border border-white/10 px-3 py-2.5 text-sm
                 focus:outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-500/20"/>
        <button data-del="${k}"
          class="rounded-2xl px-3 py-2.5 bg-white/5 border border-white/10 hover:bg-white/10 transition text-sm">
          Удалить
        </button>
      `;
      userVars.appendChild(row);
    }

    userVars.querySelectorAll("input[data-k]").forEach(inp => {
      inp.addEventListener("input", () => {
        const k = inp.getAttribute("data-k");
        getPreset().userVariables[k] = inp.value;
        persist();
        renderVarButtons();
      });
    });
    userVars.querySelectorAll("button[data-del]").forEach(btn => {
      btn.addEventListener("click", () => {
        const k = btn.getAttribute("data-del");
        delete getPreset().userVariables[k];
        persist();
        renderUserVars();
        renderVarButtons();
      });
    });
  }

  // -----------------------------
  // AI test
  // -----------------------------
  async function testAI() {
    if (location.protocol === "file:") {
      testOut.textContent = "❌ AI недоступен при открытии как file://. Запустите через vercel dev.";
      return;
    }

    testOut.textContent = "⏳ Тестируем…";
    savePresetMeta();
    savePresetSettings();
    saveEditorToPrompt();

    // Save style too (if present)
    if (hasDocxStyleControls) applyDocxQuickToStyle();

    persist();

    const p = getPreset();
    const gen = getGen(p);
    const pr = p.prompts[promptSelect.value];

    // Provide a rich vars set so templates don't fail
    const vars = {
      topic: "Демонстрационная тема проекта",
      date: new Date().toLocaleDateString("ru-RU"),
      year: String(new Date().getFullYear()),
      accessDate: new Date().toLocaleDateString("ru-RU"),
      chaptersCount: String(gen.chaptersCount),
      sectionsPerChapter: String(gen.sectionsPerChapter),
      chapterTitle: "ГЛАВА 1. Демонстрация",
      sectionId: "1.1",
      sectionTitle: "Понятие и значение",
      chaptersTitles: "ГЛАВА 1..., ГЛАВА 2...",
      chaptersOutline: "ГЛАВА 1: 1.1..., 1.2...\nГЛАВА 2: 2.1..., 2.2...",
      plan: "ГЛАВА 1: 1.1..., 1.2...\nГЛАВА 2: 2.1..., 2.2...",
      generated_text_no_title_tail: "…(пример хвоста уже написанного текста)…",
      generated_text_no_title: "",
      generated_text_no_title_len: "0",
      generated_summary: "Краткое резюме уже написанного (пример).",
      ...(p.userVariables || {})
    };

    const sys = PC.renderTemplate(pr.system, vars);
    let user = PC.renderTemplate(pr.user, vars);

    if (p.settings?.autoAddContext) {
      const addSummary = p.settings?.autoAddSummary && !String(pr.user || "").includes("{{generated_summary}}");
      const addTail = p.settings?.autoAddTail && !String(pr.user || "").includes("{{generated_text_no_title_tail}}");
      const chunks = [];
      if (addSummary && vars.generated_summary) chunks.push("КОНТЕКСТ (SUMMARY):\n" + vars.generated_summary);
      if (addTail && vars.generated_text_no_title_tail) chunks.push("КОНТЕКСТ (TAIL):\n" + vars.generated_text_no_title_tail);
      if (chunks.length) user += "\n\n---\n" + chunks.join("\n\n");
    }

    try {
      const r = await fetch("/api/deepseek", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: pr.model,
          messages: [
            { role: "system", content: sys },
            { role: "user", content: user }
          ],
          temperature: pr.temperature,
          max_tokens: pr.max_tokens
        })
      });

      const data = await r.json();
      if (!r.ok) {
        testOut.textContent = "❌ Ошибка: " + (data?.error || r.statusText);
        return;
      }
      testOut.textContent = data?.choices?.[0]?.message?.content ?? JSON.stringify(data, null, 2);
    } catch (e) {
      testOut.textContent = "❌ Failed to fetch. Запускайте через vercel dev.\n\n" + e.message;
    }
  }

  // -----------------------------
  // Import/Export
  // -----------------------------
  function exportAll() {
    const blob = new Blob([JSON.stringify(allPresets, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "projectai-presets.json";
    a.click();
    URL.revokeObjectURL(url);
  }

  async function importAll(file) {
    const txt = await file.text();
    try {
      const parsed = JSON.parse(txt);
      if (!parsed || typeof parsed !== "object") throw new Error("invalid json");

      allPresets = parsed;
      if (Object.keys(allPresets).length === 0) allPresets = PC.defaultPresets();

      for (const k of Object.keys(allPresets)) {
        allPresets[k] = PC.migratePreset ? PC.migratePreset(allPresets[k]) : allPresets[k];
        // Ensure docxStyle exists after import
        if (hasDocxStyleControls) ensureDocxStyle(allPresets[k]);
      }

      persist();
      const activeId = PC.getActivePresetId();
      const first = allPresets[activeId] ? activeId : Object.keys(allPresets)[0];
      setPreset(first);
      alert("✅ Импорт выполнен.");
    } catch (e) {
      alert("❌ Не удалось импортировать JSON: " + e.message);
    } finally {
      importFile.value = "";
    }
  }

  // -----------------------------
  // Preset actions
  // -----------------------------
  function makeActive() {
    PC.setActivePresetId(currentPresetId);
    refreshPresetDropdown();
    updateActiveBadge();
    alert("✅ Пресет сделан активным.");
  }

  function deletePreset() {
    const keys = Object.keys(allPresets);
    if (keys.length <= 1) {
      alert("Нельзя удалить последний пресет.");
      return;
    }
    const p = getPreset();
    if (!confirm(`Удалить пресет «${p.name || currentPresetId}»?`)) return;

    const activeId = PC.getActivePresetId();
    delete allPresets[currentPresetId];

    const nextId = Object.keys(allPresets)[0];
    if (activeId === currentPresetId) PC.setActivePresetId(nextId);

    persist();
    setPreset(nextId);
  }

  function createPreset() {
    const name = (createName.value || "").trim();
    if (!name) { alert("Введите название нового пресета."); return; }
    const baseChoice = createBase.value;
    const id = `preset_${Date.now()}_${Math.floor(Math.random() * 1000)}`;

    let basePreset;
    if (baseChoice === "current") basePreset = deepClone(getPreset());
    else {
      const defaults = PC.defaultPresets();
      basePreset = deepClone(defaults[baseChoice] || defaults.school);
    }

    basePreset.id = id;
    basePreset.name = name;
    basePreset.description = basePreset.description || "";

    if (PC.migratePreset) basePreset = PC.migratePreset(basePreset);
    if (hasDocxStyleControls) basePreset.settings.docxStyle = normalizeDocxStyle(basePreset.settings.docxStyle || defaultDocxStyle());

    allPresets[id] = basePreset;
    persist();
    createName.value = "";
    setPreset(id);
    alert("✅ Пресет создан.");
  }

  // -----------------------------
  // Bindings
  // -----------------------------
  presetSelect.addEventListener("change", () => setPreset(presetSelect.value));
  presetName.addEventListener("input", () => {
    savePresetMeta();
    persist();
    contextTitle.textContent = presetName.value || currentPresetId;
  });
  presetDesc.addEventListener("input", () => {
    savePresetMeta();
    persist();
    contextSubtitle.textContent = presetDesc.value || "";
  });

  makeActiveBtn.addEventListener("click", makeActive);
  deletePresetBtn.addEventListener("click", deletePreset);
  createBtn.addEventListener("click", createPreset);

  [autoAddContext, autoAddSummary, autoAddTail, tailChars, updateSummaryAfterStage, updateSummaryAfterSection].forEach(el => {
    el.addEventListener("change", () => { savePresetSettings(); persist(); });
  });
  [chaptersCount, sectionsPerChapter].forEach(el => {
    el.addEventListener("change", () => { savePresetSettings(); persist(); renderStages(); renderVarButtons(); });
  });

  syncChaptersBtn.addEventListener("click", syncChapterStages);
  addStageBtn.addEventListener("click", () => toggleAddStage(true));
  cancelAddStageBtn.addEventListener("click", () => toggleAddStage(false));
  newStageType.addEventListener("change", updateNewStageControls);
  confirmAddStageBtn.addEventListener("click", addStage);

  tabStages.addEventListener("click", () => showTab("stages"));
  tabPrompts.addEventListener("click", () => showTab("prompts"));
  if (hasStyleUI) tabStyle.addEventListener("click", () => showTab("style"));

  promptSelect.addEventListener("change", () => loadPromptIntoEditor(promptSelect.value));
  duplicatePromptBtn.addEventListener("click", duplicatePrompt);
  renamePromptBtn.addEventListener("click", renamePromptIdSafe);
  deletePromptBtn.addEventListener("click", deletePromptSafe);

  createPromptBtn.addEventListener("click", createPrompt);

  systemEl.addEventListener("focus", () => activeTextarea = systemEl);
  userEl.addEventListener("focus", () => activeTextarea = userEl);

  saveBtn.addEventListener("click", () => {
    savePresetMeta();
    savePresetSettings();
    saveEditorToPrompt();
    if (hasDocxStyleControls) applyDocxQuickToStyle();
    persist();
    alert("✅ Сохранено.");
  });

  testBtn.addEventListener("click", testAI);

  let autosaveTimer = null;
  [modelEl, modeEl, temperatureEl, maxTokensEl, systemEl, userEl].forEach(el => {
    el.addEventListener("input", () => {
      clearTimeout(autosaveTimer);
      autosaveTimer = setTimeout(() => { saveEditorToPrompt(); persist(); }, 450);
    });
    el.addEventListener("change", () => { saveEditorToPrompt(); persist(); });
  });

  exportBtn.addEventListener("click", exportAll);
  importBtn.addEventListener("click", () => importFile.click());
  importFile.addEventListener("change", async () => {
    const f = importFile.files && importFile.files[0];
    if (!f) return;
    await importAll(f);
  });

  varSearch.addEventListener("input", renderVarButtons);
  copyHintBtn.addEventListener("click", () => {
    alert("Переменные вставляются как {{topic}}, {{generated_summary}} и т.д.\n\nЕсли включён авто-контекст — вручную вставлять summary/tail обычно не нужно.");
  });

  addVarBtn.addEventListener("click", () => {
    const k = (newVarKey.value || "").trim();
    if (!k) return;
    if (!isValidId(k)) { alert("Ключ: только латиница, цифры и _. Пример: school_name"); return; }
    const p = getPreset();
    p.userVariables = p.userVariables || {};
    if (Object.prototype.hasOwnProperty.call(p.userVariables, k)) { alert("Такая переменная уже существует."); return; }
    p.userVariables[k] = "";
    newVarKey.value = "";
    persist();
    renderUserVars();
    renderVarButtons();
  });

  // -----------------------------
  // DOCX Style bindings (optional)
  // -----------------------------
  function bindDocxStyleEventsIfAny() {
    if (!hasDocxStyleControls) return;

    // Quick fields -> style
    const quickInputs = [
      docxFontName, docxFontSizePt,
      docxMarginTop, docxMarginRight, docxMarginBottom, docxMarginLeft,
      docxLineSpacing, docxCustomLine, docxFirstLineIndent,
      docxTocEnabled,
      docxPageNumbersEnabled, docxHideOnTitle, docxPageNumbersPosition, docxPageNumbersStartAt
    ].filter(Boolean);

    quickInputs.forEach((el) => {
      const evt = (el.type === "checkbox" || el.tagName.toLowerCase() === "select") ? "change" : "input";
      el.addEventListener(evt, () => applyDocxQuickToStyle());
      if (evt !== "change") el.addEventListener("change", () => applyDocxQuickToStyle());
    });

    // Advanced JSON
    if (docxStyleFormatBtn) docxStyleFormatBtn.addEventListener("click", formatDocxJson);
    if (docxStyleApplyBtn) docxStyleApplyBtn.addEventListener("click", applyDocxJsonToStyle);
    if (docxStyleCopyBtn) docxStyleCopyBtn.addEventListener("click", copyDocxJson);

    if (docxStyleImportBtn && docxStyleImportFile) {
      docxStyleImportBtn.addEventListener("click", () => docxStyleImportFile.click());
      docxStyleImportFile.addEventListener("change", async () => {
        const f = docxStyleImportFile.files && docxStyleImportFile.files[0];
        if (!f) return;
        await importDocxJson(f);
        docxStyleImportFile.value = "";
      });
    }

    if (docxStyleResetBtn) docxStyleResetBtn.addEventListener("click", resetDocxStyle);
  }

  // -----------------------------
  // Init
  // -----------------------------
  function init() {
    allPresets = PC.loadAllPresets();
    if (!allPresets || Object.keys(allPresets).length === 0) allPresets = PC.defaultPresets();

    for (const k of Object.keys(allPresets)) {
      allPresets[k] = PC.migratePreset ? PC.migratePreset(allPresets[k]) : allPresets[k];
      if (hasDocxStyleControls) ensureDocxStyle(allPresets[k]);
    }

    currentPresetId = PC.getActivePresetId();
    if (!allPresets[currentPresetId]) currentPresetId = Object.keys(allPresets)[0];

    refreshPresetDropdown();
    setPreset(currentPresetId);

    showTab("stages");
    toggleAddStage(false);
    testOut.textContent = "—";

    bindDocxStyleEventsIfAny();
  }

  init();
})();
