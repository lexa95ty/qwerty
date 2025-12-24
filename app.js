(function () {
  const PC = window.PromptConfig;
  const $ = (id) => document.getElementById(id);

  const topicInput = $("topicInput");
  const startBtn = $("startBtn");
  const stopBtn = $("stopBtn");
  const resetBtn = $("resetBtn");
  const downloadBtn = $("downloadBtn");
  const togglePreviewBtn = $("togglePreviewBtn");

  const statusText = $("statusText");
  const badge = $("badge");
  const progressBar = $("progressBar");
  const progressLeft = $("progressLeft");
  const progressRight = $("progressRight");
  const stepsWrap = $("steps");

  const previewWrap = $("previewWrap");
  const preview = $("preview");
  const summaryBox = $("summaryBox");

  let run = { active: false, abort: false };
  let state = null;
  let uiStages = [];

  // AbortController for the *current* AI request (Stop button)
  let currentFetchController = null;

  // Safe DOM helpers (prevents crashes if an element is missing)
  function safeText(el, text) {
    if (el) el.textContent = text ?? "";
  }
  function safeHtml(el, html) {
    if (el) el.innerHTML = html ?? "";
  }
  function safeStyle(el, key, value) {
    if (el && el.style) el.style[key] = value;
  }

  function stripWrappingQuotesAndCommas(text) {
    if (text == null) return "";
    return String(text)
      .trim()
      .replace(/^["'`]+/, "")
      .replace(/["'`,]+$/, "")
      .trim();
  }

  function removeLeadingNumbering(text) {
    return String(text ?? "").replace(/^\s*\d+\.\s*/, "").trim();
  }

  function cleanBibliographyEntry(text) {
    let cleaned = stripWrappingQuotesAndCommas(text);
    cleaned = removeLeadingNumbering(cleaned);
    cleaned = stripWrappingQuotesAndCommas(cleaned);
    return cleaned.trim();
  }

  function parseBibliographyResponse(rawText) {
    if (rawText == null) return [];

    let text = String(rawText).trim();
    text = text.replace(/```(?:json)?/gi, "").replace(/```/g, "");

    const start = text.indexOf("{");
    const end = text.lastIndexOf("}");
    const candidate = start >= 0 && end > start ? text.slice(start, end + 1) : text;

    let obj = null;
    try {
      obj = JSON.parse(candidate);
    } catch (_) {
      obj = null;
    }

    if (obj && Array.isArray(obj.items)) {
      return obj.items.map(cleanBibliographyEntry).filter(Boolean);
    }

    const lines = candidate
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
      .filter((line) => {
        if (/^```/.test(line)) return false;
        if (/^\s*[{}\[\],]+\s*$/.test(line)) return false;
        if (/(^|[^a-z])(meta|total|books|websites|items)\b/i.test(line)) return false;
        if (/"[^"]+"\s*:\s*/.test(line)) return false;
        return true;
      })
      .map(cleanBibliographyEntry)
      .filter(Boolean)
      .filter((line) => /[А-Яа-яЁё]/.test(line) || /(https?:\/\/|URL:)/i.test(line));

    return lines;
  }

  function initState() {
    state = {
      topic: "",
      outlineText: "",
      chapters: [],
      introText: "",
      conclusionText: "",
      bibliography: [],
      appendix: [],
      generatedSummary: "",
      title: {
        organization:
          "Муниципальное бюджетное общеобразовательное учреждение\n«Средняя общеобразовательная школа № 1»",
        student: "Иванов Иван Иванович",
        class: "9Г",
        teacher: "Петрова Мария Сергеевна",
        teacherSubject: "биологии",
        city: "Москва",
        year: String(new Date().getFullYear()),
      },
    };
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    }[c]));
  }

  function setStatus(kind, text) {
    safeText(statusText, text || "");
    if (!badge) return;
    if (kind === "idle") {
      badge.textContent = "Готово";
      badge.className =
        "text-xs px-2.5 py-1 rounded-full bg-white/5 border border-white/10 text-white/70";
    } else if (kind === "run") {
      badge.textContent = "Генерация";
      badge.className =
        "text-xs px-2.5 py-1 rounded-full bg-indigo-500/20 border border-indigo-300/20 text-indigo-100";
    } else if (kind === "done") {
      badge.textContent = "Готово";
      badge.className =
        "text-xs px-2.5 py-1 rounded-full bg-emerald-500/20 border border-emerald-300/20 text-emerald-100";
    } else if (kind === "stop") {
      badge.textContent = "Остановлено";
      badge.className =
        "text-xs px-2.5 py-1 rounded-full bg-amber-500/20 border border-amber-300/20 text-amber-100";
    } else if (kind === "error") {
      badge.textContent = "Ошибка";
      badge.className =
        "text-xs px-2.5 py-1 rounded-full bg-red-500/20 border border-red-300/20 text-red-100";
    }
  }

  function setButtons(running) {
    startBtn.disabled = running;
    stopBtn.disabled = !running;
    topicInput.disabled = running;
    startBtn.classList.toggle("opacity-70", running);
    startBtn.classList.toggle("cursor-not-allowed", running);
  }

  function renderSteps(activeIndex = -1, doneIndex = -1) {
    stepsWrap.innerHTML = "";
    uiStages.forEach((s, idx) => {
      const done = idx <= doneIndex;
      const active = idx === activeIndex;

      const row = document.createElement("div");
      row.className =
        "flex items-center justify-between gap-3 rounded-2xl px-4 py-3 border " +
        (active
          ? "border-white/20 bg-white/8 soft-ring"
          : done
            ? "border-white/10 bg-white/6"
            : "border-white/10 bg-transparent");

      row.innerHTML = `
        <div class="flex items-center gap-3">
          <div class="h-2.5 w-2.5 rounded-full ${done ? "bg-emerald-400" : active ? "bg-white pulse-dot" : "bg-white/30"}"></div>
          <div class="text-sm ${active ? "text-white" : "text-white/85"}">${escapeHtml(s.label || s.type)}</div>
        </div>
        <div class="text-xs ${done ? "text-emerald-200" : active ? "text-white/70" : "text-white/35"}">
          ${done ? "готово" : active ? "в процессе" : "ожидание"}
        </div>
      `;
      stepsWrap.appendChild(row);
    });
  }

  function setProgress(doneIndex) {
    const total = uiStages.length || 0;
    const doneCount = Math.max(0, Math.min(total, doneIndex + 1));
    const pct = total ? Math.round((doneCount / total) * 100) : 0;
    safeStyle(progressBar, "width", pct + "%");
    safeText(progressLeft, pct + "%");
    safeText(progressRight, `${doneCount}/${total}`);
  }

  function previewAdd(title, text) {
    if (!preview) return;
    if (preview.textContent.includes("Превью появится")) preview.innerHTML = "";
    const item = document.createElement("div");
    item.className = "rounded-xl bg-white/5 border border-white/10 p-3";
    item.innerHTML = `
      <div class="text-xs text-white/45">${escapeHtml(title)}</div>
      <div class="mt-1 text-sm text-white/80 leading-relaxed whitespace-pre-wrap">${escapeHtml(text)}</div>
    `;
    preview.appendChild(item);
    preview.scrollTop = preview.scrollHeight;
  }

  function buildChaptersOutline() {
    const out = [];
    for (const ch of state.chapters) {
      out.push(ch.title || "");
      for (const sec of (ch.sections || [])) out.push(`${sec.id}. ${sec.title}`);
      out.push("");
    }
    return out.join("\n").trim();
  }

  function buildChaptersTitles() {
    return state.chapters.map((c) => c.title).filter(Boolean).join(", ");
  }

  function buildGeneratedTextNoTitle() {
    const out = [];

    if (state.chapters && state.chapters.length) {
      out.push("СОДЕРЖАНИЕ");
      out.push(buildChaptersOutline());
    } else if (state.outlineText) {
      out.push("СОДЕРЖАНИЕ");
      out.push(state.outlineText.trim());
    }

    if (state.introText && state.introText.trim()) {
      out.push("ВВЕДЕНИЕ");
      out.push(state.introText.trim());
    }

    for (const ch of state.chapters || []) {
      if (ch.title) out.push(ch.title);
      for (const sec of ch.sections || []) {
        if (!sec.content || !String(sec.content).trim()) continue;
        out.push(`${sec.id}. ${sec.title}`);
        out.push(String(sec.content).trim());
      }
    }

    if (state.conclusionText && state.conclusionText.trim()) {
      out.push("ЗАКЛЮЧЕНИЕ");
      out.push(state.conclusionText.trim());
    }

    if (Array.isArray(state.bibliography) && state.bibliography.length) {
      out.push("СПИСОК ЛИТЕРАТУРЫ");
      state.bibliography.forEach((item, i) => out.push(`${i + 1}. ${item}`));
    }

    return out.join("\n\n").trim();
  }

  // Practical part text (used for appendix generation)
  function buildPracticalText() {
    const chapters = Array.isArray(state.chapters) ? state.chapters : [];
    if (!chapters.length) return "";

    const out = [];
    for (let i = 0; i < chapters.length; i++) {
      const ch = chapters[i] || {};
      const title = String(ch.title || "");
      const isPractical = i > 0 || /практическ/i.test(title);
      if (!isPractical) continue;

      if (title.trim()) out.push(title.trim());
      for (const sec of (ch.sections || [])) {
        const sid = String(sec?.id || "").trim();
        const st = String(sec?.title || "").trim();
        const ct = String(sec?.content || "").trim();
        if (sid || st) out.push(`${sid}${sid && st ? ". " : ""}${st}`.trim());
        if (ct) out.push(ct);
      }
      out.push("");
    }

    return out.join("\n\n").trim();
  }

  function tailText(str, maxChars) {
    if (!str) return "";
    const s = String(str);
    if (s.length <= maxChars) return s;
    return "…(обрезано)\n" + s.slice(-maxChars);
  }

  function clampInt(v, min, max, def) {
    const n = parseInt(String(v || ""), 10);
    if (!Number.isFinite(n)) return def;
    return Math.max(min, Math.min(max, n));
  }

  async function callAI(promptId, vars) {
    const preset = PC.getActivePreset();
    const pr = preset.prompts?.[promptId];
    if (!pr) throw new Error(`Prompt not found: ${promptId}`);

    const sys = PC.renderTemplate(pr.system, vars);
    let user = PC.renderTemplate(pr.user, vars);

    if (preset.settings?.autoAddContext) {
      const chunks = [];
      const wantsSummary = preset.settings?.autoAddSummary && !pr.user.includes("{{generated_summary}}");
      const wantsTail = preset.settings?.autoAddTail && !pr.user.includes("{{generated_text_no_title_tail}}");
      if (wantsSummary && vars.generated_summary) chunks.push("КОНТЕКСТ (SUMMARY):\n" + vars.generated_summary);
      if (wantsTail && vars.generated_text_no_title_tail) chunks.push("КОНТЕКСТ (TAIL):\n" + vars.generated_text_no_title_tail);
      if (chunks.length) user += "\n\n---\n" + chunks.join("\n\n");
    }

    const body = {
      model: pr.model,
      messages: [
        { role: "system", content: sys },
        { role: "user", content: user },
      ],
      temperature: pr.temperature,
      max_tokens: pr.max_tokens,
    };

    if (run.abort) {
      throw new Error("Aborted");
    }

    // Cancel previous in-flight request (defensive)
    if (currentFetchController) {
      try { currentFetchController.abort(); } catch (_) {}
    }

    currentFetchController = new AbortController();
    try {
      const r = await fetch("/api/deepseek", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal: currentFetchController.signal,
      });

      const data = await r.json().catch(() => ({}));
      if (!r.ok) {
        const msg = data?.error || data?.message || r.statusText || "API error";
        throw new Error(msg);
      }
      const text = data?.choices?.[0]?.message?.content ?? "";
      return { text, raw: data };
    } catch (e) {
      // Fetch abort should not be treated as an error in UI
      if (e && (e.name === "AbortError" || String(e.message || "").toLowerCase().includes("aborted"))) {
        throw new Error("Aborted");
      }
      throw e;
    } finally {
      currentFetchController = null;
    }
  }

  function buildVars(extra = {}) {
    const preset = PC.getActivePreset();
    const gen = preset.settings?.generation || {};
    const cc = clampInt(gen.chaptersCount, 1, 10, 2);
    const spc = clampInt(gen.sectionsPerChapter, 1, 12, 4);

    const now = new Date();
    const fullNoTitle = buildGeneratedTextNoTitle();
    const practicalText = buildPracticalText();
    const tailChars = preset.settings?.tailChars ?? 12000;
    const chaptersOutline = buildChaptersOutline();
    const plan = state.outlineText || chaptersOutline || "";

    const appendixJson = Array.isArray(state.appendix) && state.appendix.length
      ? tailText(JSON.stringify({ items: state.appendix }, null, 2), tailChars)
      : "";

    return {
      topic: state.topic,
      year: state.title.year || String(now.getFullYear()),
      accessDate: now.toLocaleDateString("ru-RU"),

      chaptersCount: String(cc),
      sectionsPerChapter: String(spc),

      chaptersTitles: buildChaptersTitles(),
      chaptersOutline,
      plan,

      chapterTitle: "",
      sectionId: "",
      sectionTitle: "",

      generated_text_no_title: fullNoTitle,
      generated_text_no_title_tail: tailText(fullNoTitle, tailChars),
      generated_text_no_title_len: String(fullNoTitle.length),

      practical_text: practicalText,
      practical_text_tail: tailText(practicalText, tailChars),
      practical_text_len: String(practicalText.length),

      appendix_json: appendixJson,

      generated_summary: state.generatedSummary || "",

      ...(preset.userVariables || {}),
      ...extra,
    };
  }

  function normalizePlan(planJson, chaptersCount, sectionsPerChapter) {
    const chapters = Array.isArray(planJson?.chapters) ? planJson.chapters : [];
    const out = [];

    for (let i = 0; i < chaptersCount; i++) {
      const chIn = chapters[i] || {};
      const title =
        chIn.title ||
        `ГЛАВА ${i + 1}. ${i === 0 ? "ТЕОРЕТИЧЕСКАЯ ЧАСТЬ" : "ПРАКТИЧЕСКАЯ ЧАСТЬ"}`;
      const secIn = Array.isArray(chIn.sections) ? chIn.sections : [];
      const sections = [];
      for (let j = 0; j < sectionsPerChapter; j++) {
        const sIn = secIn[j] || {};
        const id = `${i + 1}.${j + 1}`;
        const sTitle = sIn.title || `Раздел ${id}`;
        sections.push({ id, title: sTitle, content: "" });
      }
      out.push({ title, sections });
    }
    return out;
  }

  function normalizeAppendix(input) {
    let items = [];
    if (Array.isArray(input)) items = input;
    else if (input && typeof input === "object") {
      if (Array.isArray(input.items)) items = input.items;
      else if (Array.isArray(input.appendix)) items = input.appendix;
    }

    const out = [];
    for (const it of items) {
      if (!it || typeof it !== "object") continue;
      const type = String(it.type || "").trim().toLowerCase();

      if (type === "table") {
        const title = String(it.title || it.caption || "Таблица").trim();
        const headers = Array.isArray(it.headers) ? it.headers.map((x) => String(x)) : [];
        const rows = Array.isArray(it.rows) ? it.rows.map((r) => (Array.isArray(r) ? r.map((x) => String(x)) : [])) : [];
        const notes = String(it.notes || "").trim();
        out.push({ type: "table", title, headers, rows, notes });
        continue;
      }

      if (type === "chart") {
        const chartType = String(it.chartType || "bar").trim().toLowerCase();
        const title = String(it.title || it.caption || "Рисунок").trim();
        const labels = Array.isArray(it.labels) ? it.labels.map((x) => String(x)) : [];
        const seriesIn = Array.isArray(it.series) ? it.series : [];
        const series = seriesIn
          .map((s) => ({
            name: String(s?.name || "").trim() || "Series",
            values: Array.isArray(s?.values) ? s.values.map((v) => Number(v)).filter((n) => Number.isFinite(n)) : [],
          }))
          .filter((s) => s.values.length);
        const notes = String(it.notes || "").trim();
        out.push({ type: "chart", chartType, title, labels, series, notes });
        continue;
      }

      if (type === "text") {
        const title = String(it.title || "").trim();
        const content = String(it.content || it.text || "").trim();
        if (title || content) out.push({ type: "text", title, content });
        continue;
      }
    }

    return out;
  }

  function ensureStructure() {
    const preset = PC.getActivePreset();
    const gen = preset.settings?.generation || {};
    const cc = clampInt(gen.chaptersCount, 1, 10, 2);
    const spc = clampInt(gen.sectionsPerChapter, 1, 12, 4);

    if (state.chapters && state.chapters.length) return;

    state.chapters = [];
    for (let i = 0; i < cc; i++) {
      const title = `ГЛАВА ${i + 1}. ${i === 0 ? "ТЕОРЕТИЧЕСКАЯ ЧАСТЬ" : "ПРАКТИЧЕСКАЯ ЧАСТЬ"}`;
      const sections = [];
      for (let j = 0; j < spc; j++) {
        const id = `${i + 1}.${j + 1}`;
        sections.push({ id, title: `Раздел ${id}`, content: "" });
      }
      state.chapters.push({ title, sections });
    }
    state.outlineText = buildChaptersOutline();
  }

  async function runSummaryStage(promptId) {
    const pid = promptId || "summary";
    const vars = buildVars();
    const resp = await callAI(pid, vars);
    state.generatedSummary = (resp.text || "").trim();
    safeText(summaryBox, state.generatedSummary || "—");
    previewAdd("Summary", state.generatedSummary || "—");
  }

  async function updateSummaryAfterStageIfEnabled(stageType, summaryPromptId) {
    const preset = PC.getActivePreset();
    if (!preset.settings?.updateSummaryAfterEachStage) return;
    if (stageType === "summary") return;
    await runSummaryStage(summaryPromptId);
  }

  async function updateSummaryAfterSectionIfEnabled(summaryPromptId) {
    const preset = PC.getActivePreset();
    if (!preset.settings?.updateSummaryAfterEachSection) return;
    await runSummaryStage(summaryPromptId);
  }

  function getEnabledStages() {
    const preset = PC.getActivePreset();
    const gen = preset.settings?.generation || {};
    const stages = Array.isArray(gen.stages) ? gen.stages : [];
    const enabled = stages.filter((s) => s && s.enabled !== false);
    if (enabled.length) return enabled;

    const cc = clampInt(gen.chaptersCount, 1, 10, 2);
    const dpi = gen.defaultPromptIds || {
      plan: "plan",
      section: "section",
      intro: "intro",
      conclusion: "conclusion",
      bibliography: "bibliography",
      summary: "summary",
      appendix: "appendix",
    };
    return PC.defaultPipeline(cc, dpi);
  }

  function defaultPromptForStage(stageType, dpi) {
    if (stageType === "plan") return dpi.plan || "plan";
    if (stageType === "write_chapter") return dpi.section || "section";
    if (stageType === "intro") return dpi.intro || "intro";
    if (stageType === "conclusion") return dpi.conclusion || "conclusion";
    if (stageType === "bibliography") return dpi.bibliography || "bibliography";
    if (stageType === "appendix") return dpi.appendix || "appendix";
    if (stageType === "summary") return dpi.summary || "summary";
    if (stageType === "prompt") return dpi.section || "section";
    return "section";
  }

  async function runStage(stage, stageIndex) {
    if (run.abort) return;

    const label = stage.label || stage.type;
    setStatus("run", `Этап: ${label}`);
    renderSteps(stageIndex, stageIndex - 1);
    setProgress(stageIndex - 1);

    const preset = PC.getActivePreset();
    const gen = preset.settings?.generation || {};
    const cc = clampInt(gen.chaptersCount, 1, 10, 2);
    const spc = clampInt(gen.sectionsPerChapter, 1, 12, 4);
    const dpi = gen.defaultPromptIds || {
      plan: "plan",
      section: "section",
      intro: "intro",
      conclusion: "conclusion",
      bibliography: "bibliography",
      summary: "summary",
      appendix: "appendix",
    };

    const stagePromptId = stage.promptId || defaultPromptForStage(stage.type, dpi);

    if (stage.type === "plan") {
      const vars = buildVars({ chaptersCount: String(cc), sectionsPerChapter: String(spc) });
      const resp = await callAI(stagePromptId, vars);

      let planJson = null;
      try {
        planJson = JSON.parse(resp.text);
      } catch (_) {
        planJson = PC.extractJsonObject(resp.text);
      }
      state.chapters = normalizePlan(planJson || {}, cc, spc);
      state.outlineText = buildChaptersOutline();
      previewAdd("План", state.outlineText || "—");
      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    if (stage.type === "write_chapter") {
      ensureStructure();
      const chIndex = clampInt((stage.chapterIndex ?? 0) + 1, 1, cc, 1) - 1;
      const chapter = state.chapters[chIndex];
      if (!chapter) return;

      setStatus("run", `Этап: ${label} — параграфы`);
      for (const sec of chapter.sections) {
        if (run.abort) return;

        setStatus("run", `${label}: ${sec.id}. ${sec.title}`);
        const vars = buildVars({
          chapterTitle: chapter.title,
          sectionId: sec.id,
          sectionTitle: sec.title,
        });
        const resp = await callAI(stagePromptId, vars);
        sec.content = (resp.text || "").trim();

        previewAdd(`${sec.id}. ${sec.title}`, sec.content.slice(0, 900) + (sec.content.length > 900 ? "…" : ""));
        await updateSummaryAfterSectionIfEnabled(dpi.summary);
      }

      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    if (stage.type === "intro") {
      const vars = buildVars();
      const resp = await callAI(stagePromptId, vars);
      state.introText = (resp.text || "").trim();
      previewAdd("Введение", state.introText.slice(0, 1200) + (state.introText.length > 1200 ? "…" : ""));
      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    if (stage.type === "conclusion") {
      const vars = buildVars();
      const resp = await callAI(stagePromptId, vars);
      state.conclusionText = (resp.text || "").trim();
      previewAdd("Заключение", state.conclusionText.slice(0, 1200) + (state.conclusionText.length > 1200 ? "…" : ""));
      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    if (stage.type === "bibliography") {
      const vars = buildVars();
      const resp = await callAI(stagePromptId, vars);
      const parsedSources = parseBibliographyResponse(resp.text);
      state.bibliography = parsedSources.slice(0, 20);

      previewAdd("Источники", state.bibliography.map((x, i) => `${i + 1}. ${x}`).join("\n"));
      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    if (stage.type === "appendix") {
      const vars = buildVars();
      const resp = await callAI(stagePromptId, vars);

      let j = null;
      try {
        j = JSON.parse(resp.text);
      } catch (_) {
        j = PC.extractJsonObject(resp.text);
      }

      const normalized = normalizeAppendix(j || {});
      state.appendix = normalized;

      if (!state.appendix.length && String(resp.text || "").trim()) {
        state.appendix = [{ type: "text", title: "Приложение", content: String(resp.text).trim() }];
      }

      const previewText = state.appendix.length
        ? JSON.stringify({ items: state.appendix }, null, 2)
        : "—";
      previewAdd("Приложение", previewText.slice(0, 1500) + (previewText.length > 1500 ? "…" : ""));

      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    if (stage.type === "summary") {
      await runSummaryStage(stagePromptId);
      return;
    }

    if (stage.type === "prompt") {
      const vars = buildVars();
      const resp = await callAI(stagePromptId, vars);
      previewAdd(label, (resp.text || "").trim().slice(0, 1600));
      await updateSummaryAfterStageIfEnabled(stage.type, dpi.summary);
      return;
    }

    previewAdd("Этап", `Неизвестный тип: ${stage.type}`);
  }

  async function generateAll() {
    run.active = true;
    run.abort = false;
    setButtons(true);
    downloadBtn.disabled = true;

    uiStages = getEnabledStages();
    renderSteps(0, -1);
    setProgress(-1);
    setStatus("run", `Тема: «${state.topic}»`);

    for (let i = 0; i < uiStages.length; i++) {
      if (run.abort) break;
      renderSteps(i, i - 1);
      setProgress(i - 1);
      await runStage(uiStages[i], i);
      renderSteps(i, i);
      setProgress(i);
    }

    if (run.abort) {
      setStatus("stop", "Остановлено пользователем");
      setButtons(false);
      run.active = false;
      return;
    }

    renderSteps(-1, uiStages.length - 1);
    setProgress(uiStages.length - 1);
    setStatus("done", "Готово — можно скачать DOCX");
    setButtons(false);
    run.active = false;
    downloadBtn.disabled = false;
  }

  // =============================
  // DOCX EXPORT (profile-aware)
  // =============================
  async function downloadDocx() {
    try {
      const activePreset = (PC && typeof PC.getActivePreset === "function") ? PC.getActivePreset() : null;
      const activeDocxStyle = activePreset && activePreset.settings ? activePreset.settings.docxStyle : null;

      let generatorDocxStyle = null;
      try {
        const gen = localStorage.getItem("projectai_docx_style_generator_v1");
        if (gen) {
          const parsed = JSON.parse(gen);
          if (parsed && parsed.meta && parsed.meta.schemaVersion === "projectai-docx-style/v1") {
            generatorDocxStyle = parsed;
          }
        }
      } catch (_) {}

      const resolvedDocxStyle = activePreset ? (activeDocxStyle ?? null) : (generatorDocxStyle ?? null);
      // Prefer server-side DOCX generator to avoid CDN issues.
      const payload = {
        topic: state.topic,
        title: state.title,
        outlineText: state.outlineText,
        chapters: state.chapters,
        introText: state.introText,
        conclusionText: state.conclusionText,
        bibliography: state.bibliography,
        appendix: state.appendix,
        docxStyle: resolvedDocxStyle,
      };

      // 1) Try server-side export (works even if CDNs are blocked).
      let serverExportError = null;
      try {
        const resp = await fetch("/api/docx", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });

        if (resp.ok) {
          const blob = await resp.blob();
          const safeTopic = String(state.topic || "project")
            .trim()
            .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
            .replace(/\s+/g, "_")
            .slice(0, 60) || "project";

          const filename = `Проект_${safeTopic}.docx`;
          const url = URL.createObjectURL(blob);

          const a = document.createElement("a");
          a.href = url;
          a.download = filename;
          document.body.appendChild(a);
          a.click();
          a.remove();

          // Revoke a bit later to avoid rare download race conditions.
          setTimeout(() => URL.revokeObjectURL(url), 1500);
          return;
        }

        // Make server-side failure visible (otherwise users only see the fallback error).
        const txt = await resp.text().catch(() => "");
        serverExportError = `Серверный экспорт /api/docx вернул ${resp.status}.` + (txt ? `\n${txt.slice(0, 500)}` : "");
        console.warn(serverExportError);
      } catch (e) {
        serverExportError = `Серверный экспорт /api/docx недоступен: ${e?.message || e}`;
        console.warn(serverExportError);
      }

      alert(serverExportError || "Серверный экспорт недоступен.");
      return;
    } catch (e) {
      alert("Ошибка DOCX: " + (e?.message || e));
    }
  }

  function resetUI() {
    if (run.active) {
      run.abort = true;
      run.active = false;
    }

    if (currentFetchController) {
      try { currentFetchController.abort(); } catch (_) {}
      currentFetchController = null;
    }
    initState();

    setButtons(false);
    uiStages = [];
    setProgress(-1);
    renderSteps(-1, -1);
    setStatus("idle", "Ожидание темы");
    downloadBtn.disabled = true;

    safeHtml(preview, `<div class="text-white/60">Превью появится после начала генерации.</div>`);
    safeText(summaryBox, "—");
  }

  async function onStart() {
    if (location.protocol === "file:") {
      alert("AI недоступен при открытии как file://. Запустите через Vercel (vercel dev) или задеплойте на Vercel.");
      return;
    }
    const topic = (topicInput.value || "").trim();
    if (!topic) {
      alert("Введите тему проекта.");
      topicInput.focus();
      return;
    }

    initState();
    state.topic = topic;

    try {
      await generateAll();
    } catch (e) {
      console.error(e);
      const msg = String(e?.message || e || "");
      if (msg === "Aborted" || msg.toLowerCase().includes("aborted")) {
        setButtons(false);
        run.active = false;
        setStatus("stop", "Остановлено пользователем");
        return;
      }

      setStatus("error", "Ошибка: " + msg);
      setButtons(false);
      run.active = false;
      alert("Ошибка AI: " + msg);
    }
  }

  function onStop() {
    if (!run.active) return;
    run.abort = true;
    run.active = false;
    if (currentFetchController) {
      try { currentFetchController.abort(); } catch (_) {}
      currentFetchController = null;
    }
    setButtons(false);
    setStatus("stop", "Остановлено пользователем");
  }

  startBtn.addEventListener("click", onStart);
  stopBtn.addEventListener("click", onStop);
  resetBtn.addEventListener("click", resetUI);
  downloadBtn.addEventListener("click", downloadDocx);

  togglePreviewBtn.addEventListener("click", () => {
    previewWrap.classList.toggle("hidden");
    togglePreviewBtn.textContent = previewWrap.classList.contains("hidden") ? "Превью" : "Скрыть превью";
  });

  resetUI();
})();
