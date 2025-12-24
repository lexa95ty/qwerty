// prompt-config.js
(function () {
  const STORAGE_KEY = "projectai_presets_v4";
  const ACTIVE_KEY = "projectai_active_preset_v4";

  // DOCX style schema (per preset)
  const DOCX_STYLE_SCHEMA_VERSION = "projectai-docx-style/v1";

  // Built-in vars for templates (available in prompts)
  const BUILTIN_VARS = {
    topic: "",
    date: "",
    accessDate: "",
    chapterTitle: "",
    sectionId: "",
    sectionTitle: "",
    chaptersCount: "",
    sectionsPerChapter: "",
    generated_summary: "",
    generated_text_no_title: "",
    generated_text_no_title_tail: "",

    // Practical part helpers (for appendix / diagrams / tables)
    practical_text: "",
    practical_text_tail: "",
    practical_text_len: "",

    // Appendix (optional; can be generated as JSON)
    appendix_json: ""
  };

  function deepClone(obj) { return JSON.parse(JSON.stringify(obj)); }

  // ----------------------
  // DOCX style defaults
  // ----------------------
  function defaultDocxStyle() {
    // If user saved a profile in style-generator.html ("Сохранить локально"),
    // reuse it as default for all presets (best-effort).
    try {
      const gen = localStorage.getItem("projectai_docx_style_generator_v1");
      if (gen) {
        const parsed = JSON.parse(gen);
        if (parsed && parsed.meta && parsed.meta.schemaVersion === DOCX_STYLE_SCHEMA_VERSION) {
          return normalizeDocxStyle(parsed);
        }
      }
    } catch (_) {}

    const now = new Date().toISOString();
    return {
      meta: {
        schemaVersion: DOCX_STYLE_SCHEMA_VERSION,
        name: "Школьный (ГОСТ-подобный)",
        author: "ProjectAI",
        description: "Times New Roman 14, поля 3/1.5/2/2, межстрочный 1.5, отступ 1.25 см",
        notes: "",
        createdAt: now,
        updatedAt: now
      },
      page: {
        size: "A4",
        orientation: "portrait",
        marginsCm: { top: 2.0, right: 1.5, bottom: 2.0, left: 3.0 },
        gutterCm: 0.0,
        mirrorMargins: false,
        headerDistanceCm: 1.25,
        footerDistanceCm: 1.25,
        columns: { count: 1, spacingCm: 1.0 }
      },
      base: {
        font: { name: "Times New Roman", sizePt: 14, color: "#000000", lang: "ru-RU" },
        paragraph: {
          alignment: "justify",
          lineSpacing: "1.5",
          customLine: 1.35,
          spacingBeforePt: 0,
          spacingAfterPt: 0,
          firstLineIndentCm: 1.25,
          hangingIndentCm: 0,
          leftIndentCm: 0,
          rightIndentCm: 0,
          widowControl: true,
          keepTogether: false,
          keepWithNext: false
        }
      },
      headings: {
        h1: {
          font: { sizePt: 14, bold: true, italic: false, underline: false, allCaps: true, color: "#000000" },
          paragraph: { alignment: "center", spacingBeforePt: 24, spacingAfterPt: 24, pageBreakBefore: true, keepWithNext: true }
        },
        h2: {
          font: { sizePt: 14, bold: true, italic: false, underline: false, allCaps: false, color: "#000000" },
          paragraph: { alignment: "center", spacingBeforePt: 18, spacingAfterPt: 12, pageBreakBefore: false, keepWithNext: true }
        },
        h3: {
          font: { sizePt: 14, bold: true, italic: false, underline: false, allCaps: false, color: "#000000" },
          paragraph: { alignment: "left", spacingBeforePt: 12, spacingAfterPt: 6, pageBreakBefore: false, keepWithNext: true }
        },
        h4: {
          font: { sizePt: 14, bold: true, italic: false, underline: false, allCaps: false, color: "#000000" },
          paragraph: { alignment: "left", spacingBeforePt: 10, spacingAfterPt: 4, pageBreakBefore: false, keepWithNext: true }
        }
      },
      toc: { showPageNumbers: true, tabLeader: "dots", itemIndentCm: 0, subItemIndentCm: 1.25 },
      numbering: {
        bullets: { symbol: "•" },
        decimal: { format: "1." },
        common: { leftIndentCm: 1.25, hangingIndentCm: 0.6 }
      },
      table: {
        font: { name: "", sizePt: 12 },
        layout: { widthPercent: 100, alignment: "center" },
        borders: { sizePt: 0.75, style: "single", color: "#000000" },
        cellMarginsMm: { top: 1.5, right: 2.0, bottom: 1.5, left: 2.0 },
        headerRow: { shading: "#eaeaea", bold: true, repeatHeader: true },
        behavior: { allowRowSplit: true, keepTogether: false, keepWithNext: false }
      },
      captions: {
        style: { font: { sizePt: 12 }, paragraph: { alignment: "center" } },
        table: { position: "above", numberFormat: "Таблица {n} – " },
        figure: { position: "below", numberFormat: "Рисунок {n} – " }
      },
      headersFooters: {
        differentFirstPage: true,
        differentOddEven: false,
        headerText: "",
        footerText: "",
        pageNumbers: { enabled: true, position: "footer-center", startAt: 1, hideOnTitle: true }
      },
      bibliography: {
        numbering: { format: "1." },
        item: { alignment: "justify", hangingIndentCm: 0, spacingBeforePt: 0, spacingAfterPt: 0 }
      },
      advanced: { compatibilityJson: "", raw: "" }
    };
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function normalizeDocxStyle(input) {
    const base = defaultDocxStyle();
    const s = deepClone(input || {});
    // Shallow-deep merge (best-effort)
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
    const out = merge(base, s);

    out.meta = out.meta || {};
    out.meta.schemaVersion = DOCX_STYLE_SCHEMA_VERSION;
    out.meta.createdAt = out.meta.createdAt || nowIso();
    out.meta.updatedAt = nowIso();

    // Ensure required branches exist
    out.page = out.page || base.page;
    out.base = out.base || base.base;
    out.headings = out.headings || base.headings;
    out.toc = out.toc || base.toc;
    out.headersFooters = out.headersFooters || base.headersFooters;
    out.headersFooters.pageNumbers = out.headersFooters.pageNumbers || base.headersFooters.pageNumbers;
    out.bibliography = out.bibliography || base.bibliography;

    return out;
  }

  function defaultPrompts() {
    const commonSystem =
      "Ты помощник по написанию учебных проектов. Пиши академическим, нейтральным стилем. " +
      "Избегай воды и повторов. Не выдумывай факты. Если нужны числа — помечай как пример.";

    const plan = {
      id: "plan",
      title: "План (JSON)",
      mode: "json",
      model: "deepseek-chat",
      temperature: 0.5,
      max_tokens: 1100,
      system: commonSystem,
      user:
        "Составь план учебного проекта по теме: «{{topic}}».\n\n" +
        "Требования к структуре:\n" +
        "- Количество глав: {{chaptersCount}}\n" +
        "- Параграфов в каждой главе: {{sectionsPerChapter}}\n" +
        "- Нумерация: i.j (например 1.1, 1.2, ...)\n\n" +
        "Верни СТРОГО JSON:\n" +
        "{\n" +
        '  "chapters":[\n' +
        '    {"title":"ГЛАВА 1. ...","sections":[{"id":"1.1","title":"..."}, ...]}\n' +
        "  ]\n" +
        "}\n"
    };

    const section = {
      id: "section",
      title: "Параграф (текст)",
      mode: "text",
      model: "deepseek-chat",
      temperature: 0.7,
      max_tokens: 800,
      system: commonSystem,
      user:
        "Тема проекта: «{{topic}}».\n" +
        "Текущая глава: {{chapterTitle}}\n" +
        "Напиши раздел: {{sectionId}}. {{sectionTitle}}.\n\n" +
        "Требования:\n" +
        "- 2–4 абзаца\n" +
        "- связный академический текст\n" +
        "- без вымышленных фактов\n"
    };

    const intro = {
      id: "intro",
      title: "Введение (текст)",
      mode: "text",
      model: "deepseek-chat",
      temperature: 0.6,
      max_tokens: 900,
      system: commonSystem,
      user:
        "Напиши введение к проекту по теме: «{{topic}}».\n\n" +
        "Структура:\n" +
        "1) Вводный абзац\n" +
        "2) Актуальность\n" +
        "3) Цель (с глагола)\n" +
        "4) Задачи (4 пункта)\n" +
        "5) Гипотеза\n" +
        "6) Методы\n" +
        "7) Объект и предмет\n"
    };

    const conclusion = {
      id: "conclusion",
      title: "Заключение (текст)",
      mode: "text",
      model: "deepseek-chat",
      temperature: 0.6,
      max_tokens: 900,
      system: commonSystem,
      user:
        "Напиши заключение к проекту по теме: «{{topic}}».\n" +
        "Сделай 4–7 предложений: итоги, результаты, краткий вывод."
    };

    const bibliography = {
      id: "bibliography",
      title: "Список источников (JSON)",
      mode: "json",
      model: "deepseek-chat",
      temperature: 0.3,
      max_tokens: 700,
      system: commonSystem,
      user:
        "Составь список литературы (8–10 позиций) для проекта по теме: «{{topic}}».\n" +
        "Без выдуманных страниц/ISBN. Можно смешивать книги/статьи/сайты.\n" +
        "Для сайтов указывай URL и дату обращения: {{accessDate}}.\n\n" +
        "Верни СТРОГО JSON:\n" +
        "{\n" +
        '  "items":[ "Источник 1", "Источник 2", "..." ]\n' +
        "}\n"
    };

    const summary = {
      id: "summary",
      title: "generated_summary (резюме)",
      mode: "text",
      model: "deepseek-chat",
      temperature: 0.2,
      max_tokens: 260,
      system: commonSystem + " Ты умеешь делать краткое резюме текста, сохраняя смысл и структуру.",
      user:
        "Сделай краткое резюме уже написанного проекта (4–8 предложений) по теме: «{{topic}}».\n" +
        "Резюме должно отражать: что рассмотрено в главах и общий итог.\n\n" +
        "Текст проекта (хвост):\n{{generated_text_no_title_tail}}\n"
    };

    const appendix = {
      id: "appendix",
      title: "Приложение (JSON: таблицы/диаграммы)",
      mode: "json",
      model: "deepseek-chat",
      temperature: 0.35,
      max_tokens: 900,
      system:
        commonSystem +
        " Ты умеешь оформлять приложения: таблицы и диаграммы с аккуратными подписями, без выдумывания фактов. " +
        "Если точных данных нет — используй примерные данные и явно помечай как пример.",
      user:
        "Нужно подготовить раздел приложения (после списка литературы) для проекта по теме: «{{topic}}».\n" +
        "Сделай 1–3 объекта (в зависимости от уместности): таблицы и/или диаграммы.\n\n" +
        "Форматы объектов:\n" +
        "1) Таблица:\n" +
        "{\"type\":\"table\",\"title\":\"Таблица 1 – ...\",\"headers\":[...],\"rows\":[[...]...],\"notes\":\"(опционально)\"}\n" +
        "2) Диаграмма:\n" +
        "{\"type\":\"chart\",\"chartType\":\"bar|line|pie\",\"title\":\"Рисунок 1 – ...\",\"labels\":[...],\"series\":[{\"name\":\"...\",\"values\":[...] }],\"notes\":\"(опционально)\"}\n" +
        "3) Текстовый материал (если нужно):\n" +
        "{\"type\":\"text\",\"title\":\"Приложение А – ...\",\"content\":\"...\"}\n\n" +
        "Требования:\n" +
        "- Диаграммы должны опираться на практическую часть ниже.\n" +
        "- Если данных мало, сгенерируй небольшие примерные данные (и пометь в title или notes слово ‘пример’).\n" +
        "- Заголовки (title) делай короткими и понятными.\n" +
        "- Значения series.values — только числа.\n\n" +
        "Верни СТРОГО JSON без пояснений и без Markdown:\n" +
        "{\n  \"items\": [ ... ]\n}\n\n" +
        "Практическая часть (хвост):\n{{practical_text_tail}}\n"
    };

    return { plan, section, intro, conclusion, bibliography, summary, appendix };
  }

  function defaultPipeline(chaptersCount, defaultPromptIds) {
    const n = Math.max(1, Math.min(10, Number(chaptersCount || 2)));
    const dpi = defaultPromptIds || {
      plan: "plan",
      section: "section",
      intro: "intro",
      conclusion: "conclusion",
      bibliography: "bibliography",
      summary: "summary",
      appendix: "appendix"
    };

    const stages = [
      { id: "plan", type: "plan", enabled: true, label: "План", promptId: dpi.plan }
    ];
    for (let i = 0; i < n; i++) {
      stages.push({
        id: `ch_${i+1}`,
        type: "write_chapter",
        enabled: true,
        label: `Глава ${i+1}`,
        chapterIndex: i,
        promptId: dpi.section
      });
    }
    stages.push(
      { id: "intro", type: "intro", enabled: true, label: "Введение", promptId: dpi.intro },
      { id: "conclusion", type: "conclusion", enabled: true, label: "Заключение", promptId: dpi.conclusion },
      { id: "bibliography", type: "bibliography", enabled: true, label: "Источники", promptId: dpi.bibliography },
      { id: "appendix", type: "appendix", enabled: true, label: "Приложение", promptId: dpi.appendix }
    );
    return stages;
  }

  function makePreset(id, name, description, overrides = {}) {
    const prompts = defaultPrompts();

    const chaptersCount = overrides?.settings?.generation?.chaptersCount ?? 2;
    const sectionsPerChapter = overrides?.settings?.generation?.sectionsPerChapter ?? 4;

    const defaultPromptIds = overrides?.settings?.generation?.defaultPromptIds || {
      plan: "plan",
      section: "section",
      intro: "intro",
      conclusion: "conclusion",
      bibliography: "bibliography",
      summary: "summary",
      appendix: "appendix"
    };

    const settings = {
      autoAddContext: true,
      autoAddSummary: true,
      autoAddTail: false,
      tailChars: 12000,

      updateSummaryAfterEachStage: true,
      updateSummaryAfterEachSection: false,

      generation: {
        chaptersCount: chaptersCount,
        sectionsPerChapter: sectionsPerChapter,
        defaultPromptIds: defaultPromptIds,
        stages: defaultPipeline(chaptersCount, defaultPromptIds)
      },

      // Per-preset DOCX style profile
      docxStyle: normalizeDocxStyle(overrides?.settings?.docxStyle || defaultDocxStyle()),

      ...(overrides.settings || {})
    };

    // Ensure docxStyle always stays normalized (even if overrides.settings overwrote it)
    settings.docxStyle = normalizeDocxStyle(settings.docxStyle);

    return {
      id,
      name,
      description,
      settings,
      userVariables: { ...(overrides.userVariables || {}) },
      prompts: { ...prompts, ...(overrides.prompts || {}) }
    };
  }

  function defaultPresets() {
    return {
      school: makePreset(
        "school",
        "Школьный",
        "Лаконичный академический стиль. Удобно для 8–11 классов.",
        { settings: { autoAddTail: false, tailChars: 12000, generation: { chaptersCount: 2, sectionsPerChapter: 4 } } }
      ),
      college: makePreset(
        "college",
        "Колледж",
        "Чуть больше методики и практической части.",
        { settings: { autoAddTail: true, tailChars: 14000, generation: { chaptersCount: 2, sectionsPerChapter: 4 } } }
      ),
      coursework: makePreset(
        "coursework",
        "Курсовая",
        "Более развернутый текст (черновой режим).",
        { settings: { autoAddTail: true, tailChars: 18000, generation: { chaptersCount: 3, sectionsPerChapter: 4 } } }
      )
    };
  }

  function migratePreset(p) {
    p.settings = p.settings || {};
    // Per-preset DOCX style profile
    p.settings.docxStyle = normalizeDocxStyle(p.settings.docxStyle || defaultDocxStyle());
    p.settings.generation = p.settings.generation || { chaptersCount: 2, sectionsPerChapter: 4 };
    const gen = p.settings.generation;

    gen.chaptersCount = Number(gen.chaptersCount || 2);
    gen.sectionsPerChapter = Number(gen.sectionsPerChapter || 4);

    gen.defaultPromptIds = gen.defaultPromptIds || {
      plan: "plan",
      section: "section",
      intro: "intro",
      conclusion: "conclusion",
      bibliography: "bibliography",
      summary: "summary",
      appendix: "appendix"
    };

    // Ensure new keys exist even if defaultPromptIds already existed
    gen.defaultPromptIds.appendix = gen.defaultPromptIds.appendix || "appendix";

    p.prompts = p.prompts || defaultPrompts();
    const defs = defaultPrompts();
    for (const k of Object.keys(defs)) {
      if (!p.prompts[k]) p.prompts[k] = defs[k];
    }

    gen.stages = Array.isArray(gen.stages) ? gen.stages : defaultPipeline(gen.chaptersCount, gen.defaultPromptIds);
    if (!gen.stages.length) gen.stages = defaultPipeline(gen.chaptersCount, gen.defaultPromptIds);

    for (const st of gen.stages) {
      if (!st || typeof st !== "object") continue;
      if (!st.type) continue;

      if (!st.promptId) {
        if (st.type === "plan") st.promptId = gen.defaultPromptIds.plan;
        if (st.type === "write_chapter") st.promptId = gen.defaultPromptIds.section;
        if (st.type === "intro") st.promptId = gen.defaultPromptIds.intro;
        if (st.type === "conclusion") st.promptId = gen.defaultPromptIds.conclusion;
        if (st.type === "bibliography") st.promptId = gen.defaultPromptIds.bibliography;
        if (st.type === "summary") st.promptId = gen.defaultPromptIds.summary;
        if (st.type === "appendix") st.promptId = gen.defaultPromptIds.appendix;
      }
    }

    // Ensure appendix prompt exists
    const defs2 = defaultPrompts();
    if (!p.prompts.appendix) p.prompts.appendix = defs2.appendix;

    // Ensure appendix stage exists in pipeline (best-effort: add to end if missing)
    if (!gen.stages.some((s) => s && s.type === "appendix")) {
      gen.stages.push({
        id: `appendix_${Date.now()}`,
        type: "appendix",
        enabled: true,
        label: "Приложение",
        promptId: gen.defaultPromptIds.appendix
      });
    }
    return p;
  }

  function loadAllPresets() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultPresets();
    try {
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") return defaultPresets();
      if (Object.keys(parsed).length === 0) return defaultPresets();
      for (const k of Object.keys(parsed)) parsed[k] = migratePreset(parsed[k]);
      return parsed;
    } catch (e) {
      console.warn("PromptConfig: parse error; using defaults", e);
      return defaultPresets();
    }
  }

  function saveAllPresets(presets) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(presets));
  }

  function getActivePresetId() {
    return localStorage.getItem(ACTIVE_KEY) || "school";
  }

  function setActivePresetId(id) {
    localStorage.setItem(ACTIVE_KEY, id);
  }

  function getActivePreset() {
    const all = loadAllPresets();
    const id = getActivePresetId();
    return all[id] || all[Object.keys(all)[0]] || defaultPresets().school;
  }

  function renderTemplate(tpl, vars) {
    if (!tpl) return "";
    return tpl.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => {
      const v = vars && Object.prototype.hasOwnProperty.call(vars, key) ? vars[key] : "";
      return (v === undefined || v === null) ? "" : String(v);
    });
  }

  function extractJsonObject(text) {
    if (!text) return null;
    const s = String(text);
    const start = s.indexOf("{");
    if (start < 0) return null;
    let depth = 0;
    for (let i = start; i < s.length; i++) {
      const c = s[i];
      if (c === "{") depth++;
      else if (c === "}") depth--;
      if (depth === 0) {
        const candidate = s.slice(start, i + 1);
        try { return JSON.parse(candidate); } catch (_) { return null; }
      }
    }
    return null;
  }

  window.PromptConfig = {
    STORAGE_KEY,
    ACTIVE_KEY,
    DOCX_STYLE_SCHEMA_VERSION,
    BUILTIN_VARS,
    defaultDocxStyle,
    normalizeDocxStyle,
    defaultPresets,
    loadAllPresets,
    saveAllPresets,
    getActivePresetId,
    setActivePresetId,
    getActivePreset,
    renderTemplate,
    extractJsonObject,
    deepClone,
    defaultPipeline,
    defaultPrompts,
    makePreset,
    migratePreset
  };
})();
