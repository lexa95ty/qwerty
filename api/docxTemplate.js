const fs = require("fs/promises");
const path = require("path");
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");

const TEMPLATE_PATH = path.join(
  process.cwd(),
  "ffff",
  "school_project_template_v2.docx",
);
const TEMPLATE_FALLBACK_PATH = path.join(
  __dirname,
  "..",
  "ffff",
  "school_project_template_v2.docx",
);

let cachedTemplateBuffer = null;

async function loadTemplateBuffer() {
  if (cachedTemplateBuffer) {
    return cachedTemplateBuffer;
  }

  try {
    cachedTemplateBuffer = await fs.readFile(TEMPLATE_PATH);
    return cachedTemplateBuffer;
  } catch (error) {
    if (!error || error.code !== "ENOENT") {
      throw error;
    }
  }

  try {
    cachedTemplateBuffer = await fs.readFile(TEMPLATE_FALLBACK_PATH);
    return cachedTemplateBuffer;
  } catch (error) {
    if (!error || error.code !== "ENOENT") {
      throw error;
    }
  }

  throw new Error(
    `Template not found. Tried: ${TEMPLATE_PATH}; ${TEMPLATE_FALLBACK_PATH}`,
  );
}

async function openZipFromBuffer(buf) {
  return JSZip.loadAsync(buf);
}

async function readZipText(zip, entryPath) {
  const file = zip.file(entryPath);

  if (!file) {
    throw new Error(`Template entry not found: ${entryPath}`);
  }

  return file.async("string");
}

async function writeZipText(zip, entryPath, text) {
  zip.file(entryPath, text);
}

function normalizeSingleLine(value) {
  if (!value) {
    return "";
  }

  return String(value)
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stripMarkdownArtifacts(text) {
  if (!text) {
    return "";
  }

  const normalized = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  const withoutMarkdown = normalized
    .replace(/^\s{0,3}#{1,6}\s+/gm, "")
    .replace(/\*\*(.+?)\*\*/g, "$1")
    .replace(/__(.+?)__/g, "$1")
    .replace(/\*\*/g, "")
    .replace(/__/g, "")
    .replace(/^\s*[*-]\s+/gm, "")
    .replace(/^\s*>\s+/gm, "");

  return withoutMarkdown
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trimEnd())
    .join("\n")
    .trim();
}

function normalizeChapterTitle(title) {
  const normalized = normalizeSingleLine(title);
  const withoutPrefix = normalized
    .replace(/^(глава|раздел)\s*\d+(\.\d+)*\s*[.)-]?\s*/i, "")
    .replace(/^\d+(\.\d+)*\s*[.)-]?\s*/, "");

  return withoutPrefix.replace(/\.\s*$/, "").trim();
}

function normalizeSectionTitle(title) {
  const normalized = normalizeSingleLine(title);
  const withoutPrefix = normalized.replace(/^\d+(\.\d+)*\s*[.)-]?\s*/, "");

  return withoutPrefix.replace(/\.\s*$/, "").trim();
}

function normalizeHeadingForCompare(value) {
  if (!value) {
    return "";
  }

  const normalized = String(value)
    .toLowerCase()
    .replace(/[#!*_`]/g, "")
    .replace(/^(глава|раздел)\s*\d+(\.\d+)*\s*[.)-]?\s*/i, "")
    .replace(/^\d+(\.\d+)*\s*[.)-]?\s*/, "")
    .replace(/[^\p{L}\p{N}\s]/gu, " ");

  return normalized.replace(/\s+/g, " ").trim();
}

function stripDuplicateLeadingHeading(text, expectedTitle) {
  if (!text) {
    return "";
  }

  const normalizedText = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalizedText.split("\n");
  const firstContentIndex = lines.findIndex((line) => line.trim().length > 0);
  if (firstContentIndex === -1) {
    return normalizedText.trim();
  }

  const firstLine = lines[firstContentIndex];
  const firstLineTrimmed = firstLine.trim();
  const isMarkdownHeading = /^\s{0,3}#{1,6}\s+/.test(firstLine);
  const expectedNormalized = normalizeHeadingForCompare(expectedTitle);
  const lineNormalized = normalizeHeadingForCompare(firstLineTrimmed);

  const extractLeadingNumber = (value) => {
    const match = String(value || "").match(
      /^\s*(?:#{1,6}\s*)?(\d+(?:\.\d+)+|\d+)\b/,
    );
    return match ? match[1] : "";
  };

  const lineNumberId = extractLeadingNumber(firstLineTrimmed);
  const expectedNumberId = extractLeadingNumber(expectedTitle);
  const matchesExpected =
    (lineNormalized &&
      expectedNormalized &&
      lineNormalized === expectedNormalized) ||
    (lineNumberId &&
      expectedNumberId &&
      lineNumberId === expectedNumberId &&
      lineNormalized);

  if (isMarkdownHeading || matchesExpected) {
    lines.splice(firstContentIndex, 1);
    if (lines[firstContentIndex] !== undefined && !lines[firstContentIndex].trim()) {
      lines.splice(firstContentIndex, 1);
    }
  }

  return lines.join("\n").trim();
}

function getParagraphText(paragraph) {
  const textNodes = Array.from(paragraph.getElementsByTagName("w:t"));
  return textNodes.map((node) => node.textContent || "").join("");
}

function looksLikeTableSeparator(line) {
  const trimmed = String(line ?? "").trim();
  if (!trimmed.includes("|")) {
    return false;
  }

  const normalized = trimmed.replace(/^\|/, "").replace(/\|$/, "");
  const parts = normalized.split("|").map((part) => part.trim());
  if (parts.length < 2) {
    return false;
  }

  return parts.every((part) => /^:?-{3,}:?$/.test(part));
}

function parseTableRow(line) {
  return String(line ?? "")
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => String(cell).trim());
}

function parseTextToBlocks(text) {
  const normalized = String(text ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  const blocks = [];
  let paragraphLines = [];
  const sectionLabels = new Set([
    "Актуальность:",
    "Цель:",
    "Задачи:",
    "Гипотеза:",
    "Методы:",
    "Объект:",
    "Предмет:",
  ]);

  function flushParagraph() {
    if (paragraphLines.length === 0) {
      return;
    }
    const paragraph = paragraphLines
      .map((line) => line.trim())
      .filter(Boolean)
      .join(" ")
      .trim();
    if (paragraph) {
      blocks.push({ type: "p", text: paragraph });
    }
    paragraphLines = [];
  }

  for (let index = 0; index < lines.length; index += 1) {
    const line = String(lines[index] ?? "");
    const trimmedLine = line.trim();
    if (!trimmedLine) {
      flushParagraph();
      continue;
    }

    if (sectionLabels.has(trimmedLine)) {
      flushParagraph();
      blocks.push({ type: "p", text: trimmedLine });
      continue;
    }

    if (/^\d+\.\s+/.test(trimmedLine)) {
      flushParagraph();
      blocks.push({ type: "p", text: trimmedLine });
      continue;
    }

    const nextLine = lines[index + 1];
    if (line.includes("|") && looksLikeTableSeparator(nextLine)) {
      flushParagraph();
      const headers = parseTableRow(line);
      index += 2;
      const rows = [];
      while (index < lines.length) {
        const rowLine = String(lines[index] ?? "");
        if (!rowLine.trim()) {
          break;
        }
        if (!rowLine.includes("|")) {
          break;
        }
        rows.push(parseTableRow(rowLine));
        index += 1;
      }
      index -= 1;
      blocks.push({ type: "table", headers, rows });
      continue;
    }

    paragraphLines.push(line);
  }

  flushParagraph();

  if (blocks.length === 0) {
    return [{ type: "p", text: "" }];
  }

  return blocks;
}

function replaceParagraphTextWithSingleRun(doc, paragraph, text) {
  const firstRun = paragraph.getElementsByTagName("w:r")[0];
  const firstRunProps = firstRun
    ? firstRun.getElementsByTagName("w:rPr")[0]
    : null;

  Array.from(paragraph.childNodes).forEach((node) => {
    if (node.nodeName !== "w:pPr") {
      paragraph.removeChild(node);
    }
  });

  const run = doc.createElement("w:r");
  if (firstRunProps) {
    run.appendChild(firstRunProps.cloneNode(true));
  }

  const textNode = doc.createElement("w:t");
  if (/^\s|\s$/.test(text)) {
    textNode.setAttribute("xml:space", "preserve");
  }
  textNode.appendChild(doc.createTextNode(text));
  run.appendChild(textNode);
  paragraph.appendChild(run);
}

function replaceParagraphTextPreservingRuns(paragraph, text) {
  const textNodes = Array.from(paragraph.getElementsByTagName("w:t"));
  if (textNodes.length === 0) {
    return;
  }

  const firstTextNode = textNodes[0];
  firstTextNode.textContent = text;
  if (/^\s|\s$/.test(text)) {
    firstTextNode.setAttribute("xml:space", "preserve");
  } else {
    firstTextNode.removeAttribute("xml:space");
  }

  const runsToCheck = new Set();

  textNodes.slice(1).forEach((node) => {
    node.textContent = "";
    node.removeAttribute("xml:space");
    if (node.parentNode && node.parentNode.nodeName === "w:r") {
      runsToCheck.add(node.parentNode);
    }
  });

  runsToCheck.forEach((run) => {
    const elementChildren = Array.from(run.childNodes).filter(
      (child) => child.nodeType === 1,
    );
    const hasNonEmptyText = elementChildren.some(
      (child) => child.nodeName === "w:t" && (child.textContent || "") !== "",
    );
    const hasNonTextContent = elementChildren.some(
      (child) => child.nodeName !== "w:t" && child.nodeName !== "w:rPr",
    );
    if (!hasNonEmptyText && !hasNonTextContent && run.parentNode) {
      run.parentNode.removeChild(run);
    }
  });
}

function createTable(doc, headers, rows) {
  const table = doc.createElement("w:tbl");
  const tblPr = doc.createElement("w:tblPr");
  const tblW = doc.createElement("w:tblW");
  tblW.setAttribute("w:type", "auto");
  tblW.setAttribute("w:w", "0");
  tblPr.appendChild(tblW);
  table.appendChild(tblPr);

  function makeCell(text) {
    const tc = doc.createElement("w:tc");
    const tcPr = doc.createElement("w:tcPr");
    const tcW = doc.createElement("w:tcW");
    tcW.setAttribute("w:type", "auto");
    tcW.setAttribute("w:w", "0");
    tcPr.appendChild(tcW);
    tc.appendChild(tcPr);

    const p = doc.createElement("w:p");
    const r = doc.createElement("w:r");
    const tNode = doc.createElement("w:t");
    const value = String(text ?? "");
    if (/^\s|\s$/.test(value)) {
      tNode.setAttribute("xml:space", "preserve");
    }
    tNode.appendChild(doc.createTextNode(value));
    r.appendChild(tNode);
    p.appendChild(r);
    tc.appendChild(p);
    return tc;
  }

  if (headers && headers.length > 0) {
    const headerRow = doc.createElement("w:tr");
    headers.forEach((cell) => {
      headerRow.appendChild(makeCell(cell));
    });
    table.appendChild(headerRow);
  }

  (rows || []).forEach((row) => {
    const tr = doc.createElement("w:tr");
    (row || []).forEach((cell) => {
      tr.appendChild(makeCell(cell));
    });
    table.appendChild(tr);
  });

  return table;
}

function renderSources(doc, sourcesArray) {
  const paragraphs = Array.from(doc.getElementsByTagName("w:p"));
  const placeholderNumbers = [1, 2, 3, 4, 5];
  const sourceParagraphs = new Map();
  let totalSourceParagraphs = 0;

  paragraphs.forEach((paragraph, index) => {
    const visibleText = getParagraphText(paragraph);
    const matches = visibleText.match(/\{SOURCE_([1-5])\}/g) || [];
    if (matches.length === 0) {
      return;
    }

    if (matches.length > 1) {
      const error = new Error(
        "Source placeholders must be in separate paragraphs.",
      );
      error.statusCode = 500;
      throw error;
    }

    const number = Number(matches[0].replace(/\D+/g, ""));
    if (sourceParagraphs.has(number)) {
      const error = new Error(
        "Source placeholders must appear exactly once each.",
      );
      error.statusCode = 500;
      throw error;
    }

    sourceParagraphs.set(number, { paragraph, index });
    totalSourceParagraphs += 1;
  });

  if (totalSourceParagraphs !== 5) {
    const error = new Error(
      "Expected exactly five source placeholder paragraphs in a row.",
    );
    error.statusCode = 500;
    throw error;
  }

  const orderedParagraphs = placeholderNumbers.map((number) => {
    const entry = sourceParagraphs.get(number);
    if (!entry) {
      const error = new Error(
        "Expected {SOURCE_1}..{SOURCE_5} placeholders in consecutive paragraphs.",
      );
      error.statusCode = 500;
      throw error;
    }
    return entry;
  });

  const firstIndex = orderedParagraphs[0].index;
  const isConsecutive = orderedParagraphs.every(
    (entry, offset) => entry.index === firstIndex + offset,
  );
  if (!isConsecutive) {
    const error = new Error(
      "Source placeholders must be placed in five consecutive paragraphs.",
    );
    error.statusCode = 500;
    throw error;
  }

  const sources = Array.isArray(sourcesArray) ? sourcesArray : [];
  if (sources.length > 20) {
    console.warn("Sources list has more than 20 entries.");
  }
  const cappedSources = sources.slice(0, 20);
  const count = Math.min(cappedSources.length, 20);

  const orderedParagraphNodes = orderedParagraphs.map(
    (entry) => entry.paragraph,
  );
  const parent = orderedParagraphNodes[0].parentNode;

  if (count === 0) {
    replaceParagraphTextWithSingleRun(doc, orderedParagraphNodes[0], "—");
    orderedParagraphNodes.slice(1).forEach((paragraph) => {
      if (paragraph.parentNode) {
        paragraph.parentNode.removeChild(paragraph);
      }
    });
  } else if (count <= 5) {
    orderedParagraphNodes.forEach((paragraph, index) => {
      if (index < count) {
        replaceParagraphTextWithSingleRun(doc, paragraph, cappedSources[index]);
      } else if (paragraph.parentNode) {
        paragraph.parentNode.removeChild(paragraph);
      }
    });
  } else {
    orderedParagraphNodes.forEach((paragraph, index) => {
      replaceParagraphTextWithSingleRun(doc, paragraph, cappedSources[index]);
    });

    let insertAfter = orderedParagraphNodes[4];
    for (let index = 5; index < count; index += 1) {
      const cloned = orderedParagraphNodes[4].cloneNode(true);
      replaceParagraphTextWithSingleRun(doc, cloned, cappedSources[index]);
      const nextSibling = insertAfter.nextSibling;
      if (nextSibling) {
        parent.insertBefore(cloned, nextSibling);
      } else {
        parent.appendChild(cloned);
      }
      insertAfter = cloned;
    }
  }

  const remainingSources = Array.from(doc.getElementsByTagName("w:p")).filter(
    (paragraph) => /\{SOURCE_[1-5]\}/.test(getParagraphText(paragraph)),
  );
  if (remainingSources.length > 0) {
    const error = new Error(
      "Source placeholders were not fully replaced in the document.",
    );
    error.statusCode = 500;
    throw error;
  }
}

function ensureChildElement(doc, parent, tagName, insertBeforeNode = null) {
  const existing = parent.getElementsByTagName(tagName)[0];
  if (existing) {
    return existing;
  }

  const created = doc.createElement(tagName);
  if (insertBeforeNode) {
    parent.insertBefore(created, insertBeforeNode);
  } else {
    parent.appendChild(created);
  }
  return created;
}

function enforceTableStyle(doc) {
  const tables = Array.from(doc.getElementsByTagName("w:tbl"));

  tables.forEach((table) => {
    const cells = Array.from(table.getElementsByTagName("w:tc"));
    cells.forEach((cell) => {
      const paragraphs = Array.from(cell.getElementsByTagName("w:p"));
      paragraphs.forEach((paragraph) => {
        const firstChild = paragraph.firstChild;
        const paragraphProps = ensureChildElement(
          doc,
          paragraph,
          "w:pPr",
          firstChild,
        );
        const align = ensureChildElement(doc, paragraphProps, "w:jc");
        align.setAttribute("w:val", "left");
        const indent = ensureChildElement(doc, paragraphProps, "w:ind");
        indent.setAttribute("w:left", "0");
        indent.setAttribute("w:firstLine", "0");
      });
    });

    const runs = Array.from(table.getElementsByTagName("w:r"));
    runs.forEach((run) => {
      const firstChild = run.firstChild;
      const runProps = ensureChildElement(doc, run, "w:rPr", firstChild);
      const size = ensureChildElement(doc, runProps, "w:sz");
      size.setAttribute("w:val", "28");
      const color = ensureChildElement(doc, runProps, "w:color");
      color.setAttribute("w:val", "000000");
    });
  });
}

async function updateSettingsXml(zip) {
  const existingSettings = zip.file("word/settings.xml");
  const settingsXml = existingSettings
    ? await readZipText(zip, "word/settings.xml")
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>`;
  const parser = new DOMParser();
  const settingsDoc = parser.parseFromString(settingsXml, "application/xml");
  const settingsNode = settingsDoc.getElementsByTagName("w:settings")[0];

  if (!settingsNode) {
    throw new Error("Settings root <w:settings> not found.");
  }

  const updateFieldsNodes = Array.from(
    settingsNode.getElementsByTagName("w:updateFields"),
  );
  if (updateFieldsNodes.length > 0) {
    updateFieldsNodes.forEach((node) => {
      node.setAttribute("w:val", "true");
    });
  } else {
    const updateFields = settingsDoc.createElement("w:updateFields");
    updateFields.setAttribute("w:val", "true");
    settingsNode.appendChild(updateFields);
  }

  const serializer = new XMLSerializer();
  const updatedSettingsXml = serializer.serializeToString(settingsDoc);
  await writeZipText(zip, "word/settings.xml", updatedSettingsXml);
}

function markTocFieldsDirty(doc) {
  const paragraphs = Array.from(doc.getElementsByTagName("w:p"));

  paragraphs.forEach((paragraph) => {
    const instrNodes = Array.from(paragraph.getElementsByTagName("w:instrText"));
    const hasTocField = instrNodes.some((node) =>
      String(node.textContent || "").includes("TOC"),
    );
    if (!hasTocField) {
      return;
    }

    const fieldChars = Array.from(paragraph.getElementsByTagName("w:fldChar"));
    const beginField = fieldChars.find(
      (node) => node.getAttribute("w:fldCharType") === "begin",
    );
    if (beginField) {
      beginField.setAttribute("w:dirty", "true");
    }
  });
}

async function renderDocxFromTemplate(templateBuffer, templateData) {
  const zip = await openZipFromBuffer(templateBuffer);
  await updateSettingsXml(zip);
  const documentXml = await readZipText(zip, "word/document.xml");
  const parser = new DOMParser();
  const doc = parser.parseFromString(documentXml, "application/xml");
  renderSources(doc, templateData?.__sources || []);
  markTocFieldsDirty(doc);
  const paragraphs = Array.from(doc.getElementsByTagName("w:p"));
  const placeholderPattern = /\{[A-Z0-9_]+\}/g;

  paragraphs.forEach((paragraph) => {
    const visibleText = getParagraphText(paragraph);
    const trimmedText = visibleText.trim();
    const blockMatch = trimmedText.match(/^\{[A-Z0-9_]+\}$/);

    if (blockMatch) {
      const placeholder = blockMatch[0];
      const key = placeholder.slice(1, -1);
      if (!key.startsWith("SOURCE_")) {
        const value = templateData?.[key];
        if (value !== undefined && value !== null) {
          const blocks = parseTextToBlocks(value);
          const parent = paragraph.parentNode;
          let insertedTable = false;
          blocks.forEach((block) => {
            if (block.type === "table") {
              parent.insertBefore(
                createTable(doc, block.headers || [], block.rows || []),
                paragraph,
              );
              insertedTable = true;
              return;
            }

            const cloned = paragraph.cloneNode(true);
            replaceParagraphTextWithSingleRun(
              doc,
              cloned,
              String(block.text ?? ""),
            );
            parent.insertBefore(cloned, paragraph);
          });
          parent.removeChild(paragraph);
          if (insertedTable) {
            enforceTableStyle(doc);
          }
          return;
        }
      }
    }

    const placeholders = visibleText.match(placeholderPattern);
    if (!placeholders) {
      return;
    }

    let updatedText = visibleText;
    new Set(placeholders).forEach((placeholder) => {
      const key = placeholder.slice(1, -1);
      if (key.startsWith("SOURCE_")) {
        return;
      }
      if (Object.prototype.hasOwnProperty.call(templateData || {}, key)) {
        updatedText = updatedText.split(placeholder).join(
          String(templateData[key] ?? ""),
        );
      }
    });

    if (updatedText !== visibleText) {
      replaceParagraphTextPreservingRuns(paragraph, updatedText);
    }
  });

  enforceTableStyle(doc);

  const serializer = new XMLSerializer();
  const updatedXml = serializer.serializeToString(doc);
  await writeZipText(zip, "word/document.xml", updatedXml);

  const remainingPlaceholders = updatedXml.match(placeholderPattern) || [];
  const uniqueRemaining = Array.from(new Set(remainingPlaceholders));
  if (uniqueRemaining.length > 0) {
    console.error("Remaining placeholders:", uniqueRemaining);
    throw new Error(
      `Unresolved placeholders: ${uniqueRemaining.join(", ")}`,
    );
  }

  return zip.generateAsync({ type: "nodebuffer" });
}

function buildTemplateData(payload) {
  const data = payload ?? {};
  const title = data.title ?? {};
  const dash = "—";
  const currentYear = String(new Date().getFullYear());
  const safeTopic = normalizeSingleLine(data.topic) || dash;

  const organizationLines = normalizeSingleLine(title.organization)
    ? title.organization.split(/\r?\n/)
    : [];
  const orgName = normalizeSingleLine(organizationLines[0]);
  const orgAddress = normalizeSingleLine(organizationLines[1]);

  const schoolName = normalizeSingleLine(title.schoolName) || orgName || dash;
  const schoolAddress = normalizeSingleLine(title.schoolAddress) || orgAddress;

  const sanitizeTitle = (value) => stripMarkdownArtifacts(value);
  const sanitizeText = (value) => stripMarkdownArtifacts(value);

  const templateData = {
    WORK_TITLE: safeTopic,
    STUDENT_NAME: normalizeSingleLine(title.student) || dash,
    CLASS: normalizeSingleLine(title.class) || dash,
    SUPERVISOR_NAME: normalizeSingleLine(title.teacher) || dash,
    SUBJECT:
      normalizeSingleLine(title.subject ?? title.teacherSubject) || dash,
    CITY: normalizeSingleLine(title.city) || dash,
    YEAR: normalizeSingleLine(title.year) || currentYear,
    SCHOOL_NAME: schoolName,
    SCHOOL_ADDRESS: schoolAddress,
    INTRO: sanitizeText(data.introText) || dash,
    CONCLUSION: sanitizeText(data.conclusionText) || dash,
  };

  const chapters = Array.isArray(data.chapters) ? data.chapters : [];

  for (let chapterIndex = 0; chapterIndex < 2; chapterIndex += 1) {
    const chapterNumber = chapterIndex + 1;
    const chapter = chapters[chapterIndex];
    const chapterTitleKey = `CH${chapterNumber}_TITLE`;
    const chapterTitleValue = chapter?.title
      ? normalizeChapterTitle(sanitizeTitle(chapter.title))
      : "";

    if (!chapterTitleValue) {
      console.warn(`Missing data for ${chapterTitleKey}`);
    }

    templateData[chapterTitleKey] = chapterTitleValue || dash;

    for (let sectionIndex = 0; sectionIndex < 4; sectionIndex += 1) {
      const sectionNumber = sectionIndex + 1;
      const section = chapter?.sections?.[sectionIndex];
      const titleKey = `CH${chapterNumber}_${sectionNumber}_TITLE`;
      const textKey = `CH${chapterNumber}_${sectionNumber}_TEXT`;
      const sectionTitleValue = section?.title
        ? normalizeSectionTitle(sanitizeTitle(section.title))
        : "";
      const sectionTextValue = section?.content ?? "";

      if (!sectionTitleValue) {
        console.warn(`Missing data for ${titleKey}`);
      }

      if (!sectionTextValue) {
        console.warn(`Missing data for ${textKey}`);
      }

      templateData[titleKey] = sectionTitleValue || dash;
      const sanitizedSectionText = stripDuplicateLeadingHeading(
        sanitizeText(sectionTextValue),
        sectionTitleValue,
      );
      templateData[textKey] = sanitizedSectionText || dash;
    }
  }

  const appendix = Array.isArray(data.appendix) ? data.appendix : [];
  const firstAppendix = appendix[0] ?? {};
  const secondAppendix = appendix[1] ?? {};

  templateData.APP1_TITLE = normalizeSingleLine(sanitizeTitle(firstAppendix.title));
  templateData.APP1_CONTENT = sanitizeText(firstAppendix.content);
  templateData.APP2_TITLE = normalizeSingleLine(sanitizeTitle(secondAppendix.title));
  templateData.APP2_CONTENT = sanitizeText(secondAppendix.content);

  if (appendix.length > 2) {
    const extraEntries = appendix.slice(2).map((entry) => {
      const pieces = [];

      if (entry?.title) {
        pieces.push(sanitizeTitle(entry.title));
      }

      if (entry?.content) {
        pieces.push(sanitizeText(entry.content));
      }

      return pieces.join("\n\n");
    });

    const extraContent = extraEntries.filter(Boolean).join("\n\n---\n\n");

    if (extraContent) {
      templateData.APP2_CONTENT = templateData.APP2_CONTENT
        ? `${templateData.APP2_CONTENT}\n\n---\n\n${extraContent}`
        : extraContent;
    }
  }

  const bibliography = Array.isArray(data.bibliography)
    ? data.bibliography
    : [];
  const sources = bibliography
    .map((entry) => normalizeSingleLine(entry))
    .filter(Boolean);

  templateData.__sources = sources.slice(0, 20);

  return templateData;
}

module.exports = {
  TEMPLATE_PATH,
  loadTemplateBuffer,
  openZipFromBuffer,
  readZipText,
  writeZipText,
  normalizeSingleLine,
  normalizeChapterTitle,
  normalizeSectionTitle,
  buildTemplateData,
  renderDocxFromTemplate,
};
