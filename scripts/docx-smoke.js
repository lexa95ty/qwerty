const crypto = require("crypto");
const assert = require("assert");
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");
const {
  loadTemplateBuffer,
  buildTemplateData,
  renderDocxFromTemplate,
} = require("../api/docxTemplate");

function sha256(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

function buildParagraphs(prefix, count) {
  return Array.from({ length: count }, (_, index) => {
    const number = index + 1;
    return `${prefix} paragraph ${number}.`;
  }).join("\n\n");
}

function collectHeadingStyleIds(stylesDoc) {
  const styles = Array.from(stylesDoc.getElementsByTagName("w:style"));
  const ids = new Set();

  styles.forEach((style) => {
    if ((style.getAttribute("w:type") || "").toLowerCase() !== "paragraph") {
      return;
    }
    const styleId = style.getAttribute("w:styleId");
    const nameNode = style.getElementsByTagName("w:name")[0];
    const name = nameNode ? nameNode.getAttribute("w:val") || "" : "";
    const normalized = name.toLowerCase();
    if (
      normalized.includes("heading 1") ||
      normalized.includes("заголовок 1") ||
      styleId === "Heading1" ||
      styleId === "1"
    ) {
      ids.add(styleId);
    }
  });

  return ids;
}

function assertHeadingStylesHavePageBreak(stylesDoc, headingStyleIds) {
  headingStyleIds.forEach((styleId) => {
    const style = Array.from(stylesDoc.getElementsByTagName("w:style")).find(
      (node) => node.getAttribute("w:styleId") === styleId,
    );
    assert.ok(style, `Heading style ${styleId} is missing`);
    const paragraphProps = style.getElementsByTagName("w:pPr")[0];
    assert.ok(
      paragraphProps &&
        paragraphProps.getElementsByTagName("w:pageBreakBefore").length > 0,
      `Heading style ${styleId} must include pageBreakBefore`,
    );
  });
}

function isHeadingParagraph(paragraph, headingStyleIds) {
  const paragraphProps = paragraph.getElementsByTagName("w:pPr")[0];
  if (!paragraphProps) {
    return false;
  }
  const styleNode = paragraphProps.getElementsByTagName("w:pStyle")[0];
  if (!styleNode) {
    return false;
  }
  return headingStyleIds.has(styleNode.getAttribute("w:val"));
}

function paragraphHasPageBreak(paragraph) {
  const breaks = Array.from(paragraph.getElementsByTagName("w:br"));
  return breaks.some((node) => {
    const type = node.getAttribute("w:type");
    return !type || type === "page" || type === "section";
  });
}

function paragraphHasVisibleText(paragraph) {
  const texts = Array.from(paragraph.getElementsByTagName("w:t"));
  return texts.some((node) => (node.textContent || "").trim() !== "");
}

function isStandalonePageBreakParagraph(paragraph) {
  if (!paragraphHasPageBreak(paragraph)) {
    return false;
  }
  if (paragraphHasVisibleText(paragraph)) {
    return false;
  }
  const allowed = new Set(["w:pPr", "w:r", "w:rPr", "w:br", "w:lastRenderedPageBreak"]);
  const allNodes = Array.from(paragraph.getElementsByTagName("*"));
  return allNodes.every((node) => allowed.has(node.nodeName));
}

function assertNoStandaloneBreakBeforeHeadings(documentXml, headingStyleIds) {
  const doc = new DOMParser().parseFromString(documentXml, "application/xml");
  const body = doc.getElementsByTagName("w:body")[0];
  if (!body) {
    return;
  }
  const children = Array.from(body.childNodes);

  children.forEach((node, index) => {
    if (node.nodeType !== 1 || node.nodeName !== "w:p") {
      return;
    }
    if (!isHeadingParagraph(node, headingStyleIds)) {
      return;
    }
    for (let prevIndex = index - 1; prevIndex >= 0; prevIndex -= 1) {
      const prev = children[prevIndex];
      if (prev.nodeType !== 1) {
        continue;
      }
      if (prev.nodeName !== "w:p") {
        break;
      }
      assert.ok(
        !isStandalonePageBreakParagraph(prev),
        "Standalone page break paragraph found directly before Heading 1",
      );
      break;
    }
  });
}

function paragraphText(paragraph) {
  const texts = Array.from(paragraph.getElementsByTagName("w:t"));
  return texts.map((node) => node.textContent || "").join("");
}

function isInsideTable(node) {
  let current = node;
  while (current) {
    if (current.nodeName === "w:tbl") {
      return true;
    }
    current = current.parentNode;
  }
  return false;
}

function getFirstLineIndent(paragraph) {
  const paragraphProps = paragraph.getElementsByTagName("w:pPr")[0];
  const indent = paragraphProps?.getElementsByTagName("w:ind")[0];
  const value = indent?.getAttribute("w:firstLine");
  return value != null ? Number(value) : null;
}

function assertParagraphsHaveFirstLineIndent(paragraphs, label) {
  paragraphs.forEach((paragraph) => {
    const firstLine = getFirstLineIndent(paragraph);
    assert.ok(
      Number.isFinite(firstLine) && firstLine > 0,
      `${label} paragraph is missing first line indent`,
    );
  });
}

function buildSamplePayload() {
  const chapters = Array.from({ length: 2 }, (_, chapterIndex) => {
    const chapterNumber = chapterIndex + 1;
    const sections = Array.from({ length: 4 }, (_, sectionIndex) => {
      const sectionNumber = sectionIndex + 1;
      return {
        title: `Глава ${chapterNumber}.${sectionNumber} Раздел`,
        content: buildParagraphs(
          `Chapter ${chapterNumber} section ${sectionNumber}`,
          2,
        ),
      };
    });

    return {
      title: `Глава ${chapterNumber} Тестовая тема`,
      sections,
    };
  });

  const bibliography = Array.from({ length: 12 }, (_, index) => {
    return `Источник ${index + 1}: Автор ${index + 1}, Название.`;
  });

  return {
    topic: "Проверка шаблона DOCX",
    title: {
      student: "Тестовый Студент",
      class: "10A",
      teacher: "Преподаватель",
      subject: "Информатика",
      city: "Москва",
      year: String(new Date().getFullYear()),
      schoolName: "Школа №1",
      schoolAddress: "ул. Примерная, 1",
    },
    introText: buildParagraphs("Intro", 5),
    conclusionText: buildParagraphs("Conclusion", 5),
    chapters,
    bibliography,
    appendix: [
      {
        title: "Приложение A",
        content: buildParagraphs("Appendix A", 2),
      },
      {
        title: "Приложение B",
        content: buildParagraphs("Appendix B", 2),
      },
    ],
  };
}

async function main() {
  const templateBuffer = await loadTemplateBuffer();
  const templateHashBefore = sha256(templateBuffer);

  const payload = buildSamplePayload();
  const templateData = buildTemplateData(payload);
  const resultBuffer = await renderDocxFromTemplate(
    templateBuffer,
    templateData,
  );

  const templateBufferAfter = await loadTemplateBuffer();
  const templateHashAfter = sha256(templateBufferAfter);

  assert.strictEqual(
    templateHashAfter,
    templateHashBefore,
    "Template hash changed after rendering",
  );

  const zip = await JSZip.loadAsync(resultBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const settingsXml = await zip.file("word/settings.xml").async("string");
  const stylesXml = await zip.file("word/styles.xml").async("string");
  const stylesDoc = new DOMParser().parseFromString(stylesXml, "application/xml");

  const textWithoutTags = documentXml.replace(/<[^>]+>/g, "");
  assert.ok(
    !/{[A-Z0-9_]+}/.test(textWithoutTags),
    "Unresolved placeholders found in document.xml",
  );

  assert.ok(
    /<w:updateFields[^>]*w:val="true"/.test(settingsXml),
    "settings.xml does not enable updateFields",
  );

  const styleMatches = documentXml.match(/w:pStyle w:val="a"/g) || [];
  assert.strictEqual(
    styleMatches.length,
    12,
    "Unexpected number of bibliography styles",
  );

  const headingStyleIds = collectHeadingStyleIds(stylesDoc);
  assert.ok(headingStyleIds.size > 0, "Heading 1 style not found");
  assertHeadingStylesHavePageBreak(stylesDoc, headingStyleIds);
  assertNoStandaloneBreakBeforeHeadings(documentXml, headingStyleIds);

  const doc = new DOMParser().parseFromString(documentXml, "application/xml");
  const paragraphs = Array.from(doc.getElementsByTagName("w:p")).filter(
    (paragraph) => !isInsideTable(paragraph),
  );
  const introParagraphs = paragraphs.filter((paragraph) =>
    /Intro paragraph \d/.test(paragraphText(paragraph)),
  );
  const conclusionParagraphs = paragraphs.filter((paragraph) =>
    /Conclusion paragraph \d/.test(paragraphText(paragraph)),
  );
  const sectionParagraphs = paragraphs.filter((paragraph) =>
    /Chapter \d section \d paragraph \d/.test(paragraphText(paragraph)),
  );
  const appendixParagraphs = paragraphs.filter((paragraph) =>
    /Appendix [AB] paragraph \d/.test(paragraphText(paragraph)),
  );

  assertParagraphsHaveFirstLineIndent(introParagraphs, "Intro");
  assertParagraphsHaveFirstLineIndent(conclusionParagraphs, "Conclusion");
  assertParagraphsHaveFirstLineIndent(sectionParagraphs, "Section");
  assertParagraphsHaveFirstLineIndent(appendixParagraphs, "Appendix");

  const headingParagraphs = paragraphs.filter((paragraph) =>
    isHeadingParagraph(paragraph, headingStyleIds),
  );
  headingParagraphs.forEach((paragraph) => {
    const firstLine = getFirstLineIndent(paragraph);
    assert.ok(
      !firstLine || firstLine <= 0,
      "Heading 1 paragraph must not gain first line indent",
    );
  });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
