const crypto = require("crypto");
const assert = require("assert");
const JSZip = require("jszip");
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
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
