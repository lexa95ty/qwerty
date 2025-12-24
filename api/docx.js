const { buildTemplateData, loadTemplateBuffer, renderDocxFromTemplate } = require("./docxTemplate");

function safeString(value) {
  return String(value ?? "");
}

function sanitizeAsciiFilename(name) {
  const base = safeString(name || "project")
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
    .replace(/\s+/g, "_")
    .slice(0, 80) || "project";
  return base.replace(/[^a-zA-Z0-9._-]+/g, "_") || "project";
}

function encodeRFC5987(str) {
  return encodeURIComponent(safeString(str)).replace(/[!'()*]/g, (c) =>
    "%" + c.charCodeAt(0).toString(16).toUpperCase(),
  );
}

function contentDisposition(filenameUtf8) {
  const ascii = sanitizeAsciiFilename(filenameUtf8).toLowerCase().endsWith(".docx")
    ? sanitizeAsciiFilename(filenameUtf8)
    : sanitizeAsciiFilename(filenameUtf8) + ".docx";
  const encoded = encodeRFC5987(filenameUtf8);
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encoded}`;
}

function parseJsonPayload(body) {
  if (body == null) {
    return {};
  }

  let payload = body;
  if (Buffer.isBuffer(payload)) {
    payload = payload.toString("utf8");
  }

  if (typeof payload === "string") {
    try {
      payload = JSON.parse(payload);
    } catch (error) {
      const err = new Error("Invalid JSON payload");
      err.statusCode = 400;
      throw err;
    }
  }

  if (payload && typeof payload === "object") {
    return payload;
  }

  return {};
}

function logPayloadWarnings(payload) {
  const chapters = Array.isArray(payload?.chapters) ? payload.chapters : [];
  if (chapters.length < 2) {
    console.warn("Chapters length is less than 2; template expects two chapters.");
  }

  chapters.slice(0, 2).forEach((chapter, index) => {
    const sections = Array.isArray(chapter?.sections) ? chapter.sections : [];
    if (sections.length < 4) {
      console.warn(
        `Chapter ${index + 1} has fewer than 4 sections; template expects 4 sections.`,
      );
    }
  });

  const bibliography = Array.isArray(payload?.bibliography) ? payload.bibliography : [];
  if (bibliography.length > 20) {
    console.warn(
      "Bibliography has more than 20 entries; extra entries will be ignored.",
    );
  }

  const appendix = Array.isArray(payload?.appendix) ? payload.appendix : [];
  if (appendix.length > 2) {
    console.warn(
      "Appendix has more than 2 entries; extra entries will be merged into Appendix 2 or trimmed.",
    );
  }
}

module.exports = async function handler(req, res) {
  try {
    console.log("DOCX endpoint hit", new Date().toISOString());
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method Not Allowed" });
    }

    let payload;
    try {
      payload = parseJsonPayload(req.body);
    } catch (err) {
      if (err?.statusCode === 400) {
        return res.status(400).json({ error: err.message });
      }
      throw err;
    }
    if (
      !payload ||
      (typeof payload === "object" && !Array.isArray(payload) && Object.keys(payload).length === 0)
    ) {
      return res.status(400).json({
        error: "Invalid payload",
        message: "Payload is empty or invalid.",
      });
    }

    logPayloadWarnings(payload);

    const topic = safeString(payload.topic || "").trim();
    const filenameUtf8 = `Проект_${topic || "project"}.docx`;

    const templateBuffer = await loadTemplateBuffer();
    const templateData = buildTemplateData(payload);
    const buf = await renderDocxFromTemplate(templateBuffer, templateData);

    res.statusCode = 200;
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    res.setHeader("Content-Disposition", contentDisposition(filenameUtf8));
    res.setHeader("Content-Length", String(buf.length));
    res.setHeader("Cache-Control", "no-store");
    res.end(buf);
  } catch (err) {
    console.error("DOCX_EXPORT_FAILED", err.stack || err);
    return res.status(500).json({
      error: err?.message ?? String(err),
      stack: process.env.NODE_ENV === "development" ? err?.stack : undefined,
    });
  }
};
