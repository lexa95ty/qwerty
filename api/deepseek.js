module.exports = async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed" });
      return;
    }

    const apiKey = process.env.DEEPSEEK_API_KEY;
    if (!apiKey) {
      res.status(500).json({ error: "Missing DEEPSEEK_API_KEY in environment" });
      return;
    }

    const payload = req.body || {};
    if (!payload.model || !payload.messages) {
      res.status(400).json({ error: "Invalid payload (model/messages required)" });
      return;
    }

    const r = await fetch("https://api.deepseek.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: payload.model,
        messages: payload.messages,
        temperature: payload.temperature ?? 0.7,
        max_tokens: payload.max_tokens ?? 800
      })
    });

    const data = await r.json().catch(() => ({}));
    if (!r.ok) {
      res.status(r.status).json({
        error: data?.error?.message || data?.error || data?.message || "DeepSeek error",
        details: data
      });
      return;
    }

    res.status(200).json(data);
  } catch (e) {
    res.status(500).json({ error: e.message || "Server error" });
  }
}
