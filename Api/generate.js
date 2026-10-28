const Anthropic = require("@anthropic-ai/sdk");
const PptxGenJS = require("pptxgenjs");
const { Document, Packer, Paragraph, HeadingLevel } = require("docx");
const mammoth = require("mammoth");

const client = new Anthropic(); // bruker ANTHROPIC_API_KEY fra miljøvariabler
const MODEL = "claude-sonnet-5-5";

const tool = {
  name: "create_document",
  description: "Lag innholdet til et dokument eller en presentasjon.",
  input_schema: {
    type: "object",
    properties: {
      title: { type: "string" },
      sections: {
        type: "array",
        description: "Hver seksjon blir ett lysbilde (pptx) eller en overskrift med innhold (docx).",
        items: {
          type: "object",
          properties: {
            heading: { type: "string" },
            paragraphs: { type: "array", items: { type: "string" } },
            bullets: { type: "array", items: { type: "string" } }
          },
          required: ["heading"]
        }
      }
    },
    required: ["title", "sections"]
  }
};

async function buildContent(prompt, files, format) {
  const content = [];
  for (const f of files) {
    if (f.type.startsWith("image/")) {
      content.push({ type: "image", source: { type: "base64", media_type: f.type, data: f.data } });
    } else if (f.type === "application/pdf") {
      content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: f.data } });
    } else if (f.name.toLowerCase().endsWith(".docx")) {
      const r = await mammoth.extractRawText({ buffer: Buffer.from(f.data, "base64") });
      content.push({ type: "text", text: `Innhold fra filen ${f.name}:\n${r.value}` });
    } else {
      content.push({ type: "text", text: `Innhold fra filen ${f.name}:\n${Buffer.from(f.data, "base64").toString("utf8")}` });
    }
  }
  content.push({
    type: "text",
    text: `Oppgave: ${prompt || "Lag et godt dokument av materialet over."}\n\nFormat: ${format === "pptx" ? "PowerPoint-presentasjon (korte, tydelige punkter per lysbilde)" : "Word-dokument"}. Skriv på samme språk som oppgaven.`
  });
  return content;
}

async function makePptx(d) {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  const first = pres.addSlide();
  first.background = { color: "1F2A44" };
  first.addText(d.title, { x: 0.7, y: 1.8, w: 8.6, h: 1.5, fontSize: 38, bold: true, color: "FFFFFF" });
  for (const s of d.sections) {
    const slide = pres.addSlide();
    slide.addText(s.heading, { x: 0.5, y: 0.3, w: 9, h: 0.9, fontSize: 28, bold: true, color: "1F2A44" });
    const items = [
      ...(s.paragraphs || []).map(t => ({ text: t, options: { breakLine: true, paraSpaceAfter: 8 } })),
      ...(s.bullets || []).map(t => ({ text: t, options: { bullet: true, breakLine: true, paraSpaceAfter: 6 } }))
    ];
    if (items.length) slide.addText(items, { x: 0.6, y: 1.3, w: 8.8, h: 3.9, fontSize: 18, valign: "top" });
  }
  return { base64: await pres.write({ outputType: "base64" }), ext: "pptx" };
}

async function makeDocx(d) {
  const children = [new Paragraph({ text: d.title, heading: HeadingLevel.TITLE })];
  for (const s of d.sections) {
    children.push(new Paragraph({ text: s.heading, heading: HeadingLevel.HEADING_1 }));
    (s.paragraphs || []).forEach(t => children.push(new Paragraph({ text: t, spacing: { after: 160 } })));
    (s.bullets || []).forEach(t => children.push(new Paragraph({ text: t, bullet: { level: 0 } })));
  }
  const doc = new Document({ sections: [{ children }] });
  return { base64: await Packer.toBase64String(doc), ext: "docx" };
}

module.exports = async (req, res) => {
  if (req.method !== "POST") return res.status(405).json({ error: "Kun POST" });
  if (process.env.APP_PASSWORD && req.headers["x-password"] !== process.env.APP_PASSWORD) {
    return res.status(401).json({ error: "Feil passord" });
  }
  try {
    const { prompt = "", format = "pptx", files = [] } = req.body || {};
    const response = await client.messages.create({
      model: MODEL,
      max_tokens: 8000,
      tools: [tool],
      tool_choice: { type: "tool", name: "create_document" },
      messages: [{ role: "user", content: await buildContent(prompt, files, format) }]
    });
    const block = response.content.find(b => b.type === "tool_use");
    if (!block) throw new Error("Fikk ikke noe innhold fra AI-en");
    const out = format === "pptx" ? await makePptx(block.input) : await makeDocx(block.input);
    const safe = block.input.title.replace(/[^\p{L}\p{N} _-]/gu, "").trim().slice(0, 60) || "dokument";
    res.status(200).json({ filename: `${safe}.${out.ext}`, base64: out.base64 });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message || "Noe gikk galt" });
  }
};
