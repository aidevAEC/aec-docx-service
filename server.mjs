import express from "express";
import fetch from "node-fetch";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const app = express();
app.use(express.json({ limit: "10mb" }));

app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/render", async (req, res) => {
  try {
    const { templateUrl, data } = req.body || {};
    if (!templateUrl || !data) {
      return res.status(400).json({ error: "templateUrl and data required" });
    }

    const tplRes = await fetch(templateUrl);
    if (!tplRes.ok) throw new Error(`Fetch template failed: ${tplRes.status}`);
    const arrayBuf = await tplRes.arrayBuffer();
    const tplBuf = Buffer.from(arrayBuf);

    const zip = new PizZip(tplBuf);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    doc.setData(data);
    doc.render();

    const out = doc.getZip().generate({ type: "nodebuffer" });
    const filename = (data.FileName || "AEC Letter Draft.docx").replace(/[\\/:*?\"<>|]+/g, "_");
    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${filename}"`
    });
    res.send(out);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message || "Render failed" });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("AEC DOCX service running on port " + PORT));
