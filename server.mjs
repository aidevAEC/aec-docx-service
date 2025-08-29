import express from "express";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const app = express();
app.use(express.json({ limit: "8mb" }));

const API_KEY = process.env.API_KEY || "dev-key";

app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/render", async (req, res) => {
  try {
    if (req.headers["x-api-key"] !== API_KEY) {
      return res.status(401).json({ error: "unauthorized" });
    }
    const { templateUrl, data } = req.body || {};
    if (!templateUrl || !data) {
      return res.status(400).json({ error: "templateUrl and data are required" });
    }

    const r = await fetch(templateUrl);
    if (!r.ok) return res.status(400).json({ error: "failed_to_fetch_template", status: r.status });

    const ab = await r.arrayBuffer();
    const zip = new PizZip(Buffer.from(ab));

    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    doc.setData(data);

    try { doc.render(); }
    catch (e) {
      return res.status(400).json({
        error: "template_render_error",
        message: e.message,
        details: e.properties?.errors?.map(er => ({
          id: er.id, tag: er.properties?.tag, explanation: er.properties?.explanation
        })) || []
      });
    }

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition",'attachment; filename="Filled.docx"');
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: "server_error", message: err.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log("listening on", port));
