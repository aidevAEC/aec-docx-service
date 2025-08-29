// server.mjs
// Tiny DOCX render service for n8n + docxtemplater

import express from "express";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const app = express();

// accept JSON bodies
app.use(express.json({ limit: "10mb" }));

// (optional but handy) allow browser / n8n calls from anywhere
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "content-type, x-api-key");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

// serve static templates from /templates so you can use
// https://YOUR-APP.onrender.com/templates/AEC_Proposal_Template_v1.docx
app.use("/templates", express.static("templates", { immutable: true, maxAge: "1y" }));

const API_KEY = process.env.API_KEY || "dev-key";

// health check
app.get("/health", (_req, res) => res.json({ ok: true }));

// main render endpoint
app.post("/render", async (req, res) => {
  try {
    // simple API key check
    if (req.headers["x-api-key"] !== API_KEY) {
      return res.status(401).json({ error: "unauthorized" });
    }

    const { templateUrl, data, fileName } = req.body || {};
    if (!templateUrl || !data) {
      return res.status(400).json({ error: "templateUrl and data are required" });
    }

    // fetch the .docx template
    const r = await fetch(templateUrl);
    if (!r.ok) {
      return res.status(400).json({
        error: "failed_to_fetch_template",
        status: r.status,
        url: templateUrl
      });
    }

    // load the .docx into docxtemplater
    let zip;
    try {
      const ab = await r.arrayBuffer();
      zip = new PizZip(Buffer.from(ab));
    } catch (e) {
      return res.status(400).json({ error: "invalid_docx_zip", message: e.message });
    }

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // set data and render
    doc.setData(data);

    try {
      doc.render();
    } catch (e) {
      // return helpful tag errors to the caller
      const details =
        e.properties?.errors?.map((er) => ({
          id: er.id,
          tag: er.properties?.tag,
          explanation: er.properties?.explanation,
        })) || [];
      return res.status(400).json({
        error: "template_render_error",
        message: e.message,
        details,
      });
    }

    // send the filled docx back
    const buf = doc.getZip().generate({ type: "nodebuffer" });
    const outName = ((fileName || data?.FileName || "Filled") + ".docx").replace(/"/g, "");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    return res.send(buf);
  } catch (err) {
    return res.status(500).json({ error: "server_error", message: err.message });
  }
});

// start the server
const port = process.env.PORT || 3000;
app.listen(port, () => console.log("listening on", port));
