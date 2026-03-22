const express = require("express");
const cors = require("cors");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const crypto = require("node:crypto");
const { buildPptxBuffer } = require("./pptx-builder");

const PORT = Number(process.env.PORT || 8787);
const JOB_ROOT = process.env.LUCY_JOB_ROOT || path.join(os.tmpdir(), "lucy-pptx-ninja-jobs");
const MAX_JSON_BODY = process.env.LUCY_MAX_JSON_BODY || "200mb";

const app = express();
app.use(cors());
app.use(express.json({ limit: MAX_JSON_BODY }));

function getJobDir(jobId) {
  return path.join(JOB_ROOT, jobId);
}

function getMetaPath(jobId) {
  return path.join(getJobDir(jobId), "meta.json");
}

function getSlidePath(jobId, index) {
  return path.join(getJobDir(jobId), "slides", `${String(index).padStart(5, "0")}.json`);
}

function getOutputPath(jobId) {
  return path.join(getJobDir(jobId), "output.pptx");
}

async function ensureJobRoot() {
  await fsp.mkdir(JOB_ROOT, { recursive: true });
}

async function readJobMeta(jobId) {
  const metaPath = getMetaPath(jobId);
  const raw = await fsp.readFile(metaPath, "utf8");
  return JSON.parse(raw);
}

async function writeJobMeta(jobId, patch) {
  const current = await readJobMeta(jobId);
  const next = { ...current, ...patch, updatedAt: new Date().toISOString() };
  await fsp.writeFile(getMetaPath(jobId), JSON.stringify(next, null, 2), "utf8");
  return next;
}

function getBaseUrl(req) {
  const envBase = process.env.PUBLIC_BASE_URL || process.env.LUCY_PUBLIC_BASE_URL;
  if (envBase) return envBase.replace(/\/$/, "");
  return `${req.protocol}://${req.get("host")}`;
}

async function buildJob(jobId) {
  try {
    await writeJobMeta(jobId, { status: "building", error: null });
    const meta = await readJobMeta(jobId);
    const slidesDir = path.join(getJobDir(jobId), "slides");
    const slideFiles = (await fsp.readdir(slidesDir))
      .filter((name) => name.endsWith(".json"))
      .sort((a, b) => a.localeCompare(b));

    const slides = [];
    for (const fileName of slideFiles) {
      const raw = await fsp.readFile(path.join(slidesDir, fileName), "utf8");
      slides.push(JSON.parse(raw));
    }

    if (!slides.length) {
      throw new Error("No slides uploaded for this job.");
    }

    const buffer = await buildPptxBuffer(slides);
    await fsp.writeFile(getOutputPath(jobId), buffer);
    await writeJobMeta(jobId, { status: "completed", completedAt: new Date().toISOString() });
  } catch (error) {
    await writeJobMeta(jobId, {
      status: "failed",
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

app.get("/health", async (_req, res) => {
  await ensureJobRoot();
  res.json({ ok: true });
});

app.post("/api/jobs", async (req, res) => {
  await ensureJobRoot();
  const jobId = crypto.randomUUID();
  const filename = typeof req.body?.filename === "string" && req.body.filename.trim()
    ? req.body.filename.trim()
    : "Lucy_batch.pptx";

  const meta = {
    id: jobId,
    filename,
    status: "uploading",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    slideCount: 0,
    error: null
  };

  await fsp.mkdir(path.join(getJobDir(jobId), "slides"), { recursive: true });
  await fsp.writeFile(getMetaPath(jobId), JSON.stringify(meta, null, 2), "utf8");
  res.status(201).json({ jobId, status: meta.status });
});

app.put("/api/jobs/:jobId/slides/:index", async (req, res) => {
  const { jobId, index } = req.params;
  const slideIndex = Number(index);
  if (!Number.isInteger(slideIndex) || slideIndex < 0) {
    res.status(400).json({ error: "Slide index must be a non-negative integer." });
    return;
  }

  if (!fs.existsSync(getMetaPath(jobId))) {
    res.status(404).json({ error: "Job not found." });
    return;
  }

  await fsp.writeFile(getSlidePath(jobId, slideIndex), JSON.stringify(req.body), "utf8");
  const slidesDir = path.join(getJobDir(jobId), "slides");
  const slideCount = (await fsp.readdir(slidesDir)).filter((name) => name.endsWith(".json")).length;
  const meta = await writeJobMeta(jobId, { slideCount });
  res.json({ ok: true, status: meta.status, slideCount });
});

app.post("/api/jobs/:jobId/finalize", async (req, res) => {
  const { jobId } = req.params;
  if (!fs.existsSync(getMetaPath(jobId))) {
    res.status(404).json({ error: "Job not found." });
    return;
  }

  const meta = await writeJobMeta(jobId, { status: "queued" });
  void buildJob(jobId);
  res.status(202).json({ ok: true, status: meta.status });
});

app.get("/api/jobs/:jobId/status", async (req, res) => {
  const { jobId } = req.params;
  if (!fs.existsSync(getMetaPath(jobId))) {
    res.status(404).json({ error: "Job not found." });
    return;
  }

  const meta = await readJobMeta(jobId);
  const baseUrl = getBaseUrl(req);
  res.json({
    ...meta,
    downloadUrl: meta.status === "completed" ? `${baseUrl}/api/jobs/${jobId}/download` : null
  });
});

app.get("/api/jobs/:jobId/download", async (req, res) => {
  const { jobId } = req.params;
  if (!fs.existsSync(getMetaPath(jobId))) {
    res.status(404).json({ error: "Job not found." });
    return;
  }

  const meta = await readJobMeta(jobId);
  const outputPath = getOutputPath(jobId);
  if (meta.status !== "completed" || !fs.existsSync(outputPath)) {
    res.status(409).json({ error: "Job is not ready yet." });
    return;
  }

  res.download(outputPath, meta.filename || "Lucy_batch.pptx");
});

app.listen(PORT, async () => {
  await ensureJobRoot();
  console.log(`Lucy PPTX backend listening on :${PORT}`);
});
