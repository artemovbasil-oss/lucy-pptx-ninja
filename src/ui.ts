// src/ui.ts (v0.5.1-dev) — PPT builder: smart BG + masked images + cancel UI
import PptxGenJS from "pptxgenjs";

const exportBtn = document.getElementById("export") as HTMLButtonElement;
const cancelBtn = document.getElementById("cancel") as HTMLButtonElement;
const refreshBtn = document.getElementById("refresh") as HTMLButtonElement;

const statusEl = document.getElementById("status") as HTMLDivElement;
const barEl = document.getElementById("bar") as HTMLDivElement;
const pctEl = document.getElementById("pct") as HTMLDivElement;
const progTextEl = document.getElementById("progText") as HTMLDivElement;
const tinyHintEl = document.getElementById("tinyHint") as HTMLDivElement;

const listEl = document.getElementById("list") as HTMLDivElement;
const slidesCardEl = document.getElementById("slidesCard") as HTMLDivElement;
const footerEl = document.getElementById("footer") as HTMLDivElement;

const ctaTextEl = exportBtn.querySelector(".ctaText") as HTMLSpanElement | null;

function pxToIn(px: number) { return px / 96; }
function clamp(n: number, a: number, b: number) { return Math.max(a, Math.min(b, n)); }

// Text size calibration
const FONT_SCALE = 0.705;
function pxToPt(px: number) { return px * FONT_SCALE; }

// Text box fixes
const TEXT_BOX_W_PAD_PX = 10;
const TEXT_BOX_H_PAD_PX = 2;
const TEXT_NUDGE_X_PX = -2;
const TEXT_HEIGHT_PAD_PX = 4;

// Baseline offsets
function getTextNudgeYPx(fontSizePx: number): number {
  if (fontSizePx >= 28) return 1;
  if (fontSizePx >= 16) return 2;
  return 1;
}

// Shape radius calibration
const RADIUS_SCALE = 1.10;

function opacityToTransparencyPct(opacity01: number | undefined) {
  const o = typeof opacity01 === "number" ? clamp(opacity01, 0, 1) : 1;
  return Math.round((1 - o) * 100);
}

function figmaRadiusPxToRectRadiusRatio(radiusPx: number | undefined, wPx: number, hPx: number) {
  const r0 = typeof radiusPx === "number" ? Math.max(0, radiusPx) : 0;
  if (r0 <= 0) return 0;
  const r = r0 * RADIUS_SCALE;
  const halfMin = Math.max(1, Math.min(wPx, hPx) / 2);
  return clamp(r / halfMin, 0, 1);
}

function mapFontFamily(figmaFamily: string | undefined): string {
  const f = (figmaFamily || "").trim();
  const key = f.toLowerCase();

  if (key.includes("inter")) return "Calibri";
  if (key.includes("sf pro") || key.includes("san francisco")) return "Arial";
  if (key.includes("helvetica")) return "Helvetica";
  if (key.includes("graphik")) return "Arial";
  if (key.includes("roboto")) return "Calibri";
  if (key.includes("manrope")) return "Calibri";
  if (key.includes("montserrat")) return "Calibri";
  if (key.includes("poppins")) return "Calibri";

  if (key.includes("calibri")) return "Calibri";
  if (key.includes("arial")) return "Arial";
  if (key.includes("times")) return "Times New Roman";
  if (key.includes("georgia")) return "Georgia";
  if (key.includes("verdana")) return "Verdana";
  if (key.includes("tahoma")) return "Tahoma";

  return f || "Calibri";
}

function uint8ToBase64(u8: Uint8Array): string {
  let s = "";
  const chunk = 0x8000;
  for (let i = 0; i < u8.length; i += chunk) {
    s += String.fromCharCode(...u8.subarray(i, i + chunk));
  }
  // @ts-ignore
  return btoa(s);
}

function setStatus(msg: string) { statusEl.textContent = msg; }

function setProgress(phase: string, current: number, total: number, label?: string, text?: string) {
  const t = Math.max(1, total);
  const c = clamp(current, 0, t);
  const p = Math.round((c / t) * 100);

  barEl.style.width = `${p}%`;
  pctEl.textContent = `${p}%`;
  progTextEl.textContent = label ? label : `${c}/${t}`;
  if (tinyHintEl) tinyHintEl.textContent = phase ? String(phase) : "";
  if (text) setStatus(text);
}

let isBusy = false;

function setBusy(next: boolean, ctaLabel?: string) {
  isBusy = next;

  exportBtn.disabled = next;
  refreshBtn.disabled = next;

  if (ctaTextEl) ctaTextEl.textContent = ctaLabel || (next ? "Exporting…" : "Export PPTX");
  if (next) exportBtn.classList.add("isLoading");
  else exportBtn.classList.remove("isLoading");

  if (slidesCardEl) {
    if (next) slidesCardEl.classList.add("locked");
    else slidesCardEl.classList.remove("locked");
  }

  if (footerEl) {
    if (next) footerEl.classList.add("showCancel");
    else footerEl.classList.remove("showCancel");
  }
}

type FrameInfo = { id: string; name: string; width: number; height: number };
let currentFrames: FrameInfo[] = [];

function getOrderedFrameIdsFromDOM(): string[] {
  const els = Array.from(listEl.querySelectorAll(".item")) as HTMLElement[];
  return els.map((el) => String(el.dataset.id)).filter(Boolean);
}

function renderList(frames: FrameInfo[]) {
  listEl.innerHTML = "";
  if (!frames.length) {
    const empty = document.createElement("div");
    empty.className = "muted";
    empty.textContent = "No frames selected. Select frames in Figma and click refresh.";
    listEl.appendChild(empty);
    exportBtn.disabled = true;
    return;
  }

  exportBtn.disabled = isBusy ? true : false;

  for (const f of frames) {
    const row = document.createElement("div");
    row.className = "item";
    row.draggable = true;
    row.dataset.id = f.id;

    const handle = document.createElement("div");
    handle.className = "handle";
    handle.textContent = "⋮⋮";

    const center = document.createElement("div");
    const name = document.createElement("div");
    name.className = "name";
    name.textContent = f.name;

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = `${Math.round(f.width)}×${Math.round(f.height)} px`;

    center.appendChild(name);
    center.appendChild(meta);

    row.appendChild(handle);
    row.appendChild(center);
    row.appendChild(document.createElement("div"));

    row.addEventListener("dragstart", (ev) => {
      if (isBusy) return;
      row.classList.add("dragging");
      ev.dataTransfer?.setData("text/plain", f.id);
      ev.dataTransfer?.setDragImage(row, 12, 12);
    });

    row.addEventListener("dragend", () => {
      row.classList.remove("dragging");
      Array.from(listEl.querySelectorAll(".item")).forEach((el) => el.classList.remove("over"));
      const ids = getOrderedFrameIdsFromDOM();
      currentFrames = ids.map((id) => currentFrames.find((x) => x.id === id)).filter(Boolean) as FrameInfo[];
    });

    row.addEventListener("dragover", (e) => {
      if (isBusy) return;
      e.preventDefault();
      row.classList.add("over");
    });

    row.addEventListener("dragleave", () => row.classList.remove("over"));

    row.addEventListener("drop", (e) => {
      if (isBusy) return;
      e.preventDefault();
      row.classList.remove("over");
      const draggedId = e.dataTransfer?.getData("text/plain");
      if (!draggedId) return;

      const draggedEl = listEl.querySelector(`.item[data-id="${draggedId}"]`) as HTMLElement | null;
      if (!draggedEl) return;
      if (draggedEl === row) return;

      listEl.insertBefore(draggedEl, row);
    });

    listEl.appendChild(row);
  }
}

refreshBtn.onclick = () => {
  if (isBusy) return;
  parent.postMessage({ pluginMessage: { type: "REQUEST_SELECTION" } }, "*");
};

exportBtn.onclick = () => {
  if (isBusy) return;

  const ids = getOrderedFrameIdsFromDOM();
  if (!ids.length) {
    setStatus("No frames selected.");
    return;
  }

  setBusy(true, "Exporting…");
  setProgress("prepare", 0, 1, "Starting…", "Preparing export…");
  parent.postMessage({ pluginMessage: { type: "EXPORT_PPTX_ORDERED", frameIds: ids } }, "*");
};

cancelBtn.onclick = () => {
  if (!isBusy) return;
  setProgress("cancel", 0, 1, "Cancelling…", "Stopping export…");
  parent.postMessage({ pluginMessage: { type: "CANCEL_EXPORT" } }, "*");
};

// Ask selection on open
parent.postMessage({ pluginMessage: { type: "REQUEST_SELECTION" } }, "*");

// ---- Batch PPTX builder ----
type ExportSlide = {
  name: string;
  width: number;
  height: number;
  scale: number;
  bgPngBytes: number[];
  bgShape?: { fill: string; opacity: number } | null;
  items: Array<any>;
};

function buildTransformForSlide(targetWpx: number, targetHpx: number, srcWpx: number, srcHpx: number) {
  const s = Math.min(targetWpx / srcWpx, targetHpx / srcHpx);
  const outW = srcWpx * s;
  const outH = srcHpx * s;
  const ox = (targetWpx - outW) / 2;
  const oy = (targetHpx - outH) / 2;
  return { s, ox, oy, outW, outH };
}

async function buildPptxFromSlides(filename: string, slides: ExportSlide[]) {
  const targetWpx = Math.max(...slides.map((s) => s.width));
  const targetHpx = Math.max(...slides.map((s) => s.height));

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "FIGMA_BATCH", width: pxToIn(targetWpx), height: pxToIn(targetHpx) });
  pptx.layout = "FIGMA_BATCH";

  for (let si = 0; si < slides.length; si++) {
    const sd = slides[si];
    setProgress("pptx", si, slides.length, `Building slide ${si + 1}/${slides.length}`, sd.name);

    const trf = buildTransformForSlide(targetWpx, targetHpx, sd.width, sd.height);
    const slide = pptx.addSlide();

    // 1) Smart background (shape)
    if (sd.bgShape && sd.bgShape.fill) {
      const tPct = opacityToTransparencyPct(sd.bgShape.opacity ?? 1);
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: pxToIn(targetWpx),
        h: pxToIn(targetHpx),
        fill: { color: sd.bgShape.fill, transparency: tPct },
        line: { color: sd.bgShape.fill, transparency: 100, width: 0 }
      });
    }

    // 2) Background PNG fallback
    if (sd.bgPngBytes && sd.bgPngBytes.length > 0) {
      const bgBytes = new Uint8Array(sd.bgPngBytes);
      const bgB64 = uint8ToBase64(bgBytes);

      slide.addImage({
        data: "data:image/png;base64," + bgB64,
        x: pxToIn(trf.ox),
        y: pxToIn(trf.oy),
        w: pxToIn(trf.outW),
        h: pxToIn(trf.outH)
      });
    }

    const items = (sd.items || []).slice().sort((a, b) => (a.z ?? 0) - (b.z ?? 0));

    for (const it of items) {
      const sx = (v: number) => trf.ox + v * trf.s;
      const sy = (v: number) => trf.oy + v * trf.s;
      const sw = (v: number) => v * trf.s;
      const sh = (v: number) => v * trf.s;

      if (it.kind === "maskedImage") {
        const bytes = new Uint8Array(it.pngBytes);
        const b64 = uint8ToBase64(bytes);

        const crop = it.crop || { x: 0, y: 0, w: 1, h: 1 };
        slide.addImage({
          data: "data:image/png;base64," + b64,
          x: pxToIn(sx(it.x)),
          y: pxToIn(sy(it.y)),
          w: pxToIn(sw(it.w)),
          h: pxToIn(sh(it.h)),
          crop: {
            x: clamp(Number(crop.x || 0), 0, 1),
            y: clamp(Number(crop.y || 0), 0, 1),
            w: clamp(Number(crop.w || 1), 0, 1),
            h: clamp(Number(crop.h || 1), 0, 1)
          }
        });
        continue;
      }

      if (it.kind === "raster") {
        const bytes = new Uint8Array(it.pngBytes);
        const b64 = uint8ToBase64(bytes);
        slide.addImage({
          data: "data:image/png;base64," + b64,
          x: pxToIn(sx(it.x)),
          y: pxToIn(sy(it.y)),
          w: pxToIn(sw(it.w)),
          h: pxToIn(sh(it.h))
        });
        continue;
      }

      if (it.kind === "shape") {
        const x = it.x ?? 0;
        const y = it.y ?? 0;
        const w = it.w ?? 10;
        const h = it.h ?? 10;

        const opacity = typeof it.opacity === "number" ? it.opacity : 1;
        const tPct = opacityToTransparencyPct(opacity);

        const fillProps = it.fill ? { color: it.fill, transparency: tPct } : undefined;
        const lineProps = it.stroke ? { color: it.stroke.color, width: pxToIn(sw(it.stroke.width)), transparency: tPct } : undefined;

        if (it.shape === "rect") {
          const radiusPx = typeof it.radius === "number" ? it.radius : 0;
          const rr = figmaRadiusPxToRectRadiusRatio(radiusPx * trf.s, w * trf.s, h * trf.s);
          slide.addShape(pptx.ShapeType.roundRect, {
            x: pxToIn(sx(x)),
            y: pxToIn(sy(y)),
            w: pxToIn(sw(w)),
            h: pxToIn(sh(h)),
            fill: fillProps,
            line: lineProps,
            rectRadius: rr
          });
        } else if (it.shape === "ellipse") {
          slide.addShape(pptx.ShapeType.ellipse, {
            x: pxToIn(sx(x)),
            y: pxToIn(sy(y)),
            w: pxToIn(sw(w)),
            h: pxToIn(sh(h)),
            fill: fillProps,
            line: lineProps
          });
        } else if (it.shape === "line") {
          slide.addShape(pptx.ShapeType.line, {
            x: pxToIn(sx(x)),
            y: pxToIn(sy(y)),
            w: pxToIn(sw(it.w)),
            h: pxToIn(sh(it.h)),
            line: lineProps ?? { color: it.stroke.color, width: pxToIn(sw(it.stroke.width)), transparency: tPct }
          });
        }
        continue;
      }

      if (it.kind === "text") {
        if (!it.text || String(it.text).length === 0) continue;

        const baseFsPx = Number(it.fontSize || 14);
        const effFsPx = baseFsPx * trf.s;

        const xNudge = -2 * trf.s;
        const yNudge = getTextNudgeYPx(effFsPx) * trf.s;

        const xPx = sx((it.x ?? 0) + xNudge);
        const yPx = sy((it.y ?? 0) + yNudge);

        const wPad = (10 + 2) * trf.s;
        const hPad = 2 * trf.s;

        const wPx = sw((it.w ?? 10) + wPad);
        const hPx = sh((it.h ?? 10) + 4 + hPad);

        const opacity = typeof it.opacity === "number" ? it.opacity : 1;
        const tPct = opacityToTransparencyPct(opacity);

        const lhPx = typeof it.lineHeightPx === "number" ? it.lineHeightPx * trf.s : null;
        const lineSpacingPt = lhPx ? Math.max(1, Math.round(pxToPt(lhPx))) : undefined;

        slide.addText(String(it.text), {
          x: pxToIn(xPx),
          y: pxToIn(yPx),
          w: pxToIn(wPx),
          h: pxToIn(hPx),
          margin: 0,
          inset: 0,
          fontFace: mapFontFamily(it.fontFamily),
          fontSize: Math.max(1, Math.round(pxToPt(effFsPx))),
          bold: !!it.bold,
          italic: !!it.italic,
          color: it.color || "000000",
          align: it.align || "left",
          valign: "top",
          transparency: tPct,
          ...(lineSpacingPt ? { lineSpacing: lineSpacingPt } : {})
        });
      }
    }
  }

  setProgress("pptx", slides.length, slides.length, "Writing file…", "Finalizing PPTX…");
  const arrayBuffer = await pptx.write("arraybuffer");
  const outBytes = new Uint8Array(arrayBuffer as ArrayBuffer);

  const blob = new Blob([outBytes], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "Lucy_batch.pptx";
  a.click();
  URL.revokeObjectURL(url);

  setProgress("done", 1, 1, `Done — ${slides.length} slides`, "Export complete ✅");
}

window.onmessage = async (event) => {
  const msg = event.data?.pluginMessage;
  if (!msg) return;

  if (msg.type === "STATUS") { setStatus(msg.text); return; }

  if (msg.type === "ERROR") {
    setStatus("Error:\n" + msg.text);
    setBusy(false, "Export PPTX");
    return;
  }

  if (msg.type === "SELECTION_FRAMES") {
    currentFrames = msg.frames || [];
    renderList(currentFrames);
    setStatus(currentFrames.length ? `Selected frames: ${currentFrames.length}` : "Select one or more frames.");
    setProgress("idle", 0, 1, "Idle");
    setBusy(false, "Export PPTX");
    return;
  }

  if (msg.type === "PROGRESS") {
    setProgress(msg.phase || "export", msg.current || 0, msg.total || 1, msg.label, msg.text);
    return;
  }

  if (msg.type === "CANCELLED") {
    setProgress("cancelled", 0, 1, "Cancelled", "Export cancelled.");
    setBusy(false, "Export PPTX");
    return;
  }

  if (msg.type === "BATCH_BG_AND_ITEMS_V051") {
    const slides: ExportSlide[] = msg.slides || [];
    if (!slides.length) {
      setStatus("Error: empty batch.");
      setBusy(false, "Export PPTX");
      return;
    }
    await buildPptxFromSlides(msg.filename ?? "Lucy_batch.pptx", slides);
    setBusy(false, "Export PPTX");
  }
};