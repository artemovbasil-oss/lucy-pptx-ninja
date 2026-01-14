// src/ui.ts (v0.5.4-dev) — Font handling improvements + bgShape support + V051 handler
import PptxGenJS from "pptxgenjs";

// Keep these in sync with your release notes
const UI_VERSION = "v0.6";
const UI_HIGHLIGHT = "Batch export + PDF";

const exportBtn = document.getElementById("export") as HTMLButtonElement;
const cancelBtn = document.getElementById("cancel") as HTMLButtonElement | null;
const statusEl = document.getElementById("status") as HTMLDivElement;
const barEl = document.getElementById("bar") as HTMLDivElement;
// main progress percent is intentionally hidden (we show a bar + label instead)
const pctEl = document.getElementById("pct") as HTMLDivElement | null;
const progTextEl = document.getElementById("progText") as HTMLDivElement;
const stateDotEl = document.getElementById("stateDot") as HTMLDivElement;

const listEl = document.getElementById("list") as HTMLDivElement;
const slidesCardEl = document.getElementById("slidesCard") as HTMLDivElement;
const footerEl = document.getElementById("footer") as HTMLDivElement;

const formatSelect = document.getElementById("formatSelect") as HTMLDivElement | null;
const qualitySelect = document.getElementById("qualitySelect") as HTMLDivElement | null;

const versionEl = document.getElementById("version") as HTMLDivElement | null;

// Busy overlay elements
const busyOverlayEl = document.getElementById("busyOverlay") as HTMLDivElement | null;
// Overlay uses a ring indicator (no numeric percent)
const ringEl = document.querySelector(".ring") as HTMLDivElement | null;
const overlayHintEl = document.getElementById("overlayHint") as HTMLDivElement | null;
const overlayCancelBtn = document.getElementById("overlayCancel") as HTMLButtonElement | null;


const ctaTextEl = exportBtn.querySelector(".ctaText") as HTMLSpanElement | null;

if (versionEl) versionEl.textContent = `${UI_VERSION} · ${UI_HIGHLIGHT}`;

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

/**
 * FONT STRATEGY
 * - Default: keep original font family from Figma (best match on your machine).
 * - Optional: client-safe mapping (Calibri/Arial/etc.) for environments without custom fonts.
 */
const CLIENT_SAFE_FONTS = false; // <- switch to true if you want "safe for any client" output

function normalizeFontName(name: string | undefined): string {
  return (name || "").trim().replace(/\s+/g, " ");
}

// Conservative mapping for client-safe mode (feel free to tune)
function mapToClientSafeFont(figmaFamily: string): string {
  const f = figmaFamily.toLowerCase();

  if (f.includes("sf pro") || f.includes("san francisco")) return "Arial";
  if (f.includes("helvetica")) return "Helvetica";
  if (f.includes("inter")) return "Calibri"; // common office default
  if (f.includes("graphik")) return "Arial";
  if (f.includes("roboto")) return "Calibri";
  if (f.includes("manrope")) return "Calibri";
  if (f.includes("montserrat")) return "Calibri";
  if (f.includes("poppins")) return "Calibri";

  if (f.includes("calibri")) return "Calibri";
  if (f.includes("arial")) return "Arial";
  if (f.includes("times")) return "Times New Roman";
  if (f.includes("georgia")) return "Georgia";
  if (f.includes("verdana")) return "Verdana";
  if (f.includes("tahoma")) return "Tahoma";

  return "Calibri";
}

// Default mode: preserve original font name; fallback only if empty.
function mapFontFamily(figmaFamily: string | undefined): string {
  const clean = normalizeFontName(figmaFamily);
  if (!clean) return "Calibri";
  return CLIENT_SAFE_FONTS ? mapToClientSafeFont(clean) : clean;
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

type UiState = "idle" | "processing" | "success" | "error";
function setState(state: UiState) {
  if (!stateDotEl) return;
  stateDotEl.classList.remove("idle", "processing", "success", "error");
  stateDotEl.classList.add(state);
}

function setProgress(phase: string, current: number, total: number, label?: string, text?: string) {
  const t = Math.max(1, total);
  const c = clamp(current, 0, t);
  const p = Math.round((c / t) * 100);

  barEl.style.width = `${p}%`;
  if (pctEl) pctEl.textContent = "";
  if (isBusy && ringEl) ringEl.style.setProperty("--p", String(p));
  if (isBusy && overlayHintEl && phase) overlayHintEl.textContent = String(phase);
  progTextEl.textContent = label ? label : `${c}/${t}`;
  if (text) setStatus(text);
}

let isBusy = false;
let uiCancelRequested = false;

function showBusyOverlay(show: boolean) {
  if (!busyOverlayEl) return;
  if (show) {
    busyOverlayEl.classList.add("show");
    busyOverlayEl.setAttribute("aria-hidden", "false");
    if (ringEl) ringEl.style.setProperty("--p", "0");
  } else {
    busyOverlayEl.classList.remove("show");
    busyOverlayEl.setAttribute("aria-hidden", "true");
  }
}

function setBusy(next: boolean, ctaLabel?: string) {
  isBusy = next;
  // show overlay only during export
  showBusyOverlay(next);

  setState(next ? "processing" : "idle");

  exportBtn.disabled = next;

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

type FrameInfo = { id: string; name: string; width: number; height: number; thumbBytes?: number[] | null };
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
    empty.textContent = "No frames selected. Select one or more frames (or a section) in Figma.";
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

    const thumbWrap = document.createElement("div");
    thumbWrap.className = "thumb";
    if (f.thumbBytes && f.thumbBytes.length > 0) {
      const img = document.createElement("img");
      img.alt = f.name;
      img.src = `data:image/png;base64,${uint8ToBase64(new Uint8Array(f.thumbBytes))}`;
      thumbWrap.appendChild(img);
    } else {
      thumbWrap.classList.add("thumbEmpty");
    }

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
    row.appendChild(thumbWrap);
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

exportBtn.onclick = () => {
  if (isBusy) return;

  const ids = getOrderedFrameIdsFromDOM();
  if (!ids.length) {
    setStatus("No frames selected.");
    return;
  }

  uiCancelRequested = false;
  setBusy(true, "Exporting…");
  setProgress("prepare", 0, 1, "Starting…", "Preparing export…");
  parent.postMessage({
    pluginMessage: {
      type: "EXPORT_PPTX_ORDERED",
      frameIds: ids,
      format: getSegmentedValue(formatSelect, "pptx"),
      quality: getSegmentedValue(qualitySelect, "best")
    }
  }, "*");
};

cancelBtn?.addEventListener('click', () => {
  if (!isBusy) return;
  uiCancelRequested = true;
  setProgress("cancel", 0, 1, "Cancelling…", "Stopping export…");
  parent.postMessage({ pluginMessage: { type: "CANCEL_EXPORT" } }, "*");
});

// Overlay Cancel (only in overlay)
overlayCancelBtn?.addEventListener('click', () => {
  if (!isBusy) return;
  uiCancelRequested = true;
  setProgress('cancel', 0, 1, 'Cancelling…', 'Stopping export…');
  parent.postMessage({ pluginMessage: { type: 'CANCEL_EXPORT' } }, '*');
});

// Ask selection on open
parent.postMessage({ pluginMessage: { type: "REQUEST_SELECTION" } }, "*");

// ---- Batch PPTX builder ----
type ExportSlide = {
  name: string;
  width: number;
  height: number;
  scale: number;
  bgPngBytes: number[]; // empty if bgShape is used
  bgShape?: { fill: string; opacity: number } | null; // optional Smart BG
  fullPngBytes?: number[] | null;
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

function formatFontsUsed(fonts: Set<string>) {
  const arr = Array.from(fonts).filter(Boolean).sort((a, b) => a.localeCompare(b));
  if (!arr.length) return "";
  return `Fonts used: ${arr.join(", ")}${CLIENT_SAFE_FONTS ? " (client-safe mapping ON)" : ""}`;
}

async function buildPptxFromSlides(filename: string, slides: ExportSlide[]) {
  if (uiCancelRequested) throw new Error("CANCELLED_UI");
  const targetWpx = Math.max(...slides.map((s) => s.width));
  const targetHpx = Math.max(...slides.map((s) => s.height));

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "FIGMA_BATCH", width: pxToIn(targetWpx), height: pxToIn(targetHpx) });
  pptx.layout = "FIGMA_BATCH";

  const fontsUsed = new Set<string>();

  for (let si = 0; si < slides.length; si++) {
    if (uiCancelRequested) throw new Error("CANCELLED_UI");
    const sd = slides[si];
    setProgress("pptx", si, slides.length, `Building slide ${si + 1}/${slides.length}`, sd.name);

    const trf = buildTransformForSlide(targetWpx, targetHpx, sd.width, sd.height);
    const slide = pptx.addSlide();

    // Background: either PNG OR Smart BG shape
    const hasBgPng = Array.isArray(sd.bgPngBytes) && sd.bgPngBytes.length > 0;
    const hasBgShape = !!sd.bgShape && !!sd.bgShape.fill;

    if (hasBgPng) {
      const bgBytes = new Uint8Array(sd.bgPngBytes);
      const bgB64 = uint8ToBase64(bgBytes);
      slide.addImage({
        data: "data:image/png;base64," + bgB64,
        x: pxToIn(trf.ox),
        y: pxToIn(trf.oy),
        w: pxToIn(trf.outW),
        h: pxToIn(trf.outH)
      });
    } else if (hasBgShape) {
      const op = typeof sd.bgShape!.opacity === "number" ? sd.bgShape!.opacity : 1;
      const tPct = opacityToTransparencyPct(op);
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: pxToIn(targetWpx),
        h: pxToIn(targetHpx),
        fill: { color: String(sd.bgShape!.fill), transparency: tPct },
        line: { color: String(sd.bgShape!.fill), transparency: 100 }
      });
    }

    const items = (sd.items || []).slice().sort((a, b) => (a.z ?? 0) - (b.z ?? 0));

    for (const it of items) {
      if (uiCancelRequested) throw new Error("CANCELLED_UI");
      const sx = (v: number) => trf.ox + v * trf.s;
      const sy = (v: number) => trf.oy + v * trf.s;
      const sw = (v: number) => v * trf.s;
      const sh = (v: number) => v * trf.s;

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

        const xNudge = TEXT_NUDGE_X_PX * trf.s;
        const yNudge = getTextNudgeYPx(effFsPx) * trf.s;

        const xPx = sx((it.x ?? 0) + xNudge);
        const yPx = sy((it.y ?? 0) + yNudge);

        const wPad = TEXT_BOX_W_PAD_PX * trf.s;
        const hPad = TEXT_BOX_H_PAD_PX * trf.s;

        const wPx = sw((it.w ?? 10) + wPad);
        const hPx = sh((it.h ?? 10) + TEXT_HEIGHT_PAD_PX + hPad);

        const opacity = typeof it.opacity === "number" ? it.opacity : 1;
        const tPct = opacityToTransparencyPct(opacity);

        const lhPx = typeof it.lineHeightPx === "number" ? it.lineHeightPx * trf.s : null;
        const lineSpacingPt = lhPx ? Math.max(1, Math.round(pxToPt(lhPx))) : undefined;

        const fontFace = mapFontFamily(it.fontFamily);
        fontsUsed.add(fontFace);

        const rawText = String(it.text);
        const finalText = it.uppercase ? rawText.toUpperCase() : rawText;

        slide.addText(finalText, {
          x: pxToIn(xPx),
          y: pxToIn(yPx),
          w: pxToIn(wPx),
          h: pxToIn(hPx),
          margin: 0,
          inset: 0,
          fontFace,
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
  if (uiCancelRequested) throw new Error("CANCELLED_UI");
  const arrayBuffer = await pptx.write("arraybuffer");
  const outBytes = new Uint8Array(arrayBuffer as ArrayBuffer);

  const blob = new Blob([outBytes], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "Lucy_batch.pptx";
  a.click();
  URL.revokeObjectURL(url);

  // Show fonts list as a useful hint for matching on client machines
  const fontsHint = formatFontsUsed(fontsUsed);
  if (tinyHintEl && fontsHint) tinyHintEl.textContent = fontsHint;

  setProgress("done", 1, 1, `Done — ${slides.length} slides`, "Export complete ✅");
  setState("success");
}

function hexToRgba(hex: string, alpha: number): string {
  const clean = hex.replace("#", "");
  const r = parseInt(clean.slice(0, 2), 16) || 0;
  const g = parseInt(clean.slice(2, 4), 16) || 0;
  const b = parseInt(clean.slice(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function drawRoundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
  const radius = Math.max(0, Math.min(r, Math.min(w, h) / 2));
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.arcTo(x + w, y, x + w, y + h, radius);
  ctx.arcTo(x + w, y + h, x, y + h, radius);
  ctx.arcTo(x, y + h, x, y, radius);
  ctx.arcTo(x, y, x + w, y, radius);
  ctx.closePath();
}

function createCanvas(width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width));
  canvas.height = Math.max(1, Math.round(height));
  return canvas;
}

async function loadImageFromBytes(bytes: number[]): Promise<HTMLImageElement> {
  const blob = new Blob([new Uint8Array(bytes)], { type: "image/png" });
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error("Image load failed"));
    });
    return img;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function jpgQualityForMode(mode: string): number {
  if (mode === "low") return 0.45;
  if (mode === "medium") return 0.7;
  return 0.9;
}

function rasterScaleForMode(mode: string): number {
  if (mode === "low") return 0.6;
  if (mode === "medium") return 0.8;
  return 1;
}

function buildPdfBytes(pages: { jpgBytes: Uint8Array; imgWidth: number; imgHeight: number; pageWidth: number; pageHeight: number }[]): Uint8Array {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  let offset = 0;
  const xref: number[] = [];

  function pushString(s: string) {
    const bytes = encoder.encode(s);
    chunks.push(bytes);
    offset += bytes.length;
  }

  function pushBytes(b: Uint8Array) {
    chunks.push(b);
    offset += b.length;
  }

  pushString("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");

  const objects: Array<{ id: number; content: string; stream?: Uint8Array }> = [];
  const catalogId = 1;
  const pagesId = 2;
  let nextId = 3;
  const pageRefs: number[] = [];

  for (const page of pages) {
    const imgId = nextId++;
    const contentId = nextId++;
    const pageId = nextId++;

    objects.push({
      id: imgId,
      content: `<< /Type /XObject /Subtype /Image /Width ${page.imgWidth} /Height ${page.imgHeight} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpgBytes.length} >>`,
      stream: page.jpgBytes
    });

    const contentStream = `q ${page.pageWidth} 0 0 ${page.pageHeight} 0 0 cm /Im${imgId} Do Q`;
    objects.push({
      id: contentId,
      content: `<< /Length ${contentStream.length} >>`,
      stream: encoder.encode(contentStream)
    });

    objects.push({
      id: pageId,
      content: `<< /Type /Page /Parent ${pagesId} 0 R /Resources << /XObject << /Im${imgId} ${imgId} 0 R >> >> /MediaBox [0 0 ${page.pageWidth} ${page.pageHeight}] /Contents ${contentId} 0 R >>`
    });

    pageRefs.push(pageId);
  }

  objects.push({
    id: pagesId,
    content: `<< /Type /Pages /Kids [${pageRefs.map((r) => `${r} 0 R`).join(" ")}] /Count ${pageRefs.length} >>`
  });

  objects.push({
    id: catalogId,
    content: `<< /Type /Catalog /Pages ${pagesId} 0 R >>`
  });

  objects.sort((a, b) => a.id - b.id);

  for (const obj of objects) {
    xref.push(offset);
    pushString(`${obj.id} 0 obj\n`);
    if (obj.stream) {
      pushString(obj.content);
      pushString("\nstream\n");
      pushBytes(obj.stream);
      pushString("\nendstream\nendobj\n");
    } else {
      pushString(obj.content);
      pushString("\nendobj\n");
    }
  }

  const xrefOffset = offset;
  pushString(`xref\n0 ${xref.length + 1}\n`);
  pushString("0000000000 65535 f \n");
  for (const pos of xref) {
    pushString(`${String(pos).padStart(10, "0")} 00000 n \n`);
  }
  pushString(`trailer\n<< /Size ${xref.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`);

  const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
  const out = new Uint8Array(totalLength);
  let cursor = 0;
  for (const c of chunks) {
    out.set(c, cursor);
    cursor += c.length;
  }
  return out;
}

async function buildPdfFromSlides(filename: string, slides: ExportSlide[], qualityMode: string) {
  if (uiCancelRequested) throw new Error("CANCELLED_UI");
  const targetWpx = Math.max(...slides.map((s) => s.width));
  const targetHpx = Math.max(...slides.map((s) => s.height));
  const jpgQuality = jpgQualityForMode(qualityMode);
  const pages: { jpgBytes: Uint8Array; imgWidth: number; imgHeight: number; pageWidth: number; pageHeight: number }[] = [];

  for (let si = 0; si < slides.length; si++) {
    if (uiCancelRequested) throw new Error("CANCELLED_UI");
    const sd = slides[si];
    setProgress("pdf", si, slides.length, `Building page ${si + 1}/${slides.length}`, sd.name);

    const trf = buildTransformForSlide(targetWpx, targetHpx, sd.width, sd.height);
    const rasterScale = rasterScaleForMode(qualityMode);
    const canvas = createCanvas(targetWpx * rasterScale, targetHpx * rasterScale);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas unavailable");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    if (sd.fullPngBytes && sd.fullPngBytes.length > 0) {
      const img = await loadImageFromBytes(sd.fullPngBytes);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    } else {
      if (Array.isArray(sd.bgPngBytes) && sd.bgPngBytes.length > 0) {
        const img = await loadImageFromBytes(sd.bgPngBytes);
        ctx.drawImage(
          img,
          trf.ox * rasterScale,
          trf.oy * rasterScale,
          trf.outW * rasterScale,
          trf.outH * rasterScale
        );
      } else if (sd.bgShape?.fill) {
        ctx.fillStyle = hexToRgba(sd.bgShape.fill, typeof sd.bgShape.opacity === "number" ? sd.bgShape.opacity : 1);
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }

      const items = (sd.items || []).slice().sort((a, b) => (a.z ?? 0) - (b.z ?? 0));
      for (const it of items) {
        if (uiCancelRequested) throw new Error("CANCELLED_UI");
        const sx = (v: number) => (trf.ox + v * trf.s) * rasterScale;
        const sy = (v: number) => (trf.oy + v * trf.s) * rasterScale;
        const sw = (v: number) => v * trf.s * rasterScale;
        const sh = (v: number) => v * trf.s * rasterScale;

        if (it.kind === "raster") {
          const img = await loadImageFromBytes(it.pngBytes);
          ctx.drawImage(img, sx(it.x), sy(it.y), sw(it.w), sh(it.h));
          continue;
        }

        if (it.kind === "maskedImage") {
          const img = await loadImageFromBytes(it.pngBytes);
          const srcW = img.width;
          const srcH = img.height;
          const sx0 = it.crop.x * srcW;
          const sy0 = it.crop.y * srcH;
          const sw0 = it.crop.w * srcW;
          const sh0 = it.crop.h * srcH;
          ctx.drawImage(img, sx0, sy0, sw0, sh0, sx(it.x), sy(it.y), sw(it.w), sh(it.h));
          continue;
        }

        if (it.kind === "shape") {
          const x = sx(it.x ?? 0);
          const y = sy(it.y ?? 0);
          const w = sw(it.w ?? 10);
          const h = sh(it.h ?? 10);
          const opacity = typeof it.opacity === "number" ? it.opacity : 1;

          ctx.save();
          ctx.globalAlpha = opacity;

          if (it.shape === "rect") {
            const radius = typeof it.radius === "number" ? it.radius * trf.s : 0;
            drawRoundRect(ctx, x, y, w, h, radius);
            if (it.fill) {
              ctx.fillStyle = hexToRgba(it.fill, 1);
              ctx.fill();
            }
            if (it.stroke) {
              ctx.strokeStyle = hexToRgba(it.stroke.color, 1);
              ctx.lineWidth = sw(it.stroke.width);
              ctx.stroke();
            }
          } else if (it.shape === "ellipse") {
            ctx.beginPath();
            ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
            if (it.fill) {
              ctx.fillStyle = hexToRgba(it.fill, 1);
              ctx.fill();
            }
            if (it.stroke) {
              ctx.strokeStyle = hexToRgba(it.stroke.color, 1);
              ctx.lineWidth = sw(it.stroke.width);
              ctx.stroke();
            }
          } else if (it.shape === "line") {
            ctx.beginPath();
            ctx.moveTo(x, y);
            ctx.lineTo(x + w, y + h);
            ctx.strokeStyle = it.stroke ? hexToRgba(it.stroke.color, 1) : "#000";
            ctx.lineWidth = it.stroke ? sw(it.stroke.width) : 1;
            ctx.stroke();
          }

          ctx.restore();
          continue;
        }

        if (it.kind === "text") {
          if (!it.text || String(it.text).length === 0) continue;
          const baseFsPx = Number(it.fontSize || 14);
          const effFsPx = baseFsPx * trf.s * rasterScale;
          const xPx = sx((it.x ?? 0) + TEXT_NUDGE_X_PX);
          const yPx = sy((it.y ?? 0) + getTextNudgeYPx(effFsPx));
          const wPx = sw((it.w ?? 10) + TEXT_BOX_W_PAD_PX);
          const hPx = sh((it.h ?? 10) + TEXT_HEIGHT_PAD_PX + TEXT_BOX_H_PAD_PX);
          const opacity = typeof it.opacity === "number" ? it.opacity : 1;
          const fontFace = mapFontFamily(it.fontFamily);
          const rawText = String(it.text);
          const finalText = it.uppercase ? rawText.toUpperCase() : rawText;
          const lines = finalText.split("\n");
          const lineHeight = typeof it.lineHeightPx === "number" ? it.lineHeightPx * trf.s * rasterScale : effFsPx * 1.2;

          ctx.save();
          ctx.globalAlpha = opacity;
          ctx.fillStyle = hexToRgba(it.color || "000000", 1);
          ctx.font = `${it.italic ? "italic " : ""}${it.bold ? "bold " : ""}${Math.max(1, Math.round(effFsPx))}px ${fontFace}`;
          ctx.textBaseline = "top";

          for (let li = 0; li < lines.length; li++) {
            const line = lines[li];
            const metrics = ctx.measureText(line);
            let drawX = xPx;
            if (it.align === "center") drawX = xPx + (wPx - metrics.width) / 2;
            if (it.align === "right") drawX = xPx + wPx - metrics.width;
            const drawY = yPx + li * lineHeight;
            if (drawY > yPx + hPx) break;
            ctx.fillText(line, drawX, drawY);
          }

          ctx.restore();
        }
      }
    }

    const jpgDataUrl = canvas.toDataURL("image/jpeg", jpgQuality);
    const jpgBytes = Uint8Array.from(atob(jpgDataUrl.split(",")[1]), (c) => c.charCodeAt(0));
    pages.push({
      jpgBytes,
      imgWidth: canvas.width,
      imgHeight: canvas.height,
      pageWidth: Math.round(targetWpx * 0.75),
      pageHeight: Math.round(targetHpx * 0.75)
    });
  }

  setProgress("pdf", slides.length, slides.length, "Writing file…", "Finalizing PDF…");
  if (uiCancelRequested) throw new Error("CANCELLED_UI");
  const pdfBytes = buildPdfBytes(pages);
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "Lucy_batch.pdf";
  a.click();
  URL.revokeObjectURL(url);

  setProgress("done", 1, 1, `Done — ${slides.length} slides`, "Export complete ✅");
  setState("success");
}

window.onmessage = async (event) => {
  const msg = event.data?.pluginMessage;
  if (!msg) return;

  if (msg.type === "STATUS") { setStatus(msg.text); return; }

  if (msg.type === "ERROR") {
    setStatus("Error:\n" + msg.text);
    setBusy(false, "Export PPTX");
    setState("error");
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
    uiCancelRequested = false;
    return;
  }

  // ✅ New protocol (Smart BG + items) — matches current code.ts
  if (msg.type === "BATCH_BG_AND_ITEMS_V051") {
    const slides: ExportSlide[] = msg.slides || [];
    if (!slides.length) {
      setStatus("Error: empty batch.");
      setBusy(false, "Export PPTX");
      setState("error");
      return;
    }
    try {
      if (msg.format === "pdf") {
        await buildPdfFromSlides((msg.filename ?? "Lucy_batch.pdf").replace(/\.pptx$/i, ".pdf"), slides, msg.quality || "best");
      } else {
        await buildPptxFromSlides(msg.filename ?? "Lucy_batch.pptx", slides);
      }
    } catch (err: any) {
      if (err?.message === "CANCELLED_UI") {
        setProgress("cancelled", 0, 1, "Cancelled", "Export cancelled.");
        setState("idle");
        return;
      }
      throw err;
    } finally {
      setBusy(false, "Export PPTX");
      uiCancelRequested = false;
    }
    return;
  }

  // legacy handler (keep if you still receive it somewhere)
  if (msg.type === "BATCH_BG_AND_ITEMS_V040") {
    const slides: ExportSlide[] = msg.slides || [];
    if (!slides.length) {
      setStatus("Error: empty batch.");
      setBusy(false, "Export PPTX");
      setState("error");
      return;
    }
    try {
      if (msg.format === "pdf") {
        await buildPdfFromSlides((msg.filename ?? "Lucy_batch.pdf").replace(/\.pptx$/i, ".pdf"), slides, msg.quality || "best");
      } else {
        await buildPptxFromSlides(msg.filename ?? "Lucy_batch.pptx", slides);
      }
    } catch (err: any) {
      if (err?.message === "CANCELLED_UI") {
        setProgress("cancelled", 0, 1, "Cancelled", "Export cancelled.");
        setState("idle");
        return;
      }
      throw err;
    } finally {
      setBusy(false, "Export PPTX");
      uiCancelRequested = false;
    }
    return;
  }
};
function getSegmentedValue(groupEl: HTMLDivElement | null, fallback: string) {
  const active = groupEl?.querySelector<HTMLButtonElement>(".segment.isActive");
  return active?.dataset.value || fallback;
}

function setupSegmented(groupEl: HTMLDivElement | null, fallback: string) {
  if (!groupEl) return;
  const buttons = Array.from(groupEl.querySelectorAll<HTMLButtonElement>(".segment"));
  if (!buttons.some((btn) => btn.classList.contains("isActive"))) {
    const match = buttons.find((btn) => btn.dataset.value === fallback);
    if (match) match.classList.add("isActive");
  }
  groupEl.addEventListener("click", (event) => {
    const target = (event.target as HTMLElement).closest<HTMLButtonElement>(".segment");
    if (!target || !groupEl.contains(target)) return;
    buttons.forEach((btn) => btn.classList.remove("isActive"));
    target.classList.add("isActive");
  });
}

setupSegmented(formatSelect, "pptx");
setupSegmented(qualitySelect, "best");
