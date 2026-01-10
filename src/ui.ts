// src/ui.ts (v0.3.7-dev)
import PptxGenJS from "pptxgenjs";

const exportBtn = document.getElementById("export") as HTMLButtonElement;
const statusEl = document.getElementById("status") as HTMLDivElement;

function setStatus(msg: string) {
  statusEl.textContent = msg;
}

function pxToIn(px: number) {
  return px / 96;
}

function clamp(n: number, a: number, b: number) {
  return Math.max(a, Math.min(b, n));
}

// --- TEXT SIZE CALIBRATION ---
// PptxGenJS fontSize is POINTS. Figma gives px.
// Base conversion is px*0.75, then we calibrate down a bit.
const FONT_SCALE = 0.705;
function pxToPt(px: number) {
  return px * FONT_SCALE;
}

// --- TEXT BOX FIXES (wrapping) ---
// PowerPoint text boxes have internal insets by default. We remove them and add width padding.
const TEXT_BOX_W_PAD_PX = 10;
const TEXT_BOX_H_PAD_PX = 2;

// Your previous X nudge
const TEXT_NUDGE_X_PX = -2;

// Auto Y nudge by font size (PPT baseline fix)
function getTextNudgeYPx(fontSizePx: number): number {
  if (fontSizePx >= 28) return 1;  // big headers: slightly down
  if (fontSizePx >= 16) return 2;  // main UI text: down a bit more
  return 1;                        // small text: down
}

// Extra height padding you already had
const TEXT_HEIGHT_PAD_PX = 4;

// --- SHAPE RADIUS CALIBRATION ---
const RADIUS_SCALE = 1.10;

// PptxGenJS uses transparency 0..100 (percent).
function opacityToTransparencyPct(opacity01: number | undefined) {
  const o = typeof opacity01 === "number" ? clamp(opacity01, 0, 1) : 1;
  return Math.round((1 - o) * 100);
}

// PptxGenJS roundRect uses rectRadius ratio (0..1)
function figmaRadiusPxToRectRadiusRatio(radiusPx: number | undefined, wPx: number, hPx: number) {
  const r0 = typeof radiusPx === "number" ? Math.max(0, radiusPx) : 0;
  if (r0 <= 0) return 0;

  const r = r0 * RADIUS_SCALE;
  const halfMin = Math.max(1, Math.min(wPx, hPx) / 2);
  return clamp(r / halfMin, 0, 1);
}

// --- FONT MAPPING (reduces random substitutions in PPT) ---
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

exportBtn.onclick = () => {
  setStatus("Lucy: requesting export…");
  parent.postMessage({ pluginMessage: { type: "EXPORT_PPTX" } }, "*");
};

window.onmessage = async (event) => {
  const msg = event.data?.pluginMessage;
  if (!msg) return;

  if (msg.type === "STATUS") setStatus(msg.text);
  if (msg.type === "ERROR") setStatus("Error:\n" + msg.text);

  if (msg.type === "FRAME_BG_AND_ITEMS_V031") {
    try {
      setStatus("Lucy: building PPTX v0.3.7-dev…");

      const bgBytes = new Uint8Array(msg.bgPngBytes);
      const bgB64 = uint8ToBase64(bgBytes);

      const wIn = pxToIn(msg.frame.width);
      const hIn = pxToIn(msg.frame.height);

      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: "FIGMA", width: wIn, height: hIn });
      pptx.layout = "FIGMA";

      const slide = pptx.addSlide();

      // 1) Clean background
      slide.addImage({
        data: "data:image/png;base64," + bgB64,
        x: 0,
        y: 0,
        w: wIn,
        h: hIn
      });

      // 2) Overlays in z-order
      const items = (msg.items as Array<any>).slice().sort((a, b) => (a.z ?? 0) - (b.z ?? 0));

      for (const it of items) {
        if (it.kind === "raster") {
          const bytes = new Uint8Array(it.pngBytes);
          const b64 = uint8ToBase64(bytes);

          slide.addImage({
            data: "data:image/png;base64," + b64,
            x: pxToIn(it.x),
            y: pxToIn(it.y),
            w: pxToIn(it.w),
            h: pxToIn(it.h)
          });
          continue;
        }

        if (it.kind === "shape") {
          const x = it.x ?? 0;
          const y = it.y ?? 0;
          const w = it.w ?? 10;
          const h = it.h ?? 10;

          const opacity = typeof it.opacity === "number" ? it.opacity : 1;
          const tr = opacityToTransparencyPct(opacity);

          const fillProps = it.fill ? { color: it.fill, transparency: tr } : undefined;
          const lineProps = it.stroke
            ? { color: it.stroke.color, width: pxToIn(it.stroke.width), transparency: tr }
            : undefined;

          if (it.shape === "rect") {
            const radiusPx = typeof it.radius === "number" ? it.radius : 0;
            const rr = figmaRadiusPxToRectRadiusRatio(radiusPx, w, h);

            slide.addShape(pptx.ShapeType.roundRect, {
              x: pxToIn(x),
              y: pxToIn(y),
              w: pxToIn(w),
              h: pxToIn(h),
              fill: fillProps,
              line: lineProps,
              rectRadius: rr
            });
          } else if (it.shape === "ellipse") {
            slide.addShape(pptx.ShapeType.ellipse, {
              x: pxToIn(x),
              y: pxToIn(y),
              w: pxToIn(w),
              h: pxToIn(h),
              fill: fillProps,
              line: lineProps
            });
          } else if (it.shape === "line") {
            slide.addShape(pptx.ShapeType.line, {
              x: pxToIn(x),
              y: pxToIn(y),
              w: pxToIn(w),
              h: pxToIn(h),
              line: lineProps ?? { color: it.stroke.color, width: pxToIn(it.stroke.width), transparency: tr }
            });
          }
          continue;
        }

        if (it.kind === "text") {
          if (!it.text || String(it.text).length === 0) continue;

          const fsPx = Number(it.fontSize || 14);

          const xPx = (it.x ?? 0) + TEXT_NUDGE_X_PX;
          const yPx = (it.y ?? 0) + getTextNudgeYPx(fsPx);

          const wPx = (it.w ?? 10) + TEXT_BOX_W_PAD_PX;
          const hPx = (it.h ?? 10) + TEXT_HEIGHT_PAD_PX + TEXT_BOX_H_PAD_PX;

          const opacity = typeof it.opacity === "number" ? it.opacity : 1;
          const tr = opacityToTransparencyPct(opacity);

          // line-height best effort: px -> pt
          const lhPx = typeof it.lineHeightPx === "number" ? it.lineHeightPx : null;
          const lineSpacingPt = lhPx ? Math.max(1, Math.round(pxToPt(lhPx))) : undefined;

          slide.addText(String(it.text), {
            x: pxToIn(xPx),
            y: pxToIn(yPx),
            w: pxToIn(wPx),
            h: pxToIn(hPx),

            // remove internal insets
            margin: 0,
            inset: 0,

            fontFace: mapFontFamily(it.fontFamily),
            fontSize: Math.max(1, Math.round(pxToPt(fsPx))),
            bold: !!it.bold,
            italic: !!it.italic,

            color: it.color || "000000",
            align: it.align || "left",
            valign: "top",
            transparency: tr,

            ...(lineSpacingPt ? { lineSpacing: lineSpacingPt } : {})
          });
          continue;
        }
      }

      setStatus("Lucy: rendering PPTX…");
      const arrayBuffer = await pptx.write("arraybuffer");
      const outBytes = new Uint8Array(arrayBuffer as ArrayBuffer);

      const blob = new Blob([outBytes], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = msg.filename ?? "export.pptx";
      a.click();
      URL.revokeObjectURL(url);

      setStatus(`Done ✅ (items: ${items.length})`);
    } catch (e: any) {
      setStatus("Error (UI PPTX):\n" + (e?.message ?? String(e)));
    }
  }
};