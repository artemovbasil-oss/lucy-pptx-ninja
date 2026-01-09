// src/ui.ts (v0.3.5)
import PptxGenJS from "pptxgenjs";

const exportBtn = document.getElementById("export") as HTMLButtonElement;
const statusEl = document.getElementById("status") as HTMLDivElement;

function setStatus(msg: string) {
  statusEl.textContent = msg;
}

function pxToIn(px: number) {
  return px / 96;
}

// PptxGenJS fontSize is POINTS. Figma gives px.
// Base conversion is px*0.75, but empirically PPT renders a bit larger,
// so we add a calibration factor.
const FONT_SCALE = 0.70; // <-- tweak if needed (0.70..0.74 обычно хватает)

function pxToPt(px: number) {
  return px * FONT_SCALE;
}

function clamp(n: number, a: number, b: number) {
  return Math.max(a, Math.min(b, n));
}

// PptxGenJS uses transparency 0..100 (percent).
function opacityToTransparencyPct(opacity01: number | undefined) {
  const o = typeof opacity01 === "number" ? clamp(opacity01, 0, 1) : 1;
  return Math.round((1 - o) * 100);
}

// For rounded rectangles PptxGenJS supports rectRadius.
// Docs show 0..1 ratio.  [oai_citation:4‡gitbrent.github.io](https://gitbrent.github.io/PptxGenJS/docs/api-shapes/?utm_source=chatgpt.com)
function figmaRadiusPxToRectRadiusRatio(radiusPx: number | undefined, wPx: number, hPx: number) {
  const r = typeof radiusPx === "number" ? Math.max(0, radiusPx) : 0;
  if (r <= 0) return 0;
  const halfMin = Math.max(1, Math.min(wPx, hPx) / 2);
  return clamp(r / halfMin, 0, 1);
}

// Text tuning (your previous calibration)
const TEXT_NUDGE_X_PX = -2;
const TEXT_NUDGE_Y_PX = -2;
const TEXT_HEIGHT_PAD_PX = 4;

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
      setStatus("Lucy: building PPTX v0.3.5…");

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

          // Fill with transparency (0..100).  [oai_citation:5‡gitbrent.github.io](https://gitbrent.github.io/PptxGenJS/docs/types.html?utm_source=chatgpt.com)
          const fillProps =
            it.fill
              ? { color: it.fill, transparency: tr }
              : undefined;

          // Line with transparency (0..100).  [oai_citation:6‡gitbrent.github.io](https://gitbrent.github.io/PptxGenJS/docs/types.html?utm_source=chatgpt.com)
          const lineProps =
            it.stroke
              ? { color: it.stroke.color, width: pxToIn(it.stroke.width), transparency: tr }
              : undefined;

          if (it.shape === "rect") {
            const radiusPx = typeof it.radius === "number" ? it.radius : 0;
            const rr = figmaRadiusPxToRectRadiusRatio(radiusPx, w, h);

            // Use roundRect ALWAYS for rects: if rr=0, PPT renders as normal rect-ish anyway.
            // Key: rectRadius controls corner rounding.  [oai_citation:7‡gitbrent.github.io](https://gitbrent.github.io/PptxGenJS/docs/api-shapes/?utm_source=chatgpt.com)
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
            // for lines use line transparency too
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

          const x = (it.x ?? 0) + TEXT_NUDGE_X_PX;
          const y = (it.y ?? 0) + TEXT_NUDGE_Y_PX;
          const w = it.w ?? 10;
          const h = (it.h ?? 10) + TEXT_HEIGHT_PAD_PX;

          const opacity = typeof it.opacity === "number" ? it.opacity : 1;
          const tr = opacityToTransparencyPct(opacity);

          slide.addText(String(it.text), {
            x: pxToIn(x),
            y: pxToIn(y),
            w: pxToIn(w),
            h: pxToIn(h),
            fontFace: it.fontFamily || "Arial",
            fontSize: Math.max(1, Math.round(pxToPt(it.fontSize || 14))),
            bold: !!it.bold,
            italic: !!it.italic,
            color: it.color || "000000",
            align: it.align || "left",
            valign: "top",
            transparency: tr // text transparency is supported.  [oai_citation:8‡gitbrent.github.io](https://gitbrent.github.io/PptxGenJS/docs/api-text?utm_source=chatgpt.com)
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