// src/ui.ts
import PptxGenJS from "pptxgenjs";

const exportBtn = document.getElementById("export") as HTMLButtonElement;
const statusEl = document.getElementById("status") as HTMLDivElement;

function setStatus(msg: string) {
  statusEl.textContent = msg;
}

function pxToIn(px: number) {
  return px / 96;
}

// Настройки подгонки
const TEXT_NUDGE_X_PX = -2;
const TEXT_NUDGE_Y_PX = -2;
const TEXT_HEIGHT_PAD_PX = 4;

// Шейпы обычно совпадают лучше, но если понадобится — можно подвинуть тоже
const SHAPE_NUDGE_X_PX = 0;
const SHAPE_NUDGE_Y_PX = 0;

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

  if (msg.type === "FRAME_BG_PNG_TEXT_SHAPES") {
    try {
      setStatus("Lucy: building PPTX (BG + shapes + text)…");

      const bgBytes = new Uint8Array(msg.bgPngBytes);
      const b64 = uint8ToBase64(bgBytes);

      const wIn = pxToIn(msg.frame.width);
      const hIn = pxToIn(msg.frame.height);

      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: "FIGMA", width: wIn, height: hIn });
      pptx.layout = "FIGMA";

      const slide = pptx.addSlide();

      // 1) Background image (without text & safe shapes)
      slide.addImage({
        data: "data:image/png;base64," + b64,
        x: 0,
        y: 0,
        w: wIn,
        h: hIn
      });

      // 2) Shapes
      const shapes = msg.shapes as Array<any>;
      for (const s of shapes) {
        const x = (s.x ?? 0) + SHAPE_NUDGE_X_PX;
        const y = (s.y ?? 0) + SHAPE_NUDGE_Y_PX;

        if (s.kind === "rect") {
          slide.addShape(pptx.ShapeType.roundRect, {
            x: pxToIn(x),
            y: pxToIn(y),
            w: pxToIn(s.w ?? 10),
            h: pxToIn(s.h ?? 10),
            fill: s.fill ? { color: s.fill } : undefined,
            line: s.stroke ? { color: s.stroke.color, width: pxToIn(s.stroke.width) } : undefined
          });
        } else if (s.kind === "ellipse") {
          slide.addShape(pptx.ShapeType.ellipse, {
            x: pxToIn(x),
            y: pxToIn(y),
            w: pxToIn(s.w ?? 10),
            h: pxToIn(s.h ?? 10),
            fill: s.fill ? { color: s.fill } : undefined,
            line: s.stroke ? { color: s.stroke.color, width: pxToIn(s.stroke.width) } : undefined
          });
        } else if (s.kind === "line") {
          slide.addShape(pptx.ShapeType.line, {
            x: pxToIn(x),
            y: pxToIn(y),
            w: pxToIn(s.w ?? 10),
            h: pxToIn(s.h ?? 0),
            line: { color: s.stroke.color, width: pxToIn(s.stroke.width) }
          });
        }
      }

      // 3) Editable text overlay
      const texts = msg.texts as Array<any>;
      for (const t of texts) {
        if (!t.text || String(t.text).length === 0) continue;

        const x = (t.x ?? 0) + TEXT_NUDGE_X_PX;
        const y = (t.y ?? 0) + TEXT_NUDGE_Y_PX;
        const w = t.w ?? 10;
        const h = (t.h ?? 10) + TEXT_HEIGHT_PAD_PX;

        slide.addText(String(t.text), {
          x: pxToIn(x),
          y: pxToIn(y),
          w: pxToIn(w),
          h: pxToIn(h),
          fontFace: t.fontFamily || "Arial",
          fontSize: Math.max(1, Math.round(t.fontSize || 14)),
          color: t.color || "000000",
          align: t.align || "left",
          valign: "top"
        });
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

      setStatus(`Done ✅ (shapes: ${shapes.length}, texts: ${texts.length})`);
    } catch (e: any) {
      setStatus("Error (UI PPTX):\n" + (e?.message ?? String(e)));
    }
  }
};
