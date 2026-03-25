const PptxGenJS = require("pptxgenjs");

function pxToIn(px) {
  return px / 96;
}

function clamp(n, a, b) {
  return Math.max(a, Math.min(b, n));
}

const FONT_SCALE = 0.705;
const TEXT_BOX_W_PAD_PX = 10;
const TEXT_BOX_H_PAD_PX = 2;
const TEXT_NUDGE_X_PX = -2;
const TEXT_HEIGHT_PAD_PX = 4;
const RADIUS_SCALE = 1.1;

function pxToPt(px) {
  return px * FONT_SCALE;
}

function getTextNudgeYPx(fontSizePx) {
  if (fontSizePx >= 28) return 1;
  if (fontSizePx >= 16) return 2;
  return 1;
}

function opacityToTransparencyPct(opacity01) {
  const o = typeof opacity01 === "number" ? clamp(opacity01, 0, 1) : 1;
  return Math.round((1 - o) * 100);
}

function figmaRadiusPxToRectRadiusRatio(radiusPx, wPx, hPx) {
  const r0 = typeof radiusPx === "number" ? Math.max(0, radiusPx) : 0;
  if (r0 <= 0) return 0;
  const r = r0 * RADIUS_SCALE;
  const halfMin = Math.max(1, Math.min(wPx, hPx) / 2);
  return clamp(r / halfMin, 0, 1);
}

function normalizeFontName(name) {
  return (name || "").trim().replace(/\s+/g, " ");
}

function mapFontFamily(figmaFamily) {
  const clean = normalizeFontName(figmaFamily);
  return clean || "Calibri";
}

function buildTransformForSlide(targetWpx, targetHpx, srcWpx, srcHpx) {
  const s = Math.min(targetWpx / srcWpx, targetHpx / srcHpx);
  const outW = srcWpx * s;
  const outH = srcHpx * s;
  const ox = (targetWpx - outW) / 2;
  const oy = (targetHpx - outH) / 2;
  return { s, ox, oy, outW, outH };
}

function toDataUrl(bytes, mimeType) {
  return `data:${mimeType};base64,${Buffer.from(bytes).toString("base64")}`;
}

function pngDataUrlFromItem(item, keyBase) {
  const base64Key = `${keyBase}Base64`;
  const bytesKey = `${keyBase}Bytes`;
  if (typeof item?.[base64Key] === "string" && item[base64Key].length > 0) {
    return `data:image/png;base64,${item[base64Key]}`;
  }
  if (Array.isArray(item?.[bytesKey]) && item[bytesKey].length > 0) {
    return toDataUrl(item[bytesKey], "image/png");
  }
  return null;
}

async function buildPptxBuffer(slides) {
  const targetWpx = Math.max(...slides.map((s) => s.width));
  const targetHpx = Math.max(...slides.map((s) => s.height));

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "FIGMA_BATCH", width: pxToIn(targetWpx), height: pxToIn(targetHpx) });
  pptx.layout = "FIGMA_BATCH";

  for (const sd of slides) {
    const trf = buildTransformForSlide(targetWpx, targetHpx, sd.width, sd.height);
    const slide = pptx.addSlide();
    const bgDataUrl = pngDataUrlFromItem(sd, "bgPng");
    const hasBgPng = !!bgDataUrl;
    const hasBgShape = !!sd.bgShape && !!sd.bgShape.fill;

    if (hasBgPng) {
      slide.addImage({
        data: bgDataUrl,
        x: pxToIn(trf.ox),
        y: pxToIn(trf.oy),
        w: pxToIn(trf.outW),
        h: pxToIn(trf.outH)
      });
    } else if (hasBgShape) {
      const op = typeof sd.bgShape.opacity === "number" ? sd.bgShape.opacity : 1;
      const tPct = opacityToTransparencyPct(op);
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: pxToIn(targetWpx),
        h: pxToIn(targetHpx),
        fill: { color: String(sd.bgShape.fill), transparency: tPct },
        line: { color: String(sd.bgShape.fill), transparency: 100 }
      });
    }

    const items = (sd.items || []).slice().sort((a, b) => (a.z ?? 0) - (b.z ?? 0));
    for (const it of items) {
      const sx = (v) => trf.ox + v * trf.s;
      const sy = (v) => trf.oy + v * trf.s;
      const sw = (v) => v * trf.s;
      const sh = (v) => v * trf.s;

      if (it.kind === "raster") {
        const rasterDataUrl = pngDataUrlFromItem(it, "png");
        if (!rasterDataUrl) continue;
        slide.addImage({
          data: rasterDataUrl,
          x: pxToIn(sx(it.x)),
          y: pxToIn(sy(it.y)),
          w: pxToIn(sw(it.w)),
          h: pxToIn(sh(it.h))
        });
        continue;
      }

      if (it.kind === "maskedImage") {
        const maskedDataUrl = pngDataUrlFromItem(it, "png");
        if (!maskedDataUrl) continue;
        slide.addImage({
          data: maskedDataUrl,
          x: pxToIn(sx(it.x)),
          y: pxToIn(sy(it.y)),
          w: pxToIn(sw(it.w)),
          h: pxToIn(sh(it.h)),
          sizing: {
            type: "crop",
            x: it.crop.x,
            y: it.crop.y,
            w: it.crop.w,
            h: it.crop.h
          }
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
        const lineProps = it.stroke
          ? { color: it.stroke.color, width: pxToIn(sw(it.stroke.width)), transparency: tPct }
          : undefined;

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
        const rawText = String(it.text);
        const finalText = it.uppercase ? rawText.toUpperCase() : rawText;
        const baseTextOpts = {
          x: pxToIn(xPx),
          y: pxToIn(yPx),
          w: pxToIn(wPx),
          h: pxToIn(hPx),
          margin: 0,
          align: it.align || "left",
          valign: "top",
          inset: 0,
          transparency: tPct
        };

        if (Array.isArray(it.runs) && it.runs.length > 0) {
          const richRuns = it.runs
            .map((r) => {
              const runTextRaw = String(r.text || "");
              if (!runTextRaw.length) return null;
              const runText = r.uppercase ? runTextRaw.toUpperCase() : runTextRaw;
              const runFsPx = Number(r.fontSize || it.fontSize || 14) * trf.s;
              const runLhPx = typeof r.lineHeightPx === "number" ? r.lineHeightPx * trf.s : null;
              const runLineSpacingPt = runLhPx ? Math.max(1, Math.round(pxToPt(runLhPx))) : undefined;
              return {
                text: runText,
                options: {
                  fontFace: mapFontFamily(r.fontFamily || it.fontFamily),
                  fontSize: Math.max(1, Math.round(pxToPt(runFsPx))),
                  bold: !!r.bold,
                  italic: !!r.italic,
                  color: r.color || it.color || "000000",
                  ...(runLineSpacingPt ? { lineSpacing: runLineSpacingPt } : {})
                }
              };
            })
            .filter(Boolean);

          if (richRuns.length > 0) {
            slide.addText(richRuns, baseTextOpts);
            continue;
          }
        }

        slide.addText(finalText, {
          ...baseTextOpts,
          fontFace: mapFontFamily(it.fontFamily),
          fontSize: Math.max(1, Math.round(pxToPt(effFsPx))),
          bold: !!it.bold,
          italic: !!it.italic,
          color: it.color || "000000",
          ...(lineSpacingPt ? { lineSpacing: lineSpacingPt } : {})
        });
      }
    }
  }

  return pptx.write({ outputType: "nodebuffer" });
}

module.exports = { buildPptxBuffer };
