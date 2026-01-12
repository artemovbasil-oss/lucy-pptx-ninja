// src/code.ts (v0.5.3-dev) — adds: gradient RECTANGLE raster (big pills) + keeps previous fixes
figma.showUI(__html__, { width: 360, height: 520 });

function postStatus(text: string) { figma.ui.postMessage({ type: "STATUS", text }); }
function postError(text: string) { figma.ui.postMessage({ type: "ERROR", text }); }
function postProgress(phase: string, current: number, total: number, label?: string, text?: string) {
  figma.ui.postMessage({ type: "PROGRESS", phase, current, total, label, text });
}
function postCancelled() { figma.ui.postMessage({ type: "CANCELLED" }); }

let cancelRequested = false;
function throwIfCancelled() {
  if (cancelRequested) {
    const err: any = new Error("CANCELLED");
    err.__cancelled = true;
    throw err;
  }
}

function getSelectedFrames(): FrameNode[] {
  const sel = figma.currentPage.selection;
  if (!sel || sel.length === 0) return [];
  return sel.filter((n) => n.type === "FRAME") as FrameNode[];
}
function sendSelectionFrames() {
  const frames = getSelectedFrames();
  figma.ui.postMessage({
    type: "SELECTION_FRAMES",
    frames: frames.map((f) => ({ id: f.id, name: f.name, width: f.width, height: f.height }))
  });
}

function clamp(n: number, a: number, b: number) { return Math.max(a, Math.min(b, n)); }
function rgbToHex(rgb: RGB): string {
  const r = Math.round(clamp(rgb.r, 0, 1) * 255);
  const g = Math.round(clamp(rgb.g, 0, 1) * 255);
  const b = Math.round(clamp(rgb.b, 0, 1) * 255);
  return [r, g, b].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function getAbsXY(node: SceneNode): { x: number; y: number } {
  const t = node.absoluteTransform;
  return { x: t[0][2], y: t[1][2] };
}
function rectRelativeToFrame(node: SceneNode, frame: FrameNode) {
  const n = getAbsXY(node);
  const f = getAbsXY(frame);
  return { x: n.x - f.x, y: n.y - f.y, w: node.width, h: node.height };
}

type ExportText = {
  kind: "text"; z: number; id: string;
  x: number; y: number; w: number; h: number;
  text: string; fontFamily: string; fontSize: number;
  lineHeightPx?: number | null; color: string;
  align: "left" | "center" | "right" | "justify";
  opacity: number; bold: boolean; italic: boolean;
};

type ExportShape =
  | { kind: "shape"; z: number; id: string; shape: "rect" | "ellipse";
      x: number; y: number; w: number; h: number;
      fill: string | null; stroke: { color: string; width: number } | null;
      radius: number; opacity: number; }
  | { kind: "shape"; z: number; id: string; shape: "line";
      x: number; y: number; w: number; h: number;
      stroke: { color: string; width: number }; opacity: number; };

type ExportRaster = { kind: "raster"; z: number; id: string; x: number; y: number; w: number; h: number; pngBytes: number[]; };

type ExportMaskedImage = {
  kind: "maskedImage"; z: number; id: string;
  x: number; y: number; w: number; h: number;
  pngBytes: number[]; crop: { x: number; y: number; w: number; h: number };
};

type ExportItem = ExportText | ExportShape | ExportRaster | ExportMaskedImage;

type ExportSlide = {
  name: string; width: number; height: number; scale: number;
  bgPngBytes: number[]; bgShape?: { fill: string; opacity: number } | null;
  items: ExportItem[];
};

function alignMap(a: TextNode["textAlignHorizontal"]): ExportText["align"] {
  if (a === "CENTER") return "center";
  if (a === "RIGHT") return "right";
  if (a === "JUSTIFIED") return "justify";
  return "left";
}

function getFirstCharFontFamily(tn: TextNode): string {
  try {
    const len = tn.characters?.length ?? 0;
    if (len === 0) return "Arial";
    const fn = tn.getRangeFontName(0, 1) as FontName;
    return fn?.family || "Arial";
  } catch { return "Arial"; }
}
function getFirstCharFontSize(tn: TextNode): number {
  try {
    const len = tn.characters?.length ?? 0;
    if (len === 0) return 14;
    const fs = tn.getRangeFontSize(0, 1) as number;
    return typeof fs === "number" ? fs : 14;
  } catch { return 14; }
}
function getFirstCharFillHex(tn: TextNode): string {
  try {
    const len = tn.characters?.length ?? 0;
    if (len === 0) return "000000";
    const fills = tn.getRangeFills(0, 1) as readonly Paint[];
    const solid = fills?.find((p) => p.type === "SOLID") as SolidPaint | undefined;
    return solid ? rgbToHex(solid.color) : "000000";
  } catch {
    const fills = tn.fills;
    if (!fills || fills === figma.mixed) return "000000";
    const solid = (fills as readonly Paint[]).find((p) => p.type === "SOLID") as SolidPaint | undefined;
    return solid ? rgbToHex(solid.color) : "000000";
  }
}
function getFirstCharFontStyleFlags(tn: TextNode): { bold: boolean; italic: boolean } {
  try {
    const len = tn.characters?.length ?? 0;
    if (len === 0) return { bold: false, italic: false };
    const fn = tn.getRangeFontName(0, 1) as FontName;
    const style = (fn?.style || "").toLowerCase();
    return {
      bold: style.includes("bold") || style.includes("semibold") || style.includes("demibold") || style.includes("heavy") || style.includes("black"),
      italic: style.includes("italic") || style.includes("oblique")
    };
  } catch { return { bold: false, italic: false }; }
}
function getTextLineHeightPx(tn: TextNode, fontSizePx: number): number | null {
  try {
    const lh = tn.lineHeight;
    if (!lh || lh === figma.mixed) return null;
    if (lh.unit === "AUTO") return null;
    if (lh.unit === "PIXELS") return typeof lh.value === "number" ? lh.value : null;
    if (lh.unit === "PERCENT") return typeof lh.value === "number" ? (fontSizePx * lh.value) / 100 : null;
    return null;
  } catch { return null; }
}

// ---- helpers ----
function isRotationZero(node: SceneNode): boolean {
  const rot = typeof (node as any).rotation === "number" ? (node as any).rotation : 0;
  return Math.abs(rot) < 0.01;
}
function hasAnyEffects(node: SceneNode): boolean {
  if (!("effects" in node)) return false;
  const eff = (node as any).effects;
  if (!eff || eff === figma.mixed) return true;
  return (eff as readonly Effect[]).length > 0;
}
function hasImageFill(node: SceneNode): boolean {
  if (!("fills" in node)) return false;
  const fills = (node as any).fills;
  if (!fills || fills === figma.mixed) return true;
  return (fills as readonly Paint[]).some((p) => p.type === "IMAGE");
}
function hasOnlySolidFills(node: SceneNode): boolean {
  if (!("fills" in node)) return true;
  const fills = (node as any).fills;
  if (!fills || fills === figma.mixed) return false;
  return (fills as readonly Paint[]).every((p) => p.type === "SOLID");
}
function hasAnyGradientFill(node: SceneNode): boolean {
  if (!("fills" in node)) return false;
  const fills = (node as any).fills;
  if (!fills || fills === figma.mixed) return true;
  return (fills as readonly Paint[]).some((p) =>
    p.type === "GRADIENT_LINEAR" || p.type === "GRADIENT_RADIAL" || p.type === "GRADIENT_ANGULAR" || p.type === "GRADIENT_DIAMOND"
  );
}
function getSolidFill(node: SceneNode): string | null {
  if (!("fills" in node)) return null;
  const fills = (node as any).fills;
  if (!fills || fills === figma.mixed) return null;
  if (!(fills as readonly Paint[]).every((p) => p.type === "SOLID")) return null;
  const solid = (fills as readonly Paint[]).find((p) => p.type === "SOLID") as SolidPaint | undefined;
  return solid ? rgbToHex(solid.color) : null;
}
function getSolidStroke(node: SceneNode): { color: string; width: number } | null {
  if (!("strokes" in node) || !("strokeWeight" in node)) return null;
  const strokes = (node as any).strokes;
  const sw = (node as any).strokeWeight as number;
  if (!strokes || strokes === figma.mixed) return null;
  if (!(strokes as readonly Paint[]).every((p) => p.type === "SOLID")) return null;
  const solid = (strokes as readonly Paint[]).find((p) => p.type === "SOLID") as SolidPaint | undefined;
  if (!solid) return null;
  return { color: rgbToHex(solid.color), width: typeof sw === "number" ? sw : 1 };
}
function getCornerRadiusAny(node: SceneNode): number {
  const cr = (node as any).cornerRadius;
  return typeof cr === "number" ? cr : 0;
}

function isSafeEditableRect(node: RectangleNode): boolean {
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (hasAnyGradientFill(node)) return false;
  if (!hasOnlySolidFills(node)) return false;
  if ((node as any).strokes === figma.mixed) return false;
  return true;
}
function isSafeEditableEllipse(node: EllipseNode): boolean {
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (hasAnyGradientFill(node)) return false;
  if (!hasOnlySolidFills(node)) return false;
  return true;
}
function isSafeEditableLine(node: LineNode): boolean {
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  return !!getSolidStroke(node);
}

function isContainer(node: SceneNode): boolean {
  return (node.type === "FRAME" || node.type === "GROUP" || node.type === "INSTANCE" || node.type === "COMPONENT" || node.type === "COMPONENT_SET");
}
function containsTextDescendant(node: SceneNode): boolean {
  if (!("children" in node)) return false;
  const arr: SceneNode[] = [];
  for (const c of (node.children as any)) arr.push(c as SceneNode);
  while (arr.length) {
    const n = arr.pop()!;
    if ("visible" in n && (n as any).visible === false) continue;
    if (n.type === "TEXT") return true;
    if ("children" in n) for (const ch of (n.children as any)) arr.push(ch as SceneNode);
  }
  return false;
}

const RASTER_MAX_W = 420;
const RASTER_MAX_H = 420;

function isNearFullFrame(node: SceneNode, frame: FrameNode): boolean {
  const r = rectRelativeToFrame(node, frame);
  const wOk = r.w >= frame.width * 0.95;
  const hOk = r.h >= frame.height * 0.95;
  const xOk = Math.abs(r.x) <= frame.width * 0.03;
  const yOk = Math.abs(r.y) <= frame.height * 0.03;
  return wOk && hOk && xOk && yOk;
}

function shouldRasterOverlay(node: SceneNode, frame: FrameNode): boolean {
  if (!("visible" in node) || (node as any).visible === false) return false;
  if (node.type === "TEXT") return false;

  if (hasImageFill(node) && isNearFullFrame(node, frame)) return false;
  if (hasImageFill(node)) return true;

  if (node.type === "VECTOR" || node.type === "BOOLEAN_OPERATION" || node.type === "STAR" || node.type === "POLYGON") return true;
  if (node.type === "LINE" && !isSafeEditableLine(node as LineNode)) return true;

  const small = node.width <= RASTER_MAX_W && node.height <= RASTER_MAX_H;

  if (small && isContainer(node)) {
    if (containsTextDescendant(node)) return false;
    return true;
  }

  if (small) {
    if (node.type === "RECTANGLE" && !isSafeEditableRect(node as RectangleNode)) return true;
    if (node.type === "ELLIPSE" && !isSafeEditableEllipse(node as EllipseNode)) return true;
  }

  return false;
}

async function rasterizeNodePNG(node: SceneNode, scale = 2): Promise<Uint8Array> {
  throwIfCancelled();
  return await node.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: scale } });
}

// ---- Smart BG ----
function getSmartBackground(frame: FrameNode): { fill: string; opacity: number } | null {
  try {
    if (!isRotationZero(frame)) return null;
    if (hasAnyEffects(frame)) return null;
    if (hasImageFill(frame)) return null;
    if (hasAnyGradientFill(frame)) return null;
    if (!hasOnlySolidFills(frame)) return null;

    const fill = getSolidFill(frame);
    if (!fill) return null;

    const opacity = typeof frame.opacity === "number" ? frame.opacity : 1;
    return { fill, opacity };
  } catch { return null; }
}

// ---- Masks ----
function isRectMaskNode(n: SceneNode): n is RectangleNode {
  return n.type === "RECTANGLE" && (n as any).isMask === true;
}
function rectHasImageFill(r: RectangleNode): boolean {
  const fills = r.fills;
  if (!fills || fills === figma.mixed) return true;
  return (fills as readonly Paint[]).some((p) => p.type === "IMAGE");
}
function isSafeMaskPair(mask: RectangleNode, img: RectangleNode): boolean {
  if (!isRotationZero(mask) || !isRotationZero(img)) return false;
  if (hasAnyEffects(mask) || hasAnyEffects(img)) return false;
  if (!rectHasImageFill(img)) return false;
  return true;
}
function cropFromMaskAndImage(maskR: { x: number; y: number; w: number; h: number }, imgR: { x: number; y: number; w: number; h: number }) {
  const ix = imgR.x, iy = imgR.y, iw = Math.max(1e-6, imgR.w), ih = Math.max(1e-6, imgR.h);
  const mx = maskR.x, my = maskR.y, mw = maskR.w, mh = maskR.h;

  let cx = (mx - ix) / iw;
  let cy = (my - iy) / ih;
  let cw = mw / iw;
  let ch = mh / ih;

  cx = clamp(cx, 0, 1);
  cy = clamp(cy, 0, 1);
  cw = clamp(cw, 0, 1);
  ch = clamp(ch, 0, 1);

  if (cx + cw > 1) cw = clamp(1 - cx, 0, 1);
  if (cy + ch > 1) ch = clamp(1 - cy, 0, 1);

  return { x: cx, y: cy, w: cw, h: ch };
}

// ---- Container BG as shape (solid) ----
function isSafeContainerBg(node: SceneNode): boolean {
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (hasAnyGradientFill(node)) return false;
  if (!hasOnlySolidFills(node)) return false;
  const fill = getSolidFill(node);
  const stroke = getSolidStroke(node);
  return !!(fill || stroke);
}

// ---- Gradient container BG: rasterize bg-only (hide text descendants) ----
function collectTextDescendants(node: SceneNode): TextNode[] {
  const out: TextNode[] = [];
  if (!("children" in node)) return out;
  const stack: SceneNode[] = [...(node.children as readonly SceneNode[])];
  while (stack.length) {
    const n = stack.pop()!;
    if ("visible" in n && (n as any).visible === false) continue;
    if (n.type === "TEXT") out.push(n as TextNode);
    if ("children" in n) stack.push(...(n.children as readonly SceneNode[]));
  }
  return out;
}
async function rasterizeContainerBackgroundOnly(container: SceneNode, scale = 2): Promise<Uint8Array> {
  const texts = collectTextDescendants(container);
  const prev = new Map<string, boolean>();
  for (const t of texts) { prev.set(t.id, t.visible); t.visible = false; }
  try {
    return await container.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: scale } });
  } finally {
    for (const t of texts) {
      const v = prev.get(t.id);
      if (typeof v === "boolean") t.visible = v;
    }
  }
}
function isGradientContainerCandidate(node: SceneNode): boolean {
  if (!(node.type === "FRAME" || node.type === "INSTANCE" || node.type === "COMPONENT")) return false;
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (!hasAnyGradientFill(node)) return false;
  if (node.width > 1800 || node.height > 1000) return false;
  return true;
}

// ✅ NEW: big gradient rectangles (like your blue pill)
// Export as raster overlay regardless of size (bounded), because text is separate node
function isGradientRectPillCandidate(node: SceneNode): node is RectangleNode {
  if (node.type !== "RECTANGLE") return false;
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (!hasAnyGradientFill(node)) return false;
  // safety bounds (avoid accidental full-canvas gradients as overlay)
  if (node.width > 2200 || node.height > 1400) return false;
  return true;
}

async function exportOneFrame(frame: FrameNode, idx: number, total: number): Promise<ExportSlide> {
  throwIfCancelled();
  postProgress("export", idx - 1, total, `Scanning: ${frame.name}`, `Scanning frame ${idx}/${total}: ${frame.name}`);

  const items: ExportItem[] = [];
  const rasterCandidates: SceneNode[] = [];

  const toHideSet = new Set<string>();
  const toHide: SceneNode[] = [];
  function markHide(n: SceneNode) {
    if (toHideSet.has(n.id)) return;
    toHideSet.add(n.id);
    toHide.push(n);
  }

  const zById = new Map<string, number>();
  let z = 0;

  const consumedMaskIds = new Set<string>();
  const consumedMaskedContentIds = new Set<string>();

  async function tryExtractMaskPairsInContainer(container: SceneNode) {
    if (!("children" in container)) return;

    const ch = (container.children as readonly SceneNode[]).filter((n) => ("visible" in n ? (n as any).visible !== false : true));
    if (ch.length < 2) return;

    const first = ch[0];
    const second = ch[1];

    if (!isRectMaskNode(first)) return;
    if (second.type !== "RECTANGLE") return;

    const mask = first as RectangleNode;
    const img = second as RectangleNode;

    if (!isSafeMaskPair(mask, img)) return;

    const maskR = rectRelativeToFrame(mask, frame);
    const imgR = rectRelativeToFrame(img, frame);
    const maskRadius = getCornerRadiusAny(mask);

    if (maskRadius > 0.5) {
      postProgress("export", idx - 1, total, `Mask (rounded): ${frame.name}`, `Rasterizing rounded mask…`);
      const contRect = rectRelativeToFrame(container, frame);
      const bytes = await rasterizeNodePNG(container, 2);

      z += 1;
      const zVal = z;

      items.push({
        kind: "raster",
        z: zVal,
        id: `roundedMask__${container.id}`,
        x: contRect.x, y: contRect.y, w: contRect.w, h: contRect.h,
        pngBytes: Array.from(bytes)
      });

      markHide(container);
      consumedMaskIds.add(mask.id);
      consumedMaskedContentIds.add(img.id);
      return;
    }

    postProgress("export", idx - 1, total, `Mask image: ${frame.name}`, `Exporting masked image…`);
    const imgBytes = await rasterizeNodePNG(img, 2);
    const crop = cropFromMaskAndImage(maskR, imgR);

    z += 1;
    const zVal = z;
    zById.set(mask.id, zVal);
    zById.set(img.id, zVal);

    items.push({
      kind: "maskedImage",
      z: zVal,
      id: `${mask.id}__${img.id}`,
      x: maskR.x, y: maskR.y, w: maskR.w, h: maskR.h,
      pngBytes: Array.from(imgBytes),
      crop
    });

    consumedMaskIds.add(mask.id);
    consumedMaskedContentIds.add(img.id);

    markHide(mask);
    markHide(img);
  }

  async function walk(node: SceneNode) {
    if (!("visible" in node) || (node as any).visible === false) return;
    if (cancelRequested) return;

    if (node.id !== frame.id && isContainer(node)) {
      await tryExtractMaskPairsInContainer(node);
    }

    if (node.id !== frame.id) {
      if (consumedMaskIds.has(node.id) || consumedMaskedContentIds.has(node.id)) return;

      z += 1;
      zById.set(node.id, z);

      // 2) Container BG as editable shape (solid)
      if (node.type === "FRAME" || node.type === "INSTANCE" || node.type === "COMPONENT") {
        if (isSafeContainerBg(node)) {
          const r = rectRelativeToFrame(node, frame);
          const fill = getSolidFill(node);
          const stroke = getSolidStroke(node);
          const radius = getCornerRadiusAny(node);

          items.push({
            kind: "shape",
            z,
            id: node.id,
            shape: "rect",
            x: r.x, y: r.y, w: r.w, h: r.h,
            fill,
            stroke,
            radius,
            opacity: typeof (node as any).opacity === "number" ? (node as any).opacity : 1
          });

          markHide(node);
        }
      }

      // 2.1) Gradient container BG as raster (bg-only)
      if (isGradientContainerCandidate(node)) {
        const r = rectRelativeToFrame(node, frame);
        const isAlmostFull = r.w >= frame.width * 0.9 && r.h >= frame.height * 0.9;
        if (!isAlmostFull) {
          postProgress("export", idx - 1, total, `Gradient pill: ${frame.name}`, `Rasterizing gradient background…`);
          const bytes = await rasterizeContainerBackgroundOnly(node, 2);

          items.push({
            kind: "raster",
            z,
            id: `gradBg__${node.id}`,
            x: r.x, y: r.y, w: r.w, h: r.h,
            pngBytes: Array.from(bytes)
          });

          markHide(node);
        }
      }

      // ✅ NEW 2.2) Gradient RECTANGLE as raster overlay (big pills like screenshot)
      if (isGradientRectPillCandidate(node)) {
        const r = rectRelativeToFrame(node, frame);
        const isAlmostFull = r.w >= frame.width * 0.9 && r.h >= frame.height * 0.9;

        // If it's almost the entire slide, better leave it to BG PNG; otherwise export as overlay
        if (!isAlmostFull) {
          postProgress("export", idx - 1, total, `Gradient rect: ${frame.name}`, `Rasterizing gradient rectangle…`);
          const bytes = await rasterizeNodePNG(node, 2);

          items.push({
            kind: "raster",
            z,
            id: `gradRect__${node.id}`,
            x: r.x, y: r.y, w: r.w, h: r.h,
            pngBytes: Array.from(bytes)
          });

          markHide(node);
          return; // rectangle handled
        }
      }

      // 3) Text
      if (node.type === "TEXT") {
        const tn = node as TextNode;
        const r = rectRelativeToFrame(tn, frame);
        const flags = getFirstCharFontStyleFlags(tn);
        const fs = getFirstCharFontSize(tn);

        items.push({
          kind: "text",
          z,
          id: tn.id,
          x: r.x, y: r.y, w: r.w, h: r.h,
          text: tn.characters ?? "",
          fontFamily: getFirstCharFontFamily(tn),
          fontSize: fs,
          lineHeightPx: getTextLineHeightPx(tn, fs),
          color: getFirstCharFillHex(tn),
          align: alignMap(tn.textAlignHorizontal),
          opacity: typeof tn.opacity === "number" ? tn.opacity : 1,
          bold: flags.bold,
          italic: flags.italic
        });

        markHide(tn);
        return;
      }

      // 4) Basic shapes
      if (node.type === "RECTANGLE") {
        const rn = node as RectangleNode;
        if (isSafeEditableRect(rn)) {
          const r = rectRelativeToFrame(rn, frame);
          const fill = getSolidFill(rn);
          const stroke = getSolidStroke(rn);
          const radius = getCornerRadiusAny(rn);

          if (fill || stroke) {
            items.push({
              kind: "shape",
              z,
              id: rn.id,
              shape: "rect",
              x: r.x, y: r.y, w: r.w, h: r.h,
              fill, stroke, radius,
              opacity: typeof rn.opacity === "number" ? rn.opacity : 1
            });
            markHide(rn);
          }
          return;
        }
      }

      if (node.type === "ELLIPSE") {
        const en = node as EllipseNode;
        if (isSafeEditableEllipse(en)) {
          const r = rectRelativeToFrame(en, frame);
          const fill = getSolidFill(en);
          const stroke = getSolidStroke(en);

          if (fill || stroke) {
            items.push({
              kind: "shape",
              z,
              id: en.id,
              shape: "ellipse",
              x: r.x, y: r.y, w: r.w, h: r.h,
              fill, stroke, radius: 0,
              opacity: typeof en.opacity === "number" ? en.opacity : 1
            });
            markHide(en);
          }
          return;
        }
      }

      if (node.type === "LINE") {
        const ln = node as LineNode;
        if (isSafeEditableLine(ln)) {
          const r = rectRelativeToFrame(ln, frame);
          const stroke = getSolidStroke(ln)!;

          items.push({
            kind: "shape",
            z,
            id: ln.id,
            shape: "line",
            x: r.x, y: r.y, w: r.w, h: r.h,
            stroke,
            opacity: typeof ln.opacity === "number" ? ln.opacity : 1
          });
          markHide(ln);
          return;
        }
      }

      // 5) Raster fallback (icons/arrows/complex vectors etc.)
      if (shouldRasterOverlay(node, frame)) {
        rasterCandidates.push(node);
        markHide(node);
        return;
      }
    }

    if ("children" in node) {
      for (const ch of node.children as readonly SceneNode[]) {
        await walk(ch as SceneNode);
        if (cancelRequested) return;
      }
    }
  }

  await walk(frame);
  throwIfCancelled();

  // Rasterize overlays
  if (rasterCandidates.length) postProgress("export", idx - 1, total, `Rasterizing overlays: ${frame.name}`);

  const rasterItems: ExportRaster[] = [];
  for (let i = 0; i < rasterCandidates.length; i++) {
    throwIfCancelled();
    const n = rasterCandidates[i];

    postProgress("export", idx - 1, total, `Raster ${i + 1}/${rasterCandidates.length}: ${frame.name}`);

    try {
      const r = rectRelativeToFrame(n, frame);
      if (r.w <= 0 || r.h <= 0) continue;
      const bytes = await rasterizeNodePNG(n, 2);
      rasterItems.push({
        kind: "raster",
        z: zById.get(n.id) ?? 999999,
        id: n.id,
        x: r.x, y: r.y, w: r.w, h: r.h,
        pngBytes: Array.from(bytes)
      });
    } catch { /* ignore */ }
  }

  const allItems: ExportItem[] = [...items, ...rasterItems];

  // ---- Background: smart shape if possible ----
  const smartBg = getSmartBackground(frame);

  let bgPngBytes: number[] = [];
  let bgShape: ExportSlide["bgShape"] = null;

  if (smartBg) {
    bgShape = smartBg;
    bgPngBytes = [];
  } else {
    throwIfCancelled();
    postProgress("export", idx - 1, total, `Exporting background: ${frame.name}`);

    const prevVisible = new Map<string, boolean>();
    for (const n of toHide) { prevVisible.set(n.id, (n as any).visible); (n as any).visible = false; }

    let bgPng: Uint8Array;
    try {
      throwIfCancelled();
      bgPng = await frame.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: 2 } });
    } finally {
      for (const n of toHide) {
        const v = prevVisible.get(n.id);
        if (typeof v === "boolean") (n as any).visible = v;
      }
    }
    bgPngBytes = Array.from(bgPng);
  }

  throwIfCancelled();
  postProgress("export", idx, total, `Ready: ${frame.name}`);

  return {
    name: frame.name,
    width: frame.width,
    height: frame.height,
    scale: 2,
    bgPngBytes,
    bgShape,
    items: allItems
  };
}

// --- Messages ---
figma.ui.onmessage = async (msg) => {
  try {
    if (msg.type === "REQUEST_SELECTION") { sendSelectionFrames(); return; }

    if (msg.type === "CANCEL_EXPORT") {
      cancelRequested = true;
      postStatus("Cancel requested…");
      return;
    }

    if (msg.type === "EXPORT_PPTX_ORDERED") {
      cancelRequested = false;

      const ids: string[] = Array.isArray(msg.frameIds) ? msg.frameIds : [];
      if (!ids.length) { postError("No frames in export list."); return; }

      const nodes = await Promise.all(ids.map((id) => figma.getNodeByIdAsync(id)));
      const frames: FrameNode[] = nodes.filter((n): n is FrameNode => !!n && (n as any).type === "FRAME");

      if (!frames.length) { postError("Selected frames not found. Click Refresh and try again."); return; }

      postProgress("export", 0, frames.length, "Starting export…", `Exporting ${frames.length} frame(s)…`);

      const slides: ExportSlide[] = [];
      for (let i = 0; i < frames.length; i++) {
        throwIfCancelled();
        slides.push(await exportOneFrame(frames[i], i + 1, frames.length));
      }

      throwIfCancelled();

      const filename = frames.length === 1 ? `${frames[0].name}.pptx` : `Lucy_batch_${frames.length}_slides.pptx`;

      figma.ui.postMessage({ type: "BATCH_BG_AND_ITEMS_V051", filename, slides });
      postProgress("export", frames.length, frames.length, "Sent to PPTX builder", "Building PPTX…");
      return;
    }
  } catch (e: any) {
    if (e?.__cancelled || e?.message === "CANCELLED") { postCancelled(); return; }
    postError(e?.message ?? String(e));
  }
};