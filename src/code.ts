// src/code.ts (v0.3.5-dev)
figma.showUI(__html__, { width: 320, height: 320 });

function postStatus(text: string) {
  figma.ui.postMessage({ type: "STATUS", text });
}
function postError(text: string) {
  figma.ui.postMessage({ type: "ERROR", text });
}

function getSelectedFrame(): FrameNode | null {
  const sel = figma.currentPage.selection;
  if (!sel || sel.length !== 1) return null;
  return sel[0].type === "FRAME" ? (sel[0] as FrameNode) : null;
}

function clamp(n: number, a: number, b: number) {
  return Math.max(a, Math.min(b, n));
}
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
  kind: "text";
  z: number;
  id: string;
  x: number; y: number; w: number; h: number;
  text: string;
  fontFamily: string;
  fontSize: number; // px from Figma
  lineHeightPx?: number | null; // new: px line-height (best effort)
  color: string;
  align: "left" | "center" | "right" | "justify";
  opacity: number;
  bold: boolean;
  italic: boolean;
};

type ExportShape =
  | {
      kind: "shape";
      z: number;
      id: string;
      shape: "rect" | "ellipse";
      x: number; y: number; w: number; h: number;
      fill: string | null;
      stroke: { color: string; width: number } | null;
      radius: number;
      opacity: number;
    }
  | {
      kind: "shape";
      z: number;
      id: string;
      shape: "line";
      x: number; y: number; w: number; h: number;
      stroke: { color: string; width: number };
      opacity: number;
    };

type ExportRaster = {
  kind: "raster";
  z: number;
  id: string;
  x: number; y: number; w: number; h: number;
  pngBytes: number[];
};

type ExportItem = ExportText | ExportShape | ExportRaster;

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
  } catch {
    return "Arial";
  }
}
function getFirstCharFontSize(tn: TextNode): number {
  try {
    const len = tn.characters?.length ?? 0;
    if (len === 0) return 14;
    const fs = tn.getRangeFontSize(0, 1) as number;
    return typeof fs === "number" ? fs : 14;
  } catch {
    return 14;
  }
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
      bold:
        style.includes("bold") ||
        style.includes("semibold") ||
        style.includes("demibold") ||
        style.includes("heavy") ||
        style.includes("black"),
      italic: style.includes("italic") || style.includes("oblique")
    };
  } catch {
    return { bold: false, italic: false };
  }
}

// New: best-effort line-height (px)
function getTextLineHeightPx(tn: TextNode, fontSizePx: number): number | null {
  try {
    const lh = tn.lineHeight;
    if (!lh || lh === figma.mixed) return null;
    if (lh.unit === "AUTO") return null;
    if (lh.unit === "PIXELS") return typeof lh.value === "number" ? lh.value : null;
    if (lh.unit === "PERCENT") return typeof lh.value === "number" ? (fontSizePx * lh.value) / 100 : null;
    return null;
  } catch {
    return null;
  }
}

// ---------- SAFE HELPERS ----------
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
  if (!hasOnlySolidFills(node)) return false;
  if ((node as any).strokes === figma.mixed) return false;
  return true;
}
function isSafeEditableEllipse(node: EllipseNode): boolean {
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (!hasOnlySolidFills(node)) return false;
  return true;
}
function isSafeEditableLine(node: LineNode): boolean {
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  return !!getSolidStroke(node);
}

// Container background (pills/cards)
function isSafeEditableContainerBg(node: SceneNode): boolean {
  if (!("fills" in (node as any))) return false;
  if (!isRotationZero(node)) return false;
  if (hasAnyEffects(node)) return false;
  if (hasImageFill(node)) return false;
  if (!hasOnlySolidFills(node)) return false;
  if ("strokes" in (node as any) && (node as any).strokes === figma.mixed) return false;

  const fill = getSolidFill(node);
  const stroke = getSolidStroke(node);
  return !!(fill || stroke);
}

// ---------- CONTAINER TEXT DETECTOR ----------
function isContainer(node: SceneNode): boolean {
  return (
    node.type === "FRAME" ||
    node.type === "GROUP" ||
    node.type === "INSTANCE" ||
    node.type === "COMPONENT" ||
    node.type === "COMPONENT_SET"
  );
}

function containsTextDescendant(node: SceneNode): boolean {
  if (!("children" in node)) return false;
  const arr: SceneNode[] = [];
  for (const c of (node.children as any)) arr.push(c as SceneNode);

  while (arr.length) {
    const n = arr.pop()!;
    if ("visible" in n && (n as any).visible === false) continue;
    if (n.type === "TEXT") return true;
    if ("children" in n) {
      for (const ch of (n.children as any)) arr.push(ch as SceneNode);
    }
  }
  return false;
}

// ---------- RASTER POLICY ----------
const RASTER_MAX_W = 420;
const RASTER_MAX_H = 420;

function shouldRasterOverlay(node: SceneNode): boolean {
  if (!("visible" in node) || (node as any).visible === false) return false;
  if (node.type === "TEXT") return false;

  if (hasImageFill(node)) return true;

  if (
    node.type === "VECTOR" ||
    node.type === "BOOLEAN_OPERATION" ||
    node.type === "STAR" ||
    node.type === "POLYGON"
  ) return true;

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
  return await node.exportAsync({
    format: "PNG",
    constraint: { type: "SCALE", value: scale }
  });
}

// ---------- MAIN EXPORT ----------
figma.ui.onmessage = async (msg) => {
  try {
    if (msg.type !== "EXPORT_PPTX") return;

    const frame = getSelectedFrame();
    if (!frame) {
      postError("Select exactly ONE Frame.");
      return;
    }

    postStatus("Lucy: scanning layers (v0.3.5-dev)…");

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

    function walk(node: SceneNode) {
      if (!("visible" in node) || (node as any).visible === false) return;

      if (node.id !== frame.id) {
        z += 1;
        zById.set(node.id, z);

        // TEXT
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

        // SHAPES
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
                fill,
                stroke,
                radius,
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
                fill,
                stroke,
                radius: 0,
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

        // Container background as editable rect (but continue walking children)
        let exportedContainerBg = false;
        if (node.type === "FRAME" || node.type === "INSTANCE" || node.type === "COMPONENT") {
          if (isSafeEditableContainerBg(node)) {
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
            exportedContainerBg = true;
          }
        }

        // Raster overlay candidate (only if we didn't export container bg)
        if (!exportedContainerBg && shouldRasterOverlay(node)) {
          rasterCandidates.push(node);
          markHide(node);
          return;
        }
      }

      if ("children" in node) {
        for (const ch of node.children) walk(ch as SceneNode);
      }
    }

    walk(frame);

    // Rasterize
    if (rasterCandidates.length) {
      postStatus(`Lucy: rasterizing ${rasterCandidates.length} overlays…`);
    }

    const rasterItems: ExportRaster[] = [];
    for (let i = 0; i < rasterCandidates.length; i++) {
      const n = rasterCandidates[i];
      try {
        const r = rectRelativeToFrame(n, frame);
        if (r.w > frame.width * 0.98 && r.h > frame.height * 0.98) continue;
        if (r.w <= 0 || r.h <= 0) continue;

        postStatus(`Lucy: raster ${i + 1}/${rasterCandidates.length}…`);
        const bytes = await rasterizeNodePNG(n, 2);

        rasterItems.push({
          kind: "raster",
          z: zById.get(n.id) ?? 999999,
          id: n.id,
          x: r.x, y: r.y, w: r.w, h: r.h,
          pngBytes: Array.from(bytes)
        });
      } catch {
        // ignore
      }
    }

    const allItems: ExportItem[] = [...items, ...rasterItems];

    // Export clean BG (hide everything we will overlay)
    postStatus("Lucy: exporting background (clean)…");

    const prevVisible = new Map<string, boolean>();
    for (const n of toHide) {
      prevVisible.set(n.id, (n as any).visible);
      (n as any).visible = false;
    }

    const scale = 2;
    let bgPng: Uint8Array;
    try {
      bgPng = await frame.exportAsync({
        format: "PNG",
        constraint: { type: "SCALE", value: scale }
      });
    } finally {
      for (const n of toHide) {
        const v = prevVisible.get(n.id);
        if (typeof v === "boolean") (n as any).visible = v;
      }
    }

    figma.ui.postMessage({
      type: "FRAME_BG_AND_ITEMS_V031", // UI compatible
      bgPngBytes: Array.from(bgPng),
      filename: `${frame.name}.pptx`,
      frame: { name: frame.name, width: frame.width, height: frame.height, scale },
      items: allItems
    });

    postStatus(`Lucy: sent BG + overlays (${allItems.length}) ✅`);
  } catch (e: any) {
    postError(e?.message ?? String(e));
  }
};