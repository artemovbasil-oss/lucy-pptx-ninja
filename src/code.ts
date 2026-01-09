// src/code.ts
figma.showUI(__html__, { width: 320, height: 300 });

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
  id: string;
  x: number; y: number; w: number; h: number;
  text: string;
  fontFamily: string;
  fontSize: number;
  color: string;
  align: "left" | "center" | "right" | "justify";
  opacity: number;
};

type ExportShape =
  | {
      id: string;
      kind: "rect" | "ellipse";
      x: number; y: number; w: number; h: number;
      fill: string | null;
      stroke: { color: string; width: number } | null;
      radius: number; // for rect
      opacity: number;
    }
  | {
      id: string;
      kind: "line";
      x: number; y: number; w: number; h: number;
      stroke: { color: string; width: number };
      opacity: number;
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

// ---- Shape helpers (only very safe shapes) ----
function isRotationZero(node: SceneNode): boolean {
  // @ts-ignore
  const rot = typeof node.rotation === "number" ? (node.rotation as number) : 0;
  return Math.abs(rot) < 0.01;
}
function hasAnyEffects(node: SceneNode): boolean {
  if (!("effects" in node)) return false;
  const eff = node.effects;
  if (!eff || eff === figma.mixed) return true;
  return (eff as readonly Effect[]).length > 0;
}
function getSolidFill(node: SceneNode): string | null {
  if (!("fills" in node)) return null;
  const fills = node.fills;
  if (!fills || fills === figma.mixed) return null;
  const onlySolid = (fills as readonly Paint[]).every((p) => p.type === "SOLID");
  if (!onlySolid) return null;
  const solid = (fills as readonly Paint[]).find((p) => p.type === "SOLID") as SolidPaint | undefined;
  return solid ? rgbToHex(solid.color) : null;
}
function getSolidStroke(node: SceneNode): { color: string; width: number } | null {
  if (!("strokes" in node) || !("strokeWeight" in node)) return null;
  const strokes = node.strokes;
  // @ts-ignore
  const sw = node.strokeWeight as number;
  if (!strokes || strokes === figma.mixed) return null;
  const onlySolid = (strokes as readonly Paint[]).every((p) => p.type === "SOLID");
  if (!onlySolid) return null;
  const solid = (strokes as readonly Paint[]).find((p) => p.type === "SOLID") as SolidPaint | undefined;
  if (!solid) return null;
  return { color: rgbToHex(solid.color), width: typeof sw === "number" ? sw : 1 };
}

figma.ui.onmessage = async (msg) => {
  try {
    if (msg.type !== "EXPORT_PPTX") return;

    const frame = getSelectedFrame();
    if (!frame) {
      postError("Select exactly ONE Frame.");
      return;
    }

    postStatus("Lucy: collecting layers…");

    const texts: ExportText[] = [];
    const shapes: ExportShape[] = [];

    const textNodes: TextNode[] = [];
    const shapeNodes: SceneNode[] = [];

    function walk(node: SceneNode) {
      if ("visible" in node && node.visible === false) return;

      if (node.type === "TEXT") {
        const tn = node as TextNode;
        const r = rectRelativeToFrame(tn, frame);
        texts.push({
          id: tn.id,
          x: r.x, y: r.y, w: r.w, h: r.h,
          text: tn.characters ?? "",
          fontFamily: getFirstCharFontFamily(tn),
          fontSize: getFirstCharFontSize(tn),
          color: getFirstCharFillHex(tn),
          align: alignMap(tn.textAlignHorizontal),
          opacity: typeof tn.opacity === "number" ? tn.opacity : 1
        });
        textNodes.push(tn);
      }

      // safe shapes only
      const rotOk = isRotationZero(node);
      const effOk = !hasAnyEffects(node);

      if (rotOk && effOk) {
        if (node.type === "RECTANGLE") {
          const fill = getSolidFill(node);
          const stroke = getSolidStroke(node);
          // @ts-ignore
          const radius = typeof (node as any).cornerRadius === "number" ? (node as any).cornerRadius : 0;

          // если нет ни fill ни stroke — смысла нет
          if (fill || stroke) {
            const r = rectRelativeToFrame(node, frame);
            shapes.push({
              id: node.id,
              kind: "rect",
              x: r.x, y: r.y, w: r.w, h: r.h,
              fill,
              stroke,
              radius: radius ?? 0,
              opacity: typeof (node as any).opacity === "number" ? (node as any).opacity : 1
            });
            shapeNodes.push(node);
          }
        } else if (node.type === "ELLIPSE") {
          const fill = getSolidFill(node);
          const stroke = getSolidStroke(node);
          if (fill || stroke) {
            const r = rectRelativeToFrame(node, frame);
            shapes.push({
              id: node.id,
              kind: "ellipse",
              x: r.x, y: r.y, w: r.w, h: r.h,
              fill,
              stroke,
              radius: 0,
              opacity: typeof (node as any).opacity === "number" ? (node as any).opacity : 1
            });
            shapeNodes.push(node);
          }
        } else if (node.type === "LINE") {
          const stroke = getSolidStroke(node);
          if (stroke) {
            const r = rectRelativeToFrame(node, frame);
            shapes.push({
              id: node.id,
              kind: "line",
              x: r.x, y: r.y, w: r.w, h: r.h,
              stroke,
              opacity: typeof (node as any).opacity === "number" ? (node as any).opacity : 1
            });
            shapeNodes.push(node);
          }
        }
      }

      if ("children" in node) {
        for (const ch of node.children) walk(ch as SceneNode);
      }
    }

    walk(frame);

    // ---- Export BG without text + without safe shapes ----
    postStatus("Lucy: exporting background (without text & safe shapes)…");

    const prevVisible = new Map<string, boolean>();

    // hide texts + shapes
    for (const tn of textNodes) {
      prevVisible.set(tn.id, tn.visible);
      tn.visible = false;
    }
    for (const sn of shapeNodes) {
      // @ts-ignore
      prevVisible.set(sn.id, sn.visible);
      // @ts-ignore
      sn.visible = false;
    }

    const scale = 2;
    let pngBytes: Uint8Array;
    try {
      pngBytes = await frame.exportAsync({
        format: "PNG",
        constraint: { type: "SCALE", value: scale }
      });
    } finally {
      // restore
      for (const tn of textNodes) {
        const v = prevVisible.get(tn.id);
        if (typeof v === "boolean") tn.visible = v;
      }
      for (const sn of shapeNodes) {
        const v = prevVisible.get(sn.id);
        // @ts-ignore
        if (typeof v === "boolean") sn.visible = v;
      }
    }

    figma.ui.postMessage({
      type: "FRAME_BG_PNG_TEXT_SHAPES",
      bgPngBytes: Array.from(pngBytes),
      filename: `${frame.name}.pptx`,
      frame: { name: frame.name, width: frame.width, height: frame.height, scale },
      texts,
      shapes
    });

    postStatus(`Lucy: sent BG + ${texts.length} texts + ${shapes.length} shapes ✅`);
  } catch (e: any) {
    postError(e?.message ?? String(e));
  }
};
