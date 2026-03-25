// src/code.ts (v0.5.3-dev) — adds: gradient RECTANGLE raster (big pills) + keeps previous fixes
declare const __LUCY_API_BASE_URL__: string;

figma.showUI(__html__, { width: 360, height: 600 });

function safeUiPostMessage(message: Record<string, any>): boolean {
  try {
    figma.ui.postMessage(message);
    return true;
  } catch {
    return false;
  }
}

function postStatus(text: string) { safeUiPostMessage({ type: "STATUS", text }); }
function postError(text: string) { safeUiPostMessage({ type: "ERROR", text }); }
function postProgress(phase: string, current: number, total: number, label?: string, text?: string) {
  safeUiPostMessage({ type: "PROGRESS", phase, current, total, label, text });
}
function postCancelled() { safeUiPostMessage({ type: "CANCELLED" }); }

let cancelRequested = false;
const DEFAULT_REMOTE_API_BASE = "https://lucy-pptx-ninja-production.up.railway.app";
const INJECTED_REMOTE_API_BASE = (typeof __LUCY_API_BASE_URL__ === "string" ? __LUCY_API_BASE_URL__ : "").trim().replace(/\/$/, "");
const REMOTE_API_CANDIDATES = Array.from(new Set([INJECTED_REMOTE_API_BASE, DEFAULT_REMOTE_API_BASE].filter(Boolean)));
let activeRemoteApiBase = REMOTE_API_CANDIDATES[0] || "";
let exportInProgress = false;
const thumbCache = new Map<string, { width: number; height: number; thumbBytes: number[] | null }>();
let selectionRefreshInFlight = false;
let selectionRefreshQueued = false;
const REMOTE_PPTX_MIN_FRAMES = 12;

function throwIfCancelled() {
  if (cancelRequested) {
    const err: any = new Error("CANCELLED");
    err.__cancelled = true;
    throw err;
  }
}

function shouldUseRemotePptx(format: string): boolean {
  return format === "pptx" && REMOTE_API_CANDIDATES.length > 0;
}

function canToggleNodeVisibility(node: SceneNode): boolean {
  try {
    const current = node.visible;
    node.visible = current;
    return true;
  } catch {
    return false;
  }
}

function getExportScale(format: string, quality: string, remotePptx: boolean): number {
  if (format === "pdf") {
    return quality === "low" ? 1 : quality === "medium" ? 1.5 : 2;
  }

  if (remotePptx) {
    return quality === "low" ? 1 : quality === "medium" ? 1.5 : 2;
  }

  return quality === "low" ? 1 : quality === "medium" ? 2 : 3;
}

async function fetchJson<T = any>(url: string, init?: RequestInit): Promise<T> {
  const mergedHeaders: Record<string, string> = {};
  const sourceHeaders: any = init?.headers;

  if (sourceHeaders) {
    if (typeof sourceHeaders.forEach === "function") {
      sourceHeaders.forEach((value: unknown, key: unknown) => {
        mergedHeaders[String(key)] = String(value);
      });
    } else if (Array.isArray(sourceHeaders)) {
      for (const entry of sourceHeaders) {
        if (Array.isArray(entry) && entry.length >= 2) {
          mergedHeaders[String(entry[0])] = String(entry[1]);
        }
      }
    } else if (typeof sourceHeaders === "object") {
      for (const [key, value] of Object.entries(sourceHeaders)) {
        mergedHeaders[String(key)] = String(value);
      }
    }
  }

  const hasContentType = Object.keys(mergedHeaders).some((key) => key.toLowerCase() === "content-type");
  if (!hasContentType) mergedHeaders["Content-Type"] = "application/json";

  const res = await fetch(url, { ...init, headers: mergedHeaders });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(text || `Request failed: ${res.status}`);
  }
  return await res.json() as T;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchRemoteJson<T = any>(path: string, init?: RequestInit): Promise<T> {
  const orderedBases = [activeRemoteApiBase, ...REMOTE_API_CANDIDATES.filter((b) => b !== activeRemoteApiBase)];
  let lastErr: unknown = null;

  for (const base of orderedBases) {
    try {
      const data = await fetchJson<T>(`${base}${path}`, init);
      activeRemoteApiBase = base;
      return data;
    } catch (err) {
      lastErr = err;
    }
  }

  throw lastErr instanceof Error ? lastErr : new Error("Remote API unavailable");
}

function collectFramesDeep(node: SceneNode, out: FrameNode[]) {
  if ((node as any).visible === false) return;
  if (node.type === "FRAME") {
    out.push(node as FrameNode);
    return;
  }
  if ("children" in node) {
    for (const c of node.children as readonly SceneNode[]) {
      collectFramesDeep(c as SceneNode, out);
    }
  }
}

type FrameSortMode = "layout" | "name";
const FRAME_SORT_MODE: FrameSortMode = "layout";

function getFrameSortKey(frame: FrameNode): { x: number; y: number } {
  const pos = getAbsXY(frame);
  return { x: pos.x, y: pos.y };
}

function sortFrames(frames: FrameNode[]): FrameNode[] {
  if (FRAME_SORT_MODE === "name") {
    return [...frames].sort((a, b) => {
      const byName = a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" });
      if (byName !== 0) return byName;
      return a.id.localeCompare(b.id);
    });
  }

  return [...frames].sort((a, b) => {
    const pa = getFrameSortKey(a);
    const pb = getFrameSortKey(b);
    if (pa.y !== pb.y) return pa.y - pb.y;
    if (pa.x !== pb.x) return pa.x - pb.x;
    const byName = a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" });
    if (byName !== 0) return byName;
    return a.id.localeCompare(b.id);
  });
}

function getSelectedFrames(): FrameNode[] {
  const sel = figma.currentPage.selection;
  if (!sel || sel.length === 0) return [];

  const out: FrameNode[] = [];
  for (const n of sel as readonly SceneNode[]) {
    if ((n as any).visible === false) continue;
    if (n.type === "FRAME") out.push(n as FrameNode);
    else if (n.type === "SECTION") collectFramesDeep(n as unknown as SceneNode, out);
  }

  // Deduplicate
  const map = new Map<string, FrameNode>();
  for (const f of out) map.set(f.id, f);
  return sortFrames(Array.from(map.values()));
}
async function sendSelectionFrames() {
  const frames = getSelectedFrames();
  const enriched: Array<{ id: string; name: string; width: number; height: number; thumbBytes: number[] | null }> = [];
  for (const f of frames) {
    let thumbBytes: number[] | null = null;
    const cached = thumbCache.get(f.id);
    if (cached && cached.width === f.width && cached.height === f.height) {
      thumbBytes = cached.thumbBytes;
    } else if (!exportInProgress) {
      try {
        const bytes = await f.exportAsync({ format: "PNG", constraint: { type: "WIDTH", value: 96 } });
        thumbBytes = Array.from(bytes);
      } catch {
        thumbBytes = null;
      }
      thumbCache.set(f.id, { width: f.width, height: f.height, thumbBytes });
    }
    enriched.push({ id: f.id, name: f.name, width: f.width, height: f.height, thumbBytes });
  }
  safeUiPostMessage({
    type: "SELECTION_FRAMES",
    frames: enriched
  });
}

async function refreshSelectionFramesSafely() {
  if (selectionRefreshInFlight) {
    selectionRefreshQueued = true;
    return;
  }
  selectionRefreshInFlight = true;
  try {
    await sendSelectionFrames();
  } finally {
    selectionRefreshInFlight = false;
    if (selectionRefreshQueued) {
      selectionRefreshQueued = false;
      void refreshSelectionFramesSafely();
    }
  }
}

// Auto-update selection without manual refresh
figma.on("selectionchange", () => {
  try { void refreshSelectionFramesSafely(); } catch { /* ignore */ }
});

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
function getNodeBounds(node: SceneNode): { x: number; y: number; w: number; h: number } {
  const bb = (node as any).absoluteBoundingBox as { x: number; y: number; width: number; height: number } | undefined;
  if (bb) return { x: bb.x, y: bb.y, w: bb.width, h: bb.height };
  const t = node.absoluteTransform;
  return { x: t[0][2], y: t[1][2], w: node.width, h: node.height };
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
  opacity: number; bold: boolean; italic: boolean; uppercase: boolean;
  runs?: ExportTextRun[];
};

type ExportTextRun = {
  text: string;
  fontFamily: string;
  fontSize: number;
  lineHeightPx?: number | null;
  color: string;
  bold: boolean;
  italic: boolean;
  uppercase: boolean;
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
type ExportRasterRemote = { kind: "raster"; z: number; id: string; x: number; y: number; w: number; h: number; pngBase64: string; };

type ExportMaskedImage = {
  kind: "maskedImage"; z: number; id: string;
  x: number; y: number; w: number; h: number;
  pngBytes: number[]; crop: { x: number; y: number; w: number; h: number };
};
type ExportMaskedImageRemote = {
  kind: "maskedImage"; z: number; id: string;
  x: number; y: number; w: number; h: number;
  pngBase64: string; crop: { x: number; y: number; w: number; h: number };
};

type ExportItem = ExportText | ExportShape | ExportRaster | ExportMaskedImage | ExportRasterRemote | ExportMaskedImageRemote;

type ExportSlide = {
  name: string; width: number; height: number; scale: number;
  bgPngBytes: number[]; bgShape?: { fill: string; opacity: number } | null;
  bgPngBase64?: string | null;
  fullPngBytes?: number[] | null;
  items: ExportItem[];
};

function pngToBase64(bytes: Uint8Array): string {
  return figma.base64Encode(bytes);
}

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
function getIsUppercase(tn: TextNode): boolean {
  try {
    const len = tn.characters?.length ?? 0;
    if (len === 0) return false;
    const tc = tn.getRangeTextCase(0, 1) as TextCase;
    return tc === "UPPER";
  } catch {
    const tc = (tn as any).textCase as TextCase | PluginAPI["mixed"] | undefined;
    if (!tc || tc === figma.mixed) return false;
    return tc === "UPPER";
  }
}
function lineHeightToPx(lineHeight: LineHeight | PluginAPI["mixed"] | null | undefined, fontSizePx: number): number | null {
  if (!lineHeight || lineHeight === figma.mixed) return null;
  if (lineHeight.unit === "AUTO") return null;
  if (lineHeight.unit === "PIXELS") return typeof lineHeight.value === "number" ? lineHeight.value : null;
  if (lineHeight.unit === "PERCENT") return typeof lineHeight.value === "number" ? (fontSizePx * lineHeight.value) / 100 : null;
  return null;
}

function getTextLineHeightPx(tn: TextNode, fontSizePx: number): number | null {
  try {
    return lineHeightToPx(tn.lineHeight, fontSizePx);
  } catch { return null; }
}

function getDirectSolidFillHex(tn: TextNode): string | null {
  try {
    const fills = tn.fills;
    if (!fills || fills === figma.mixed) return null;
    const solid = (fills as readonly Paint[]).find((p) => p.type === "SOLID") as SolidPaint | undefined;
    return solid ? rgbToHex(solid.color) : null;
  } catch {
    return null;
  }
}

function getDirectTextLineHeightPx(tn: TextNode, fontSizePx: number): number | null {
  try {
    return lineHeightToPx(tn.lineHeight, fontSizePx);
  } catch {
    return null;
  }
}

function getSolidFillHexFromPaints(paints: readonly Paint[] | PluginAPI["mixed"] | null | undefined): string | null {
  if (!paints || paints === figma.mixed) return null;
  const solid = paints.find((p) => p.type === "SOLID") as SolidPaint | undefined;
  return solid ? rgbToHex(solid.color) : null;
}

function getDirectTextRunsPayload(tn: TextNode): ExportTextRun[] | null {
  try {
    const segments = (tn as any).getStyledTextSegments?.(["fontName", "fontSize", "fills", "textCase", "lineHeight"]) as any[] | undefined;
    if (!segments || segments.length <= 1) return null;

    const runs: ExportTextRun[] = [];
    for (const seg of segments) {
      const segText = String(seg?.characters ?? "");
      if (!segText.length) continue;

      const fontName = seg?.fontName as FontName | PluginAPI["mixed"] | undefined;
      const fontSize = seg?.fontSize as number | PluginAPI["mixed"] | undefined;
      const textCase = seg?.textCase as TextCase | PluginAPI["mixed"] | undefined;
      const color = getSolidFillHexFromPaints(seg?.fills as readonly Paint[] | PluginAPI["mixed"] | undefined);

      if (!fontName || fontName === figma.mixed) return null;
      if (typeof fontSize !== "number") return null;
      if (!color) return null;

      const style = (fontName.style || "").toLowerCase();
      const uppercase = textCase === "UPPER";
      runs.push({
        text: uppercase ? segText.toUpperCase() : segText,
        fontFamily: fontName.family || "Arial",
        fontSize,
        lineHeightPx: lineHeightToPx(seg?.lineHeight as LineHeight | PluginAPI["mixed"] | undefined, fontSize),
        color,
        bold: style.includes("bold") || style.includes("semibold") || style.includes("demibold") || style.includes("heavy") || style.includes("black"),
        italic: style.includes("italic") || style.includes("oblique"),
        uppercase
      });
    }

    return runs.length > 1 ? runs : null;
  } catch {
    return null;
  }
}

function getDirectTextPayload(tn: TextNode): Omit<ExportText, "kind" | "z" | "id" | "x" | "y" | "w" | "h"> | null {
  try {
    const text = tn.characters ?? "";
    if (!text.length) return null;

    const align = tn.textAlignHorizontal;
    if (!align || align === figma.mixed) return null;

    const fontName = tn.fontName;
    const fontSize = tn.fontSize;
    const textCase = tn.textCase;
    const opacity = typeof tn.opacity === "number" ? tn.opacity : 1;

    if (fontName && fontName !== figma.mixed && typeof fontSize === "number" && textCase !== figma.mixed) {
      const color = getDirectSolidFillHex(tn);
      if (!color) return null;

      const style = (fontName.style || "").toLowerCase();
      return {
        text,
        fontFamily: fontName.family || "Arial",
        fontSize,
        lineHeightPx: getDirectTextLineHeightPx(tn, fontSize),
        color,
        align: alignMap(align),
        opacity,
        bold: style.includes("bold") || style.includes("semibold") || style.includes("demibold") || style.includes("heavy") || style.includes("black"),
        italic: style.includes("italic") || style.includes("oblique"),
        uppercase: textCase === "UPPER"
      };
    }

    const runs = getDirectTextRunsPayload(tn);
    if (!runs || !runs.length) return null;
    const first = runs[0];
    return {
      text,
      fontFamily: first.fontFamily,
      fontSize: first.fontSize,
      lineHeightPx: first.lineHeightPx ?? null,
      color: first.color,
      align: alignMap(align),
      opacity,
      bold: first.bold,
      italic: first.italic,
      uppercase: false,
      runs
    };
  } catch {
    return null;
  }
}

function getUltraSafeTextPayload(tn: TextNode): Omit<ExportText, "kind" | "z" | "id" | "x" | "y" | "w" | "h"> | null {
  try {
    const text = tn.characters ?? "";
    if (!text.length) return null;

    const alignRaw = tn.textAlignHorizontal;
    const align = (!alignRaw || alignRaw === figma.mixed) ? "left" : alignMap(alignRaw as TextNode["textAlignHorizontal"]);

    return {
      text,
      fontFamily: "Arial",
      fontSize: 14,
      lineHeightPx: null,
      color: "000000",
      align,
      opacity: typeof tn.opacity === "number" ? tn.opacity : 1,
      bold: false,
      italic: false,
      uppercase: false
    };
  } catch {
    return null;
  }
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
function hasBlendMode(node: SceneNode): boolean {
  if (!("blendMode" in node)) return false;
  const bm = (node as any).blendMode;
  if (!bm || bm === figma.mixed) return true;
  return bm !== "NORMAL";
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

function isComplexContainerForConservativeExport(node: SceneNode): boolean {
  return node.type === "GROUP" || node.type === "INSTANCE" || node.type === "COMPONENT" || node.type === "COMPONENT_SET";
}

function countVisibleDescendants(node: SceneNode, limit = 48): number {
  if (!("children" in node)) return 0;
  const stack: SceneNode[] = [...(node.children as readonly SceneNode[])];
  let count = 0;
  while (stack.length) {
    const n = stack.pop()!;
    if ("visible" in n && (n as any).visible === false) continue;
    count += 1;
    if (count >= limit) return count;
    if ("children" in n) stack.push(...(n.children as readonly SceneNode[]));
  }
  return count;
}

function shouldRasterizeConservativeContainer(node: SceneNode, frame: FrameNode): boolean {
  if (!isComplexContainerForConservativeExport(node)) return false;
  if (isNearFullFrame(node, frame)) return false;
  return countVisibleDescendants(node, 28) >= 28 || node.width >= frame.width * 0.2 || node.height >= frame.height * 0.12;
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
function hasOverflowingDescendant(container: SceneNode): boolean {
  if (!("children" in container)) return false;
  const cBounds = getNodeBounds(container);
  const stack: SceneNode[] = [...(container.children as readonly SceneNode[])];
  while (stack.length) {
    const n = stack.pop()!;
    if ("visible" in n && (n as any).visible === false) continue;
    const b = getNodeBounds(n);
    const outLeft = b.x < cBounds.x - 0.5;
    const outTop = b.y < cBounds.y - 0.5;
    const outRight = b.x + b.w > cBounds.x + cBounds.w + 0.5;
    const outBottom = b.y + b.h > cBounds.y + cBounds.h + 0.5;
    if (outLeft || outTop || outRight || outBottom) return true;
    if ("children" in n) stack.push(...(n.children as readonly SceneNode[]));
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
  if (hasBlendMode(node)) {
    if (isContainer(node) && containsTextDescendant(node)) return false;
    return true;
  }

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
    if (frame.clipsContent === true && hasOverflowingDescendant(frame)) return null;

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
function isMaskNode(n: SceneNode): boolean {
  return (n as any).isMask === true;
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

function isRectOutsideFrame(r: { x: number; y: number; w: number; h: number }, frame: FrameNode): boolean {
  const pad = 0.5;
  return r.x < -pad || r.y < -pad || r.x + r.w > frame.width + pad || r.y + r.h > frame.height + pad;
}

function isWidePillCandidate(node: SceneNode, frame: FrameNode): boolean {
  const radius = getCornerRadiusAny(node);
  if (radius <= 0) return false;
  const r = rectRelativeToFrame(node, frame);
  if (r.w < frame.width * 0.7) return false;
  if (r.h > frame.height * 0.22) return false;
  return radius >= r.h / 2 - 1;
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
  for (const t of texts) {
    try {
      prev.set(t.id, t.visible);
      t.visible = false;
    } catch {
      // Ignore nodes whose visibility can't be toggled in current context.
    }
  }
  try {
    return await container.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: scale } });
  } finally {
    for (const t of texts) {
      const v = prev.get(t.id);
      if (typeof v === "boolean") {
        try { t.visible = v; } catch { /* ignore */ }
      }
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

async function exportOneFrame(
  frame: FrameNode,
  idx: number,
  total: number,
  exportScale: number,
  includeFullRaster: boolean,
  conservative = false
): Promise<ExportSlide> {
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

  async function tryExtractMaskPairsInContainer(container: SceneNode): Promise<boolean> {
    if (!("children" in container)) return false;

    const ch = (container.children as readonly SceneNode[]).filter((n) => ("visible" in n ? (n as any).visible !== false : true));
    if (ch.length < 2) return false;

    const first = ch[0];
    const second = ch[1];

    if (!isMaskNode(first)) return false;

    const maskNode = first;
    const contentNode = second;

    if (isRectMaskNode(maskNode) && contentNode.type === "RECTANGLE") {
      const mask = maskNode as RectangleNode;
      const img = contentNode as RectangleNode;

      if (!isSafeMaskPair(mask, img)) return false;

      const maskR = rectRelativeToFrame(mask, frame);
      const imgR = rectRelativeToFrame(img, frame);
      const maskRadius = getCornerRadiusAny(mask);

      if (maskRadius > 0.5) {
        postProgress("export", idx - 1, total, `Mask (rounded): ${frame.name}`, `Rasterizing rounded mask…`);
        const contRect = rectRelativeToFrame(container, frame);
        const bytes = await rasterizeNodePNG(container, exportScale);

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
        return true;
      }

      postProgress("export", idx - 1, total, `Mask image: ${frame.name}`, `Exporting masked image…`);
      const imgBytes = await rasterizeNodePNG(img, exportScale);
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
      return true;
    }

    postProgress("export", idx - 1, total, `Mask group: ${frame.name}`, `Rasterizing masked content…`);
    const contRect = rectRelativeToFrame(container, frame);
    const bytes = await rasterizeNodePNG(container, exportScale);

    z += 1;
    const zVal = z;

    items.push({
      kind: "raster",
      z: zVal,
      id: `maskGroup__${container.id}`,
      x: contRect.x, y: contRect.y, w: contRect.w, h: contRect.h,
      pngBytes: Array.from(bytes)
    });

    markHide(container);
    return true;
  }

  function addTextItem(tn: TextNode) {
    const r = rectRelativeToFrame(tn, frame);
    const flags = getFirstCharFontStyleFlags(tn);
    const fs = getFirstCharFontSize(tn);

    z += 1;
    zById.set(tn.id, z);

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
      italic: flags.italic,
      uppercase: getIsUppercase(tn)
    });

    markHide(tn);
  }

  async function walk(node: SceneNode, textOnly = false) {
    if (!("visible" in node) || (node as any).visible === false) return;
    if (cancelRequested) return;

    if (textOnly && node.type !== "TEXT") {
      if ("children" in node) {
        for (const ch of node.children as readonly SceneNode[]) {
          await walk(ch as SceneNode, true);
          if (cancelRequested) return;
        }
      }
      return;
    }

    if (node.id !== frame.id && isContainer(node)) {
      const handledMask = await tryExtractMaskPairsInContainer(node);
      if (handledMask) {
        if ("children" in node) {
          for (const ch of node.children as readonly SceneNode[]) {
            await walk(ch as SceneNode, true);
            if (cancelRequested) return;
          }
        }
        return;
      }
    }

    if (node.id !== frame.id) {
      if (conservative && shouldRasterizeConservativeContainer(node, frame)) {
        const r = rectRelativeToFrame(node, frame);
        const isOverflowing = frame.clipsContent === true && isRectOutsideFrame(r, frame);
        if (!isOverflowing) {
          rasterCandidates.push(node);
          markHide(node);
        }
        return;
      }

      if (consumedMaskIds.has(node.id) || consumedMaskedContentIds.has(node.id)) return;

      z += 1;
      zById.set(node.id, z);

      // 2) Container BG as editable shape (solid)
      if (node.type === "FRAME" || node.type === "INSTANCE" || node.type === "COMPONENT") {
        if ("clipsContent" in node && (node as FrameNode).clipsContent === true && hasOverflowingDescendant(node)) {
          const r = rectRelativeToFrame(node, frame);
          postProgress("export", idx - 1, total, `Clipped frame: ${frame.name}`, `Rasterizing clipped content…`);
          const bytes = await rasterizeContainerBackgroundOnly(node, exportScale);

          items.push({
            kind: "raster",
            z,
            id: `clippedFrame__${node.id}`,
            x: r.x, y: r.y, w: r.w, h: r.h,
            pngBytes: Array.from(bytes)
          });

          markHide(node);
          if ("children" in node) {
            for (const ch of node.children as readonly SceneNode[]) {
              await walk(ch as SceneNode, true);
              if (cancelRequested) return;
            }
          }
          return;
        }

        if (isSafeContainerBg(node)) {
          const r = rectRelativeToFrame(node, frame);
          const fill = getSolidFill(node);
          const stroke = getSolidStroke(node);
          const radius = getCornerRadiusAny(node);
          const isOverflowing = frame.clipsContent === true && isRectOutsideFrame(r, frame);

          if (isWidePillCandidate(node, frame)) {
            postProgress("export", idx - 1, total, `Pill background: ${frame.name}`, `Rasterizing pill background…`);
            const bytes = await rasterizeContainerBackgroundOnly(node, exportScale);

            items.push({
              kind: "raster",
              z,
              id: `pillBg__${node.id}`,
              x: r.x, y: r.y, w: r.w, h: r.h,
              pngBytes: Array.from(bytes)
            });

            markHide(node);
            if ("children" in node) {
              for (const ch of node.children as readonly SceneNode[]) {
                await walk(ch as SceneNode, true);
                if (cancelRequested) return;
              }
            }
            return;
          }

          if (!isOverflowing) {
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
      }

      // 2.1) Gradient container BG as raster (bg-only)
      if (isGradientContainerCandidate(node)) {
        const r = rectRelativeToFrame(node, frame);
        const isAlmostFull = r.w >= frame.width * 0.9 && r.h >= frame.height * 0.9;
        if (!isAlmostFull) {
          postProgress("export", idx - 1, total, `Gradient pill: ${frame.name}`, `Rasterizing gradient background…`);
          const bytes = await rasterizeContainerBackgroundOnly(node, exportScale);

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
          const bytes = await rasterizeNodePNG(node, exportScale);

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
        addTextItem(tn);
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
          const isOverflowing = frame.clipsContent === true && isRectOutsideFrame(r, frame);

          if (fill || stroke) {
            if (isWidePillCandidate(rn, frame)) {
              postProgress("export", idx - 1, total, `Pill shape: ${frame.name}`, `Rasterizing pill shape…`);
              const bytes = await rasterizeNodePNG(rn, exportScale);

              items.push({
                kind: "raster",
                z,
                id: `pillRect__${rn.id}`,
                x: r.x, y: r.y, w: r.w, h: r.h,
                pngBytes: Array.from(bytes)
              });
              markHide(rn);
              return;
            }

            if (!isOverflowing) {
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
        const r = rectRelativeToFrame(node, frame);
        const isOverflowing = frame.clipsContent === true && isRectOutsideFrame(r, frame);
        if (!isOverflowing) {
          rasterCandidates.push(node);
          markHide(node);
        }
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
      const bytes = await rasterizeNodePNG(n, exportScale);
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
    for (const n of toHide) {
      try {
        prevVisible.set(n.id, (n as any).visible);
        (n as any).visible = false;
      } catch {
        // If visibility cannot be toggled, keep node in background.
      }
    }

    let bgPng: Uint8Array;
    const prevClips = frame.clipsContent;
    try {
      throwIfCancelled();
      if (!frame.clipsContent) frame.clipsContent = true;
      bgPng = await frame.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: exportScale } });
    } finally {
      frame.clipsContent = prevClips;
      for (const n of toHide) {
        const v = prevVisible.get(n.id);
        if (typeof v === "boolean") {
          try { (n as any).visible = v; } catch { /* ignore */ }
        }
      }
    }
    bgPngBytes = Array.from(bgPng);
  }

  let fullPngBytes: number[] | null = null;
  if (includeFullRaster) {
    const prevClips = frame.clipsContent;
    try {
      if (!frame.clipsContent) frame.clipsContent = true;
      const fullPng = await frame.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: exportScale } });
      fullPngBytes = Array.from(fullPng);
    } finally {
      frame.clipsContent = prevClips;
    }
  }

  throwIfCancelled();
  postProgress("export", idx, total, `Ready: ${frame.name}`);

  return {
    name: frame.name,
    width: frame.width,
    height: frame.height,
    scale: exportScale,
    bgPngBytes,
    bgShape,
    fullPngBytes,
    items: allItems
  };
}

async function exportFramesLocally(
  frames: FrameNode[],
  exportScale: number,
  includeFullRaster: boolean,
  filename: string,
  format: string,
  quality: string
) {
  const slides: ExportSlide[] = [];
  for (let i = 0; i < frames.length; i++) {
    throwIfCancelled();
    slides.push(await exportOneFrame(frames[i], i + 1, frames.length, exportScale, includeFullRaster, false));
  }

  throwIfCancelled();

  safeUiPostMessage({
    type: "BATCH_BG_AND_ITEMS_V051",
    filename,
    slides,
    format,
    quality
  });
  postProgress("export", frames.length, frames.length, "Sent to PPTX builder", "Building PPTX…");
}

async function exportOneFrameRemoteRasterOnly(
  frame: FrameNode,
  idx: number,
  total: number,
  exportScale: number,
  label = "Safe frame render"
): Promise<ExportSlide> {
  throwIfCancelled();
  postProgress("export", idx - 1, total, `Rasterizing: ${frame.name}`, label);
  const safeScale = total >= 30 ? Math.min(exportScale, 0.75) :
    total >= 10 ? Math.min(exportScale, 0.9) :
    Math.min(exportScale, 1);
  const bg = await frame.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: safeScale } });
  throwIfCancelled();
  postProgress("export", idx, total, `Ready: ${frame.name}`);
  return {
    name: frame.name,
    width: frame.width,
    height: frame.height,
    scale: safeScale,
    bgPngBytes: [],
    bgPngBase64: pngToBase64(bg),
    bgShape: null,
    fullPngBytes: null,
    items: []
  };
}

async function exportOneFrameRemoteSafe(
  frame: FrameNode,
  idx: number,
  total: number,
  exportScale: number
): Promise<ExportSlide> {
  throwIfCancelled();
  postProgress("export", idx - 1, total, `Scanning: ${frame.name}`, `Scanning frame ${idx}/${total}: ${frame.name}`);

  const crashSafeRasterOnly = total >= 36 || frame.children.length >= 220;
  if (crashSafeRasterOnly) {
    return await exportOneFrameRemoteRasterOnly(frame, idx, total, exportScale, "Safe frame render");
  }

  const items: ExportItem[] = [];
  const flattenForStability = total >= 20 || frame.children.length >= 140;
  const ultraStableMode = total >= 30 || frame.children.length >= 180;
  let z = 0;

  function nextZ(nodeId: string) {
    z += 1;
    return z;
  }

  function collectVisibleDescendantsSafe(root: SceneNode): SceneNode[] {
    const out: SceneNode[] = [];
    const stack: SceneNode[] = [root];
    while (stack.length) {
      const n = stack.pop()!;
      if ("visible" in n && (n as any).visible === false) continue;
      if (n.id !== root.id) out.push(n);
      if ("children" in n) {
        for (const ch of n.children as readonly SceneNode[]) stack.push(ch as SceneNode);
      }
    }
    return out;
  }

  async function addRasterItemSafe(node: SceneNode, idPrefix?: string) {
    const r = rectRelativeToFrame(node, frame);
    if (r.w <= 0 || r.h <= 0) return;
    const maxSide = Math.max(r.w, r.h);
    const nodeScale = maxSide >= 4000 ? Math.min(exportScale, 0.85) :
      maxSide >= 2500 ? Math.min(exportScale, 1) :
      exportScale;

    postProgress("export", idx - 1, total, `Rasterizing: ${frame.name}`, node.name);
    const bytes = await rasterizeNodePNG(node, nodeScale);
    const itemZ = nextZ(node.id);
    items.push({
      kind: "raster",
      z: itemZ,
      id: idPrefix ? `${idPrefix}__${node.id}` : node.id,
      x: r.x, y: r.y, w: r.w, h: r.h,
      pngBase64: pngToBase64(bytes)
    });
  }

  if (flattenForStability) {
    const nodes = collectVisibleDescendantsSafe(frame);
    const hideForBg: SceneNode[] = [];
    const maxFlattenRasterItems = total >= 20 ? 36 : 64;
    let flattenRasterItems = 0;

    function containsMaskNode(root: SceneNode): boolean {
      if (isMaskNode(root)) return true;
      if (!("children" in root)) return false;
      const stack: SceneNode[] = [...(root.children as readonly SceneNode[])];
      while (stack.length) {
        const n = stack.pop()!;
        if (isMaskNode(n)) return true;
        if ("children" in n) stack.push(...(n.children as readonly SceneNode[]));
      }
      return false;
    }

    function hasRotation(node: SceneNode): boolean {
      return !isRotationZero(node);
    }

    function isSafeStandaloneRasterNode(node: SceneNode): boolean {
      if (isNearFullFrame(node, frame)) return false;
      if (isMaskNode(node)) return false;
      if (containsMaskNode(node)) return false;
      if (hasRotation(node)) return false;
      if (hasAnyEffects(node)) return false;
      if (hasBlendMode(node)) return false;
      if (isContainer(node) && ("clipsContent" in node) && (node as FrameNode).clipsContent === true) return false;
      if (node.type === "TEXT") return false;
      if (!canToggleNodeVisibility(node)) return false;

      // Respect ancestor context: if parent chain introduces clipping/mask/effects/blends/opacity,
      // keep node merged into background for visual fidelity.
      let parent = node.parent as BaseNode | null;
      while (parent && parent.id !== frame.id) {
        if ("type" in parent) {
          const p = parent as SceneNode;
          if (p.type === "INSTANCE" || p.type === "COMPONENT" || p.type === "COMPONENT_SET") return false;
          if (isMaskNode(p)) return false;
          if (hasRotation(p)) return false;
          if (hasAnyEffects(p)) return false;
          if (hasBlendMode(p)) return false;
          if ("clipsContent" in p && (p as FrameNode).clipsContent === true) return false;
          if (typeof (p as any).opacity === "number" && Math.abs(((p as any).opacity as number) - 1) > 0.001) return false;
        }
        parent = (parent as any).parent ?? null;
      }

      // If there is real editable text inside, we keep node in background and emit text separately.
      if (containsTextDescendant(node)) return false;

      // Separate simple, non-masked visual primitives as raster layers to preserve look.
      if (node.type === "RECTANGLE") {
        return true;
      }
      if (node.type === "ELLIPSE") {
        return true;
      }
      if (
        node.type === "VECTOR" ||
        node.type === "BOOLEAN_OPERATION" ||
        node.type === "STAR" ||
        node.type === "POLYGON"
      ) {
        return true;
      }
      if (
        node.type === "GROUP" ||
        node.type === "INSTANCE" ||
        node.type === "COMPONENT" ||
        node.type === "COMPONENT_SET"
      ) {
        const r = rectRelativeToFrame(node, frame);
        const area = r.w * r.h;
        const frameArea = Math.max(1, frame.width * frame.height);
        if (area / frameArea > 0.32) return false;
        if (Math.max(r.w, r.h) > Math.max(frame.width, frame.height) * 0.6) return false;
        return true;
      }

      return false;
    }

    async function addFlattenRasterNode(node: SceneNode): Promise<boolean> {
      if (flattenRasterItems >= maxFlattenRasterItems) return false;
      const r = rectRelativeToFrame(node, frame);
      if (r.w <= 0 || r.h <= 0) return false;

      const maxSide = Math.max(r.w, r.h);
      const area = r.w * r.h;
      if (maxSide > 5000 || area > 8_000_000) return false;

      const nodeScale = maxSide >= 4000 ? Math.min(exportScale, 0.8) :
        maxSide >= 2500 ? Math.min(exportScale, 0.95) :
        Math.min(exportScale, 1.1);

      postProgress("export", idx - 1, total, `Rasterizing: ${frame.name}`, node.name || "Image");
      const bytes = await rasterizeNodePNG(node, nodeScale);

      items.push({
        kind: "raster",
        z: nextZ(node.id),
        id: `flattenRaster__${node.id}`,
        x: r.x, y: r.y, w: r.w, h: r.h,
        pngBase64: pngToBase64(bytes)
      });
      hideForBg.push(node);
      flattenRasterItems += 1;
      return true;
    }

    for (const node of nodes) {
      throwIfCancelled();

      if (node.type === "TEXT") {
        const payload = ultraStableMode
          ? (getDirectTextPayload(node) || getUltraSafeTextPayload(node))
          : (getDirectTextPayload(node) || getUltraSafeTextPayload(node));
        if (!payload) continue;
        const r = rectRelativeToFrame(node, frame);
        items.push({
          kind: "text",
          z: nextZ(node.id),
          id: node.id,
          x: r.x, y: r.y, w: r.w, h: r.h,
          ...payload
        });
        hideForBg.push(node);
        continue;
      }

      if (node.type === "RECTANGLE" && isSafeEditableRect(node)) {
        const r = rectRelativeToFrame(node, frame);
        const fill = getSolidFill(node);
        const stroke = getSolidStroke(node);
        const radius = getCornerRadiusAny(node);
        if (fill || stroke) {
          items.push({
            kind: "shape",
            z: nextZ(node.id),
            id: node.id,
            shape: "rect",
            x: r.x, y: r.y, w: r.w, h: r.h,
            fill, stroke, radius,
            opacity: typeof node.opacity === "number" ? node.opacity : 1
          });
          hideForBg.push(node);
        }
        continue;
      }

      if (node.type === "ELLIPSE" && isSafeEditableEllipse(node)) {
        const r = rectRelativeToFrame(node, frame);
        const fill = getSolidFill(node);
        const stroke = getSolidStroke(node);
        if (fill || stroke) {
          items.push({
            kind: "shape",
            z: nextZ(node.id),
            id: node.id,
            shape: "ellipse",
            x: r.x, y: r.y, w: r.w, h: r.h,
            fill, stroke, radius: 0,
            opacity: typeof node.opacity === "number" ? node.opacity : 1
          });
          hideForBg.push(node);
        }
        continue;
      }

      if (node.type === "LINE" && isSafeEditableLine(node)) {
        const r = rectRelativeToFrame(node, frame);
        const stroke = getSolidStroke(node)!;
        items.push({
          kind: "shape",
          z: nextZ(node.id),
          id: node.id,
          shape: "line",
          x: r.x, y: r.y, w: r.w, h: r.h,
          stroke,
          opacity: typeof node.opacity === "number" ? node.opacity : 1
        });
        hideForBg.push(node);
        continue;
      }

      // Keep safe raster candidates as separate pictures where feasible.
      if (isSafeStandaloneRasterNode(node)) {
        try {
          const extracted = await addFlattenRasterNode(node);
          if (extracted) continue;
        } catch {
          // Keep this node in background if separate export failed.
        }
      }
    }

    const prevVisible = new Map<string, boolean>();
    for (const n of hideForBg) {
      try {
        prevVisible.set(n.id, n.visible);
        n.visible = false;
      } catch {
        // Keep node in background if visibility cannot be toggled.
      }
    }

    let bgPngBase64: string | null = null;
    try {
      throwIfCancelled();
      postProgress("export", idx - 1, total, `Rasterizing: ${frame.name}`, "Background");
      const bgScale = ultraStableMode ? Math.min(exportScale, 1) : Math.min(exportScale, 1.15);
      const bgPng = await frame.exportAsync({ format: "PNG", constraint: { type: "SCALE", value: bgScale } });
      bgPngBase64 = pngToBase64(bgPng);
    } finally {
      for (const n of hideForBg) {
        const v = prevVisible.get(n.id);
        if (typeof v === "boolean") {
          try { n.visible = v; } catch { /* ignore */ }
        }
      }
    }

    throwIfCancelled();
    postProgress("export", idx, total, `Ready: ${frame.name}`);
    return {
      name: frame.name,
      width: frame.width,
      height: frame.height,
      scale: exportScale,
      bgPngBytes: [],
      bgPngBase64,
      bgShape: null,
      fullPngBytes: null,
      items
    };
  }

  async function handleDirectChild(node: SceneNode) {
    if (!("visible" in node) || (node as any).visible === false) return;
    throwIfCancelled();

    if (node.type === "TEXT") {
      const payload = getDirectTextPayload(node) || getUltraSafeTextPayload(node);
      if (!payload) {
        return;
      }
      const r = rectRelativeToFrame(node, frame);
      items.push({
        kind: "text",
        z: nextZ(node.id),
        id: node.id,
        x: r.x, y: r.y, w: r.w, h: r.h,
        ...payload
      });
      return;
    }

    if (node.type === "RECTANGLE" && isSafeEditableRect(node)) {
      const r = rectRelativeToFrame(node, frame);
      const fill = getSolidFill(node);
      const stroke = getSolidStroke(node);
      const radius = getCornerRadiusAny(node);
      if (fill || stroke) {
        items.push({
          kind: "shape",
          z: nextZ(node.id),
          id: node.id,
          shape: "rect",
          x: r.x, y: r.y, w: r.w, h: r.h,
          fill, stroke, radius,
          opacity: typeof node.opacity === "number" ? node.opacity : 1
        });
        return;
      }
    }

    if (node.type === "ELLIPSE" && isSafeEditableEllipse(node)) {
      const r = rectRelativeToFrame(node, frame);
      const fill = getSolidFill(node);
      const stroke = getSolidStroke(node);
      if (fill || stroke) {
        items.push({
          kind: "shape",
          z: nextZ(node.id),
          id: node.id,
          shape: "ellipse",
          x: r.x, y: r.y, w: r.w, h: r.h,
          fill, stroke, radius: 0,
          opacity: typeof node.opacity === "number" ? node.opacity : 1
        });
        return;
      }
    }

    if (node.type === "LINE" && isSafeEditableLine(node)) {
      const r = rectRelativeToFrame(node, frame);
      const stroke = getSolidStroke(node)!;
      items.push({
        kind: "shape",
        z: nextZ(node.id),
        id: node.id,
        shape: "line",
        x: r.x, y: r.y, w: r.w, h: r.h,
        stroke,
        opacity: typeof node.opacity === "number" ? node.opacity : 1
      });
      return;
    }

    await addRasterItemSafe(node, isContainer(node) ? "safeContainer" : "safeRaster");
  }

  for (const child of frame.children as readonly SceneNode[]) {
    await handleDirectChild(child as SceneNode);
  }

  let bgShape: { fill: string; opacity: number } | null = { fill: "FFFFFF", opacity: 1 };
  try {
    if (isRotationZero(frame) && !hasAnyEffects(frame) && !hasImageFill(frame) && !hasAnyGradientFill(frame) && hasOnlySolidFills(frame)) {
      const fill = getSolidFill(frame);
      if (fill) bgShape = { fill, opacity: typeof frame.opacity === "number" ? frame.opacity : 1 };
    }
  } catch {
    bgShape = { fill: "FFFFFF", opacity: 1 };
  }

  throwIfCancelled();
  postProgress("export", idx, total, `Ready: ${frame.name}`);

  return {
    name: frame.name,
    width: frame.width,
    height: frame.height,
    scale: exportScale,
    bgPngBytes: [],
    bgPngBase64: null,
    bgShape,
    fullPngBytes: null,
    items
  };
}

async function exportFramesRemotelyDirect(
  frames: FrameNode[],
  exportScale: number,
  includeFullRaster: boolean,
  filename: string
) {
  const adaptiveScale = frames.length >= 80 ? Math.min(exportScale, 1) :
    frames.length >= 40 ? Math.min(exportScale, 1.15) :
    frames.length >= 20 ? Math.min(exportScale, 1.35) :
    exportScale;

  postStatus(`Using Railway backend for PPTX export… (${activeRemoteApiBase})`);
  postProgress("remote", 0, frames.length, "Creating server job…", "Connecting to Railway…");

  const created = await fetchRemoteJson<{ jobId: string }>(`/api/jobs`, {
    method: "POST",
    body: JSON.stringify({ filename, expectedSlides: frames.length })
  });

  const { jobId } = created;

  for (let i = 0; i < frames.length; i++) {
    throwIfCancelled();
    const frame = frames[i];
    let slide: ExportSlide;
    try {
      slide = await exportOneFrameRemoteSafe(
        frame,
        i + 1,
        frames.length,
        adaptiveScale
      );
    } catch (err: any) {
      if (err?.__cancelled || err?.message === "CANCELLED") throw err;
      postStatus(`Slide ${i + 1}/${frames.length} fallback to safe raster…`);
      slide = await exportOneFrameRemoteRasterOnly(
        frame,
        i + 1,
        frames.length,
        adaptiveScale,
        "Fallback raster render"
      );
    }

    throwIfCancelled();
    postStatus(`Mode: Railway backend. Uploading slide ${i + 1}/${frames.length}…`);
    postProgress("upload", i, frames.length, `Uploading slide ${i + 1}/${frames.length}`, slide.name);

    await fetchRemoteJson(`/api/jobs/${jobId}/slides/${i}`, {
      method: "PUT",
      body: JSON.stringify(slide)
    });

    postProgress("upload", i + 1, frames.length, `Uploaded ${i + 1}/${frames.length}`, slide.name);
  }

  throwIfCancelled();
  postStatus("Mode: Railway backend. Starting server build…");
  postProgress("remote", frames.length, frames.length, "Starting server build…", "Building PPTX on Railway…");

  await fetchRemoteJson(`/api/jobs/${jobId}/finalize`, {
    method: "POST",
    body: JSON.stringify({})
  });

  for (let attempt = 0; attempt < 240; attempt++) {
    throwIfCancelled();

    const status = await fetchRemoteJson<{ status: string; error?: string | null; downloadUrl?: string | null }>(
      `/api/jobs/${jobId}/status`
    );

    if (status.status === "completed" && status.downloadUrl) {
      postStatus("Mode: Railway backend. Downloading PPTX…");
      postProgress("done", 1, 1, "Done", "Server export complete ✅");
      safeUiPostMessage({
        type: "REMOTE_EXPORT_READY",
        filename,
        downloadUrl: status.downloadUrl
      });
      return;
    }

    if (status.status === "failed") {
      throw new Error(status.error || "Railway export failed.");
    }

    postStatus(`Mode: Railway backend. Server status: ${status.status}`);
    postProgress("remote", frames.length, frames.length, "Building on Railway…", `Server status: ${status.status}`);
    await sleep(1500);
  }

  throw new Error("Remote export timed out while waiting for the server.");
}

// --- Messages ---
figma.ui.onmessage = async (msg) => {
  try {
    if (msg.type === "REQUEST_SELECTION") { await refreshSelectionFramesSafely(); return; }

    if (msg.type === "CANCEL_EXPORT") {
      cancelRequested = true;
      postStatus("Cancel requested…");
      return;
    }

    if (msg.type === "EXPORT_PPTX_ORDERED") {
      cancelRequested = false;
      exportInProgress = true;
      try {

        const ids: string[] = Array.isArray(msg.frameIds) ? msg.frameIds : [];
        if (!ids.length) { postError("No frames in export list."); return; }

        const quality = String(msg.quality || "best");
        const format = String(msg.format || "pptx");
        const useRemotePptx = shouldUseRemotePptx(format) && ids.length >= REMOTE_PPTX_MIN_FRAMES;
        const exportScale = getExportScale(format, quality, useRemotePptx);
        const includeFullRaster = format === "pdf";

        const nodes = await Promise.all(ids.map((id) => figma.getNodeByIdAsync(id)));
        const frames: FrameNode[] = nodes.filter((n): n is FrameNode => !!n && (n as any).type === "FRAME");

        if (!frames.length) { postError("Selected frames not found. Click Refresh and try again."); return; }

        const filename = frames.length === 1 ? `${frames[0].name}.pptx` : `Lucy_batch_${frames.length}_slides.pptx`;

        postProgress("export", 0, frames.length, "Starting export…", `Exporting ${frames.length} frame(s)…`);

        if (useRemotePptx) {
          await exportFramesRemotelyDirect(frames, exportScale, includeFullRaster, filename);
          return;
        }

        if (format === "pptx" && shouldUseRemotePptx(format) && frames.length < REMOTE_PPTX_MIN_FRAMES) {
          postStatus(`Using local high-fidelity export for ${frames.length} slides…`);
        }
        await exportFramesLocally(frames, exportScale, includeFullRaster, filename, format, quality);
        return;
      } finally {
        exportInProgress = false;
      }
    }
  } catch (e: any) {
    if (e?.__cancelled || e?.message === "CANCELLED") { postCancelled(); return; }
    postError(e?.message ?? String(e));
  }
};
