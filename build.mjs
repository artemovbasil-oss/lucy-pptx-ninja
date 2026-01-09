import { build } from "esbuild";
import fs from "node:fs";
import path from "node:path";

const distDir = path.resolve("dist");
fs.mkdirSync(distDir, { recursive: true });

// 1) Build UI JS -> dist/ui.js
await build({
  entryPoints: ["src/ui.ts"],
  bundle: true,
  minify: true,
  platform: "browser",
  target: ["es2017"],
  format: "iife",
  outfile: "dist/ui.js"
});

// 2) Copy UI css -> dist/ui.css
fs.copyFileSync("src/ui.css", "dist/ui.css");

// 3) Read UI html template
const uiHtmlTemplate = fs.readFileSync("src/ui.html", "utf8");

// 4) Inline CSS + JS into HTML (critical for figma.showUI(__html__))
const uiCss = fs.readFileSync("dist/ui.css", "utf8");
const uiJs = fs.readFileSync("dist/ui.js", "utf8");

let uiHtmlInlined = uiHtmlTemplate;

// Replace <link rel="stylesheet" href="ui.css" /> with <style>...</style>
uiHtmlInlined = uiHtmlInlined.replace(
  /<link\s+rel=["']stylesheet["']\s+href=["']ui\.css["']\s*\/?>/i,
  `<style>\n${uiCss}\n</style>`
);

// Replace <script src="ui.js"></script> with inline script
uiHtmlInlined = uiHtmlInlined.replace(
  /<script\s+src=["']ui\.js["']\s*><\/script>/i,
  `<script>\n${uiJs}\n</script>`
);

// Write final ui.html that will be embedded as __html__
fs.writeFileSync("dist/ui.html", uiHtmlInlined, "utf8");

// 5) Bundle code.ts -> dist/code.js with __html__ inlined
await build({
  entryPoints: ["src/code.ts"],
  bundle: true,
  minify: true,
  platform: "browser",
  target: ["es2017"],
  format: "iife",
  outfile: "dist/code.js",
  define: {
    __html__: JSON.stringify(uiHtmlInlined)
  }
});

console.log("Build complete: dist/code.js + dist/ui.html (CSS/JS inlined)");
