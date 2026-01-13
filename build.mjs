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

// 2) Copy UI CSS -> dist/ui.css
fs.copyFileSync("src/ui.css", "dist/ui.css");

// 3) Copy UI HTML -> dist/ui.html (for debugging / local inspection)
fs.copyFileSync("src/ui.html", "dist/ui.html");

// 4) Inline CSS + JS into HTML string for figma.showUI(__html__)
let uiHtml = fs.readFileSync("src/ui.html", "utf8");
const uiCss = fs.readFileSync("dist/ui.css", "utf8");
const uiJs  = fs.readFileSync("dist/ui.js", "utf8");

// If the HTML already has a <link href="ui.css">, replace it; otherwise inject CSS before </head>
if (uiHtml.includes('href="ui.css"')) {
  uiHtml = uiHtml.replace(/<link[^>]*href="ui\.css"[^>]*>/g, `<style>\n${uiCss}\n</style>`);
} else {
  uiHtml = uiHtml.replace("</head>", `<style>\n${uiCss}\n</style>\n</head>`);
}

// Replace <script src="ui.js"></script> with inlined JS
uiHtml = uiHtml.replace(/<script[^>]*src="ui\.js"[^>]*><\/script>/g, `<script>\n${uiJs}\n</script>`);

// Write inlined UI for reference
fs.writeFileSync("dist/ui.inlined.html", uiHtml, "utf8");

// 5) Build plugin code -> dist/code.js with __html__ injected
await build({
  entryPoints: ["src/code.ts"],
  bundle: true,
  minify: true,
  platform: "browser",
  target: ["es2017"],
  format: "iife",
  outfile: "dist/code.js",
  define: {
    __html__: JSON.stringify(uiHtml)
  }
});

console.log("Build complete: dist/code.js + dist/ui.inlined.html (used for __html__)");
