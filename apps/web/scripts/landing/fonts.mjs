// Fonts used only at build time to draw brush lettering as SVG outlines. They are fetched from
// the google/fonts repository (all SIL OFL 1.1) into node_modules/.cache and never shipped.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import opentype from "opentype.js";

const CACHE = new URL("../../node_modules/.cache/visualtex-landing-fonts/", import.meta.url);
const SOURCE = "https://raw.githubusercontent.com/google/fonts/main/ofl";

export const FONT_FILES = {
  word: "greatvibes/GreatVibes-Regular.ttf", // hero "VisualTeX"
  display: "norican/Norican-Regular.ttf", // English slogans
  body: "caveatbrush/CaveatBrush-Regular.ttf", // English body text
  han: "mashanzheng/MaShanZheng-Regular.ttf", // Chinese
};

export async function loadFonts() {
  mkdirSync(CACHE, { recursive: true });
  const fonts = {};
  for (const [key, path] of Object.entries(FONT_FILES)) {
    const file = new URL(path.split("/")[1], CACHE);
    if (!existsSync(file)) {
      const response = await fetch(`${SOURCE}/${path}`);
      if (!response.ok) throw new Error(`Could not download ${path}: HTTP ${response.status}`);
      writeFileSync(file, Buffer.from(await response.arrayBuffer()));
    }
    fonts[key] = opentype.parse(readFileSync(file).buffer);
  }
  return fonts;
}

// opentype.js's toPathData can drop the separator between numbers ("713 0" -> "7130"),
// which silently truncates paths in browsers. Serialize commands ourselves.
export function pathD(path, digits = 1) {
  const n = (v) => +v.toFixed(digits);
  return path.commands.map((c) => {
    switch (c.type) {
      case "M":
      case "L": return `${c.type}${n(c.x)} ${n(c.y)}`;
      case "Q": return `Q${n(c.x1)} ${n(c.y1)} ${n(c.x)} ${n(c.y)}`;
      case "C": return `C${n(c.x1)} ${n(c.y1)} ${n(c.x2)} ${n(c.y2)} ${n(c.x)} ${n(c.y)}`;
      default: return "Z";
    }
  }).join("");
}
