import { compileCustomSymbolDesignerArtwork } from "./compiler.ts";
import type { CustomSymbolDesignerDocument } from "./designerTypes.ts";
import { shapeSvgMarkup } from "./shapes.ts";
import type { CustomSymbolBox } from "./symbolBox.ts";

const n = (value: number) => {
  const rounded = Math.round(value * 1000) / 1000;
  return (Object.is(rounded, -0) ? 0 : rounded).toString();
};

/**
 * The designed symbol as a standalone SVG (1000 units per em, sized in em so
 * it scales with the surrounding text, its baseline on the text's baseline):
 * the registered box, widened to where overlays reach outside it; eraser
 * strokes cut out with a mask.
 */
export function customSymbolSvgMarkup(
  document: CustomSymbolDesignerDocument,
  box: CustomSymbolBox | null,
  paint = "currentColor",
): string {
  if (!box) return "";
  const width = box.metrics.widthEm * 1000;
  const height = (box.metrics.ascentEm + box.metrics.descentEm) * 1000;
  // the drawing's extent from the box's top left
  const extent = box.artworkBounds ?? { x: 0, y: 0, width, height };
  const shapes = compileCustomSymbolDesignerArtwork(document);
  const defs: string[] = [];
  const painted: string[] = [];
  const erased: string[] = [];
  shapes.forEach((shape, index) => {
    if (shape.operation === "erase") {
      erased.push(shapeSvgMarkup({ ...shape, operation: undefined }, "black", `vtx-cs-clip-e${index}`, defs));
    } else {
      painted.push(shapeSvgMarkup(shape, paint, `vtx-cs-clip-${index}`, defs));
    }
  });
  let content = painted.join("");
  if (erased.length) {
    const x = box.left + extent.x - extent.width;
    const y = box.top + extent.y - extent.height;
    defs.push(
      `<mask id="vtx-cs-erase" maskUnits="userSpaceOnUse" x="${n(x)}" y="${n(y)}" width="${n(extent.width * 3)}" height="${n(extent.height * 3)}">` +
        `<rect x="${n(x)}" y="${n(y)}" width="${n(extent.width * 3)}" height="${n(extent.height * 3)}" fill="white"/>${erased.join("")}</mask>`,
    );
    content = `<g mask="url(#vtx-cs-erase)">${content}</g>`;
  }
  // depth below the baseline of the drawing (not only the box)
  const below = (extent.y + extent.height) / 1000 - box.metrics.ascentEm;
  const name = document.command.trim().replace(/^\\/, "");
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${n(extent.x)} ${n(extent.y)} ${n(extent.width)} ${n(extent.height)}" ` +
      `width="${n(extent.width / 1000)}em" height="${n(extent.height / 1000)}em" style="vertical-align:${n(-below)}em"` +
      `${name ? ` data-latex="\\${name}"` : ""}>`,
    defs.length ? `<defs>${defs.join("")}</defs>` : "",
    `<g transform="translate(${n(-box.left)} ${n(-box.top)})">${content}</g>`,
    "</svg>",
  ]
    .filter(Boolean)
    .join("\n");
}
