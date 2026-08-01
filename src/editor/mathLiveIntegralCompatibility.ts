import { convertLatexToMarkup, type MathfieldElement } from "mathlive";
import {
  OIINT_SIZE1_OVAL_HEIGHT_EM,
  OIINT_SIZE1_OVAL_PATH,
  OIINT_SIZE1_OVAL_WIDTH_EM,
  OIINT_SIZE2_OVAL_HEIGHT_EM,
  OIINT_SIZE2_OVAL_PATH,
  OIINT_SIZE2_OVAL_WIDTH_EM,
  OIIINT_SIZE1_OVAL_HEIGHT_EM,
  OIIINT_SIZE1_OVAL_PATH,
  OIIINT_SIZE1_OVAL_WIDTH_EM,
  OIIINT_SIZE2_OVAL_HEIGHT_EM,
  OIIINT_SIZE2_OVAL_PATH,
  OIIINT_SIZE2_OVAL_WIDTH_EM,
} from "../math/integralGlyphs.ts";

const globalStyleId = "visualtex-mathlive-contour-integral-style";
const shadowStyleId = "visualtex-mathlive-contour-integral-shadow-style";

interface OvalGeometry {
  path: string;
  width: number;
  height: number;
  shift: number;
}

function maskImage({ path, width, height }: OvalGeometry) {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" ` +
    `viewBox="0 0 ${width * 1000} ${height * 1000}">` +
    `<path fill="black" d="${path}"/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}

function ovalRule(
  className: "visualtex-oiint" | "visualtex-oiiint",
  sizeClass: "ML__small-op" | "ML__large-op",
  geometry: OvalGeometry,
) {
  const mask = maskImage(geometry);
  return `
    .ML__op-symbol.${sizeClass}.${className}::after {
      width: ${geometry.width}em;
      height: ${geometry.height}em;
      top: calc(50% + ${geometry.shift}em);
      -webkit-mask-image: ${mask};
      mask-image: ${mask};
    }
  `;
}

const oiintSmall: OvalGeometry = {
  path: OIINT_SIZE1_OVAL_PATH,
  width: OIINT_SIZE1_OVAL_WIDTH_EM,
  height: OIINT_SIZE1_OVAL_HEIGHT_EM,
  shift: 0,
};
const oiintLarge: OvalGeometry = {
  path: OIINT_SIZE2_OVAL_PATH,
  width: OIINT_SIZE2_OVAL_WIDTH_EM,
  height: OIINT_SIZE2_OVAL_HEIGHT_EM,
  shift: 0.08,
};
const oiiintSmall: OvalGeometry = {
  path: OIIINT_SIZE1_OVAL_PATH,
  width: OIIINT_SIZE1_OVAL_WIDTH_EM,
  height: OIIINT_SIZE1_OVAL_HEIGHT_EM,
  shift: 0,
};
const oiiintLarge: OvalGeometry = {
  path: OIIINT_SIZE2_OVAL_PATH,
  width: OIIINT_SIZE2_OVAL_WIDTH_EM,
  height: OIIINT_SIZE2_OVAL_HEIGHT_EM,
  shift: 0.08,
};

export const MATHLIVE_CONTOUR_INTEGRAL_STYLES = `
  .ML__op-symbol.visualtex-oiint,
  .ML__op-symbol.visualtex-oiiint {
    position: relative;
  }

  .ML__op-symbol.visualtex-oiint::after,
  .ML__op-symbol.visualtex-oiiint::after {
    content: "";
    position: absolute;
    left: 0;
    z-index: 1;
    display: block;
    pointer-events: none;
    background-color: currentColor;
    transform: translateY(-50%);
    -webkit-mask-position: 0 0;
    mask-position: 0 0;
    -webkit-mask-repeat: no-repeat;
    mask-repeat: no-repeat;
    -webkit-mask-size: 100% 100%;
    mask-size: 100% 100%;
  }

  ${ovalRule("visualtex-oiint", "ML__small-op", oiintSmall)}
  ${ovalRule("visualtex-oiint", "ML__large-op", oiintLarge)}
  ${ovalRule("visualtex-oiiint", "ML__small-op", oiiintSmall)}
  ${ovalRule("visualtex-oiiint", "ML__large-op", oiiintLarge)}
`;

export function installMathLiveContourIntegralGlobalStyle() {
  if (typeof document === "undefined") return;
  if (document.getElementById(globalStyleId)) return;
  const style = document.createElement("style");
  style.id = globalStyleId;
  style.textContent = MATHLIVE_CONTOUR_INTEGRAL_STYLES;
  document.head.append(style);
}

export function installMathLiveContourIntegralShadowStyle(
  field: MathfieldElement,
) {
  const shadowRoot = field.shadowRoot;
  if (!shadowRoot || shadowRoot.getElementById(shadowStyleId)) return;
  const style = document.createElement("style");
  style.id = shadowStyleId;
  style.textContent = MATHLIVE_CONTOUR_INTEGRAL_STYLES;
  shadowRoot.append(style);
}

export function convertVisualTexLatexToMarkup(
  ...args: Parameters<typeof convertLatexToMarkup>
) {
  installMathLiveContourIntegralGlobalStyle();
  return convertLatexToMarkup(...args);
}

installMathLiveContourIntegralGlobalStyle();
