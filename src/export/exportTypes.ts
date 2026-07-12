export interface SvgExportOptions {
  displayMode: boolean;
  fontSizePt: number;
  paddingPx: number;
  background: "transparent" | "white";
}

export interface SvgExportResult {
  svg: string;
  base64: string;
  width: number;
  height: number;
  baseline?: number;
}

export interface PngExportOptions {
  scale?: number;
  background?: "transparent" | "white";
}

export interface PngExportResult {
  blob: Blob;
  base64: string;
  width: number;
  height: number;
}
