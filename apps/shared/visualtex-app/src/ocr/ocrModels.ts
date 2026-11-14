// The local formula recognition models (PaddleOCR PP-FormulaNet+), the same in both apps; the backends download or
// bundle them.

export const OCR_MODELS = [
  {
    id: "PP-FormulaNet_plus-S",
    labelZh: "高速版 S",
    labelEn: "Fast S",
    hintZh: "248 MB · 英文公式",
    hintEn: "248 MB · English formulas",
    downloadMb: 259.6,
    storageMb: 248,
    cpuBenchmarkMs: 260.99,
  },
  {
    id: "PP-FormulaNet_plus-M",
    labelZh: "均衡版 M（推荐）",
    labelEn: "Balanced M (recommended)",
    hintZh: "592 MB · 中文与复杂公式",
    hintEn: "592 MB · Chinese and complex formulas",
    downloadMb: 620.5,
    storageMb: 592,
    cpuBenchmarkMs: 1615.8,
  },
  {
    id: "PP-FormulaNet_plus-L",
    labelZh: "高精度版 L",
    labelEn: "High accuracy L",
    hintZh: "698 MB · 内存占用数 GB",
    hintEn: "698 MB · uses several GB of memory",
    downloadMb: 731.5,
    storageMb: 698,
    cpuBenchmarkMs: 3125.58,
  },
] as const;

export type OcrModelName = (typeof OCR_MODELS)[number]["id"];
export const DEFAULT_OCR_MODEL: OcrModelName = "PP-FormulaNet_plus-M";
