// Every foreground string on the landing page. Each is drawn as brush lettering at build time,
// so editing copy means re-running `npm run build:landing-art`.
// `display: true` uses the display script (Norican) for English; `enTex` renders the English side
// with MathJax (Computer Modern) instead, for LaTeX-style card labels; `tex` renders both sides
// with MathJax (for strings that read the same in both languages).
export const copy = {
  navEditor: { zh: "编辑器", en: "Editor" },
  navFeatures: { zh: "功能", en: "Features" },
  navDownload: { zh: "下载", en: "Download" },
  langSwitch: { zh: "EN", en: "中文" },
  refProp: { tex: String.raw`\text{(2)}` },
  refThm: { tex: String.raw`\text{(3)}` },

  heroOpen: { zh: "在浏览器里打开", en: "Open in the browser" },
  heroDownload: { zh: "下载桌面版", en: "Get the desktop app" },
  noteHero: { zh: "无需安装，打开就能写", en: "no install, just start writing" },

  say1: { zh: "所见即所得", en: "What you see is what you write.", display: true },
  say2: { zh: "不必再看嵌套格式", en: "No more counting braces.", display: true },
  say3: { zh: "写一次，到处都能用", en: "Write it once, use it anywhere.", display: true },

  figLabel: { zh: "图 1 · 网页编辑器", enTex: String.raw`\textbf{Figure 1}\ \ \text{The web editor.}` },
  openEditor: { zh: "打开编辑器", en: "Open the editor" },

  propLabel: { zh: "命题 2 · 它能做什么", enTex: String.raw`\textbf{Proposition 2}\ \ \text{(What it does).}` },
  feature1: { zh: "可视化编辑，LaTeX 源码实时同步", en: "Visual editing, with the LaTeX source in sync" },
  feature2: { zh: "拍下手写或截图，识别成公式", en: "Photograph handwriting or a screenshot, get the formula" },
  feature3: { zh: "一键插入 Word 与 PowerPoint", en: "Insert into Word and PowerPoint in one step" },
  feature4: { zh: "原生 MathType 公式，无需安装 MathType", en: "Native MathType equations, no MathType install" },
  feature5: { zh: "导出 LaTeX、SVG 与 PNG", en: "Export LaTeX, SVG and PNG" },

  thmLabel: { zh: "定理 3 · 下载", enTex: String.raw`\textbf{Theorem 3}\ \ \text{(Download).}` },
  macName: { zh: "macOS", en: "macOS" },
  winName: { zh: "Windows", en: "Windows" },
  macMeta: { zh: "Apple Silicon · macOS 11+ · v{VERSION}", en: "Apple Silicon · macOS 11+ · v{VERSION}" },
  winMeta: { zh: "Windows 10 / 11 · x64 · v{VERSION}", en: "Windows 10 / 11 · x64 · v{VERSION}" },
  thisDevice: { zh: "当前设备", en: "This device" },
  full: { zh: "完整版", en: "Full edition" },
  lite: { zh: "轻量版 · 无本地 OCR", en: "Lite · no local OCR" },
  mobileNote: { zh: "桌面安装包请在电脑上下载。", en: "Download the desktop app on a computer." },
  ocrCaption: { zh: "离线 OCR 模型 · Windows", en: "Offline OCR · Windows" },
  ocrS: { zh: "OCR-S  200 MB", en: "OCR-S  200 MB" },
  ocrM: { zh: "OCR-M  426 MB", en: "OCR-M  426 MB" },
  ocrL: { zh: "OCR-L  670 MB", en: "OCR-L  670 MB" },
  allReleases: { zh: "全部版本", en: "All releases" },

  remLabel: { zh: "注 4 · 支持", enTex: String.raw`\textbf{Remark 4}\ \ \text{(Support).}` },
  qrWechat: { zh: "微信打赏", en: "WeChat Pay" },
  qrAlipay: { zh: "支付宝打赏", en: "Alipay" },
  qrGroup: { zh: "QQ 交流群", en: "QQ group" },
  qrFailed: { zh: "二维码暂时未能加载。", en: "The QR codes could not be loaded." },
  tips: { zh: "自愿打赏，不影响任何功能。", en: "Tips are optional and never unlock features." },
  group: { zh: "QQ 群 1045801770", en: "QQ group 1045801770" },

  footerName: { zh: "VisualTeX", en: "VisualTeX" },
  footerGithub: { zh: "GitHub", en: "GitHub" },
};
