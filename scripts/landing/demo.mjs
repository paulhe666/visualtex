// Figure 1 on the landing page: a typing demo of the visual editor. Each step is one thing the
// visitor would type (`key`), the formula as the editor shows it (`view`, with a caret and
// placeholders) and the LaTeX source the editor writes for it (`src`).
// \C = caret, \P = empty placeholder; both are replaced before MathJax sees them.
const R = String.raw;

export const demos = [
  [
    { key: "int", view: R`\int\C`, src: R`\int` },
    { key: "_", view: R`\int_{\C}`, src: R`\int_{}` },
    { key: "-inf", view: R`\int_{-\infty\C}`, src: R`\int_{-\infty}` },
    { key: "^", view: R`\int_{-\infty}^{\C}`, src: R`\int_{-\infty}^{}` },
    { key: "inf", view: R`\int_{-\infty}^{\infty\C}`, src: R`\int_{-\infty}^{\infty}` },
    { key: "→ e^", view: R`\int_{-\infty}^{\infty}e^{\C}`, src: R`\int_{-\infty}^{\infty}e^{}` },
    { key: "-x^2", view: R`\int_{-\infty}^{\infty}e^{-x^{2\C}}`, src: R`\int_{-\infty}^{\infty}e^{-x^{2}}` },
    { key: "→→ dx", view: R`\int_{-\infty}^{\infty}e^{-x^{2}}\,dx\C`, src: R`\int_{-\infty}^{\infty}e^{-x^{2}}\,dx` },
    { key: "=sqrt", view: R`\int_{-\infty}^{\infty}e^{-x^{2}}\,dx=\sqrt{\C}`, src: R`\int_{-\infty}^{\infty}e^{-x^{2}}\,dx=\sqrt{}` },
    { key: "pi", view: R`\int_{-\infty}^{\infty}e^{-x^{2}}\,dx=\sqrt{\pi\C}`, src: R`\int_{-\infty}^{\infty}e^{-x^{2}}\,dx=\sqrt{\pi}` },
  ],
  [
    { key: "sum", view: R`\sum_{\C}^{\P}`, src: R`\sum_{}^{}` },
    { key: "n=1", view: R`\sum_{n=1\C}^{\P}`, src: R`\sum_{n=1}^{}` },
    { key: "Tab inf", view: R`\sum_{n=1}^{\infty\C}`, src: R`\sum_{n=1}^{\infty}` },
    { key: "→ 1/", view: R`\sum_{n=1}^{\infty}\frac{1}{\C}`, src: R`\sum_{n=1}^{\infty}\frac{1}{}` },
    { key: "n^2", view: R`\sum_{n=1}^{\infty}\frac{1}{n^{2\C}}`, src: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}` },
    { key: "→→ =pi^2", view: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}=\pi^{2\C}`, src: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}=\pi^{2}` },
    { key: "→ /", view: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}=\frac{\pi^{2}}{\C}`, src: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}=\frac{\pi^{2}}{}` },
    { key: "6", view: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}=\frac{\pi^{2}}{6\C}`, src: R`\sum_{n=1}^{\infty}\frac{1}{n^{2}}=\frac{\pi^{2}}{6}` },
  ],
  [
    { key: "pmatrix", view: R`\begin{pmatrix}\C&\P\\\P&\P\end{pmatrix}`, src: R`\begin{pmatrix}&\\&\end{pmatrix}` },
    { key: "cos t", view: R`\begin{pmatrix}\cos t\C&\P\\\P&\P\end{pmatrix}`, src: R`\begin{pmatrix}\cos t&\\&\end{pmatrix}` },
    { key: "Tab -sin t", view: R`\begin{pmatrix}\cos t&-\sin t\C\\\P&\P\end{pmatrix}`, src: R`\begin{pmatrix}\cos t&-\sin t\\&\end{pmatrix}` },
    { key: "Tab sin t", view: R`\begin{pmatrix}\cos t&-\sin t\\\sin t\C&\P\end{pmatrix}`, src: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\end{pmatrix}` },
    { key: "Tab cos t", view: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\cos t\C\end{pmatrix}`, src: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\cos t\end{pmatrix}` },
    { key: "→ =e^", view: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\cos t\end{pmatrix}=e^{\C}`, src: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\cos t\end{pmatrix}=e^{}` },
    { key: "tJ", view: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\cos t\end{pmatrix}=e^{tJ\C}`, src: R`\begin{pmatrix}\cos t&-\sin t\\\sin t&\cos t\end{pmatrix}=e^{tJ}` },
  ],
];

/** Keystrokes a step costs: named keys (Tab, arrows) count once, everything else per character. */
export function keystrokes(key) {
  return key.split(" ").reduce((sum, word) => sum + (word === "Tab" ? 1 : [...word].length), 0);
}
