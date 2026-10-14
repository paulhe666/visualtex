import { MathfieldElement } from "mathlive";

export function normalizeMathfieldLatexForRegression(latex: string) {
  const field = new MathfieldElement();
  field.style.position = "fixed";
  field.style.left = "-10000px";
  field.style.top = "-10000px";
  document.body.append(field);
  try {
    field.setValue(latex, {
      mode: "math",
      format: "latex",
      insertionMode: "replaceAll",
      selectionMode: "after",
      silenceNotifications: true,
    });
    return field.value;
  } finally {
    field.remove();
  }
}
