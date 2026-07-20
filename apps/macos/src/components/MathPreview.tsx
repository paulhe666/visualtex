import { useMemo } from "react";
import { convertLatexToMarkup } from "mathlive";

interface MathPreviewProps {
  latex: string;
  className?: string;
}

export function MathPreview({ latex, className = "" }: MathPreviewProps) {
  const markup = useMemo(
    () => convertLatexToMarkup(latex, { defaultMode: "math" }),
    [latex],
  );

  return (
    <span
      className={"math-preview " + className}
      aria-hidden="true"
      dangerouslySetInnerHTML={{ __html: markup }}
    />
  );
}
