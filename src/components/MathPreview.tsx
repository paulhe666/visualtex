import { useEffect, useRef } from "react";
import { MathfieldElement } from "mathlive";

interface MathPreviewProps {
  latex: string;
  className?: string;
}

export function MathPreview({ latex, className = "" }: MathPreviewProps) {
  const hostRef = useRef<HTMLSpanElement>(null);
  const fieldRef = useRef<MathfieldElement | null>(null);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const field = new MathfieldElement();
    field.value = latex;
    field.readOnly = true;
    field.setAttribute("math-virtual-keyboard-policy", "manual");
    field.tabIndex = -1;
    field.className = "math-preview-field";
    host.replaceChildren(field);
    fieldRef.current = field;
    return () => {
      fieldRef.current = null;
      host.replaceChildren();
    };
  }, []);

  useEffect(() => {
    if (fieldRef.current && fieldRef.current.value !== latex) {
      fieldRef.current.value = latex;
    }
  }, [latex]);

  return <span ref={hostRef} className={"math-preview " + className} aria-hidden="true" />;
}
