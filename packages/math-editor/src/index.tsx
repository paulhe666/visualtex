import { useEffect, useRef } from "react";
import { MathfieldElement } from "mathlive";
import "mathlive/fonts.css";
import "mathlive/static.css";

export interface MathNodeEditorProps {
  value: string;
  disabled?: boolean;
  autoFocus?: boolean;
  onChange: (latex: string) => void;
  onCommit?: (latex: string) => void;
}

export function MathNodeEditor({
  value,
  disabled = false,
  autoFocus = false,
  onChange,
  onCommit,
}: MathNodeEditorProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const fieldRef = useRef<MathfieldElement | null>(null);
  const suppressRef = useRef(false);
  const callbacksRef = useRef({ onChange, onCommit });
  callbacksRef.current = { onChange, onCommit };

  useEffect(() => {
    if (!hostRef.current) return;
    const field = new MathfieldElement();
    field.value = value;
    field.smartFence = true;
    field.smartMode = true;
    field.mathVirtualKeyboardPolicy = "manual";
    field.className = "vt-math-field";
    field.readOnly = disabled;

    const handleInput = () => {
      if (!suppressRef.current) callbacksRef.current.onChange(field.value);
    };
    const handleChange = () => callbacksRef.current.onCommit?.(field.value);
    field.addEventListener("input", handleInput);
    field.addEventListener("change", handleChange);
    hostRef.current.appendChild(field);
    fieldRef.current = field;
    if (autoFocus && !disabled) queueMicrotask(() => field.focus());

    return () => {
      field.removeEventListener("input", handleInput);
      field.removeEventListener("change", handleChange);
      field.remove();
      fieldRef.current = null;
    };
  }, []);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field || field.value === value) return;
    suppressRef.current = true;
    field.value = value;
    suppressRef.current = false;
  }, [value]);

  useEffect(() => {
    if (fieldRef.current) fieldRef.current.readOnly = disabled;
  }, [disabled]);

  return <div ref={hostRef} className="vt-math-editor" aria-label="LaTeX formula editor" />;
}
