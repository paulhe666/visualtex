import { useEffect, useRef, useState } from "react";
import { ART_DEMO, ART_MARKS, ART_TEXT, type DemoStep } from "./art.generated";
import type { LandingLang } from "./i18n";

const TYPE_MS = 34; // per source character
const STEP_MS = 620; // pause after each step
const HOLD_MS = 2600; // pause on a finished formula

type Frame = { demo: number; step: number; typed: number };

const reduceMotion = () => window.matchMedia("(prefers-reduced-motion: reduce)").matches;
const last = (demo: number) => ART_DEMO[demo].length - 1;

/** A number drawn with MathJax digits, so counters match the rest of the sheet. */
function Count({ value }: { value: number }) {
  return (
    <span className="landing-count" aria-label={String(value)}>
      {[...String(value)].map((digit, index) => (
        <span key={index} aria-hidden="true" dangerouslySetInnerHTML={{ __html: ART_MARKS.digits[Number(digit)] }} />
      ))}
    </span>
  );
}

/** The source line, with what the current step added in ink and every brace marked. */
function Source({ prev, next, typed }: { prev: string; next: string; typed: number }) {
  // Steps only ever append or fill in; find the common prefix and type the rest.
  let common = 0;
  while (common < prev.length && common < next.length && prev[common] === next[common]) common++;
  const text = next.slice(0, common + typed);
  return (
    <code className="landing-src">
      {[...text].map((ch, index) => {
        const fresh = index >= common;
        const brace = ch === "{" || ch === "}";
        return <span key={index} className={(fresh ? "is-new" : "") + (brace ? " is-brace" : "") || undefined}>{ch}</span>;
      })}
      <span className="landing-src-caret" aria-hidden="true" />
    </code>
  );
}

/**
 * Figure 1: the visual editor types a formula a keystroke group at a time while the LaTeX
 * source, braces and all, writes itself underneath. Runs only while on screen.
 */
export function TypingDemo({ lang }: { lang: LandingLang }) {
  const root = useRef<HTMLDivElement>(null);
  const [frame, setFrame] = useState<Frame>(() =>
    reduceMotion() ? { demo: 0, step: last(0), typed: Infinity } : { demo: 0, step: 0, typed: 0 },
  );
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const el = root.current;
    if (!el) return;
    const observer = new IntersectionObserver(([entry]) => setVisible(entry.isIntersecting), { threshold: 0.25 });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    if (!visible || reduceMotion()) return;
    const steps = ART_DEMO[frame.demo];
    const step = steps[frame.step];
    const prev = frame.step > 0 ? steps[frame.step - 1].src : "";
    let common = 0;
    while (common < prev.length && common < step.src.length && prev[common] === step.src[common]) common++;
    const remaining = step.src.length - common - frame.typed;
    let delay: number, next: Frame;
    if (remaining > 0) {
      delay = TYPE_MS;
      next = { ...frame, typed: frame.typed + 1 };
    } else if (frame.step < steps.length - 1) {
      delay = STEP_MS;
      next = { ...frame, step: frame.step + 1, typed: 0 };
    } else {
      delay = HOLD_MS;
      next = { demo: (frame.demo + 1) % ART_DEMO.length, step: 0, typed: 0 };
    }
    const timer = window.setTimeout(() => setFrame(next), delay);
    return () => window.clearTimeout(timer);
  }, [frame, visible]);

  const steps = ART_DEMO[frame.demo];
  const step: DemoStep = steps[frame.step];
  const prevSrc = frame.step > 0 ? steps[frame.step - 1].src : "";
  const keys = steps.slice(0, frame.step + 1).reduce((sum, s) => sum + s.keys, 0);
  const braces = [...step.src].filter((ch) => ch === "{" || ch === "}").length;
  const ink = (key: keyof typeof ART_TEXT) => <span className="landing-ink" dangerouslySetInnerHTML={{ __html: ART_TEXT[key][lang] }} />;

  return (
    <div className="landing-demo" ref={root} aria-hidden="true">
      <div className="landing-demo-view">
        <span className="landing-demo-cap">{ink("figView")}</span>
        <div className="landing-demo-math" key={`${frame.demo}-${frame.step}`} dangerouslySetInnerHTML={{ __html: step.view }} />
        <kbd className="landing-key" key={`k${frame.demo}-${frame.step}`}>{step.key}</kbd>
      </div>
      <div className="landing-demo-source">
        <span className="landing-demo-cap">{ink("figSource")}</span>
        <Source prev={prevSrc} next={step.src} typed={frame.typed} />
      </div>
      <div className="landing-demo-stats">
        <span>{ink("figKeys")}<Count value={keys} /></span>
        <span>{ink("figBraces")}<Count value={braces} /></span>
        <span className="is-zero">{ink("figTyped")}<Count value={0} /></span>
      </div>
    </div>
  );
}
