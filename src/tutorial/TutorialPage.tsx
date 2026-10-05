import { Fragment, useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { ArrowLeft, ArrowRight, Check, Pause, Play, RotateCcw } from "lucide-react";
import { TUTORIAL_LESSONS, type TutorialLang, type TutorialLesson, type TutorialSnapshot } from "./lessons";

const LANG_KEY = "visualtex.tutorial.lang";
const PROGRESS_KEY = "visualtex.tutorial.progress.v1";
const CHEATSHEET_ID = "keys";

// Phones and tablets get the demos only; the editor needs a keyboard.
const PRACTICE_QUERY = "(max-width: 760px), (pointer: coarse)";

const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad|iPod/i.test(navigator.platform || navigator.userAgent);

const KEY_LABELS: Record<string, string> = isMac
  ? { Ctrl: "Control", Alt: "Option", Shift: "Shift", Enter: "Return", Tab: "Tab", Esc: "Esc", Mod: "⌘" }
  : { Ctrl: "Ctrl", Alt: "Alt", Shift: "Shift", Enter: "Enter", Tab: "Tab", Esc: "Esc", Mod: "Ctrl" };

const UI = {
  zh: {
    title: "VisualTeX 新手教程",
    openEditor: "打开编辑器",
    lessons: "课程",
    cheatsheet: "快捷键速查",
    watch: "演示",
    replay: "重播",
    pause: "暂停",
    play: "播放",
    practice: "动手试试",
    restart: "重来",
    practiceHint: "在下面的编辑器里完成：",
    allDone: "这一课完成了。",
    next: "下一课",
    prev: "上一课",
    phoneNote: "在电脑上打开这一页，可以在真实编辑器里练习。",
    sandboxNote: "练习区和你自己的文档互不影响。",
    keysTitle: "快捷键速查",
    more: "更多功能见编辑器菜单里的「帮助手册」。",
    langSwitch: "English",
  },
  en: {
    title: "VisualTeX Tutorial",
    openEditor: "Open the editor",
    lessons: "Lessons",
    cheatsheet: "Shortcuts",
    watch: "Demo",
    replay: "Replay",
    pause: "Pause",
    play: "Play",
    practice: "Try it",
    restart: "Start over",
    practiceHint: "Do this in the editor below:",
    allDone: "Lesson complete.",
    next: "Next lesson",
    prev: "Previous",
    phoneNote: "Open this page on a computer to practise in the real editor.",
    sandboxNote: "The practice editor never touches your own document.",
    keysTitle: "Shortcuts",
    more: "Everything else is in Help Manual in the editor menu.",
    langSwitch: "中文",
  },
} satisfies Record<TutorialLang, Record<string, string>>;

// Shortcuts the tutorial teaches, all checked against the editor's key handling.
const CHEATSHEET: { keys: string[]; zh: string; en: string }[] = [
  { keys: ["Enter"], zh: "新建一行（同类型）", en: "New row (same type)" },
  { keys: ["Ctrl", "Enter"], zh: "新建行内公式", en: "New inline row" },
  { keys: ["Alt", "Enter"], zh: "新建行间公式", en: "New display row" },
  { keys: ["Shift", "Enter"], zh: "在同一条公式里换行", en: "Line break inside a formula" },
  { keys: ["&"], zh: "对齐点（多行格式为 align 时）", en: "Alignment point (Multi-line set to align)" },
  { keys: ["Tab"], zh: "跳到下一个空位", en: "Next placeholder" },
  { keys: ["→"], zh: "离开上标、下标", en: "Leave a superscript or subscript" },
  { keys: ["\\"], zh: "输入命令，弹出候选", en: "Start a command and show suggestions" },
  { keys: ["Mod", "Z"], zh: "撤销", en: "Undo" },
  { keys: isMac ? ["Mod", "Shift", "Z"] : ["Mod", "Y"], zh: "重做", en: "Redo" },
];

interface TimelineEvent {
  t: number;
  dur: number;
  step: number;
  kind: "keys" | "type" | "click";
  keys?: string[];
  text?: string;
  label?: { zh: string; en: string };
}

interface Timeline {
  duration: number;
  width: number;
  height: number;
  events: TimelineEvent[];
}

type Progress = Record<string, number[]>;

function readStorage(key: string) {
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function writeStorage(key: string, value: string) {
  try {
    window.localStorage.setItem(key, value);
  } catch {
    // Progress simply won't persist.
  }
}

function detectLang(): TutorialLang {
  const saved = readStorage(LANG_KEY);
  if (saved === "zh" || saved === "en") return saved;
  const preferred = navigator.languages?.length ? navigator.languages : [navigator.language];
  return preferred[0]?.toLowerCase().startsWith("zh") ? "zh" : "en";
}

function readProgress(): Progress {
  try {
    const parsed = JSON.parse(readStorage(PROGRESS_KEY) ?? "{}");
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function Keys({ keys }: { keys: string[] }) {
  return (
    <span className="tut-keys">
      {keys.map((key, index) => (
        <Fragment key={index}>
          {index > 0 && <span className="tut-plus">+</span>}
          <kbd>{KEY_LABELS[key] ?? key}</kbd>
        </Fragment>
      ))}
    </span>
  );
}

/** [Ctrl+Enter] → key caps, `code` → code. */
function RichText({ text }: { text: string }) {
  const parts = text.split(/(`[^`]+`|\[[^\]]+\])/g);
  return (
    <>
      {parts.map((part, index) => {
        if (part.startsWith("`") && part.endsWith("`") && part.length > 1) return <code key={index}>{part.slice(1, -1)}</code>;
        if (part.startsWith("[") && part.endsWith("]") && part.length > 2) return <Keys key={index} keys={part.slice(1, -1).split("+")} />;
        return <Fragment key={index}>{part}</Fragment>;
      })}
    </>
  );
}

function eventChip(event: TimelineEvent, lang: TutorialLang): ReactNode {
  if (event.kind === "keys" && event.keys) return <Keys keys={event.keys} />;
  if (event.kind === "type" && event.text) return <code>{event.text}</code>;
  if (event.kind === "click" && event.label) return <span className="tut-chip-click">{event.label[lang]}</span>;
  return null;
}

function DemoPlayer({ lesson, lang, onStep }: { lesson: TutorialLesson; lang: TutorialLang; onStep: (step: number) => void }) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const [timeline, setTimeline] = useState<Timeline | null>(null);
  const [time, setTime] = useState(0);
  const [playing, setPlaying] = useState(true);
  const base = `/tutorial/${lesson.id}.${lang}`;

  useEffect(() => {
    let cancelled = false;
    setTimeline(null);
    setTime(0);
    fetch(`${base}.json`)
      .then((response) => (response.ok ? response.json() : null))
      .then((data: Timeline | null) => {
        if (!cancelled) setTimeline(data);
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [base]);

  useEffect(() => {
    const video = videoRef.current;
    if (!video) return;
    let frame = 0;
    const tick = () => {
      setTime(video.currentTime);
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(frame);
  }, [base]);

  const active = timeline?.events.filter((event) => time >= event.t && time < event.t + event.dur).at(-1) ?? null;
  const current = timeline?.events.filter((event) => time >= event.t).at(-1) ?? null;
  const step = current?.step ?? -1;
  useEffect(() => onStep(step), [step, onStep]);

  const toggle = () => {
    const video = videoRef.current;
    if (!video) return;
    if (video.paused) void video.play();
    else video.pause();
  };
  const replay = () => {
    const video = videoRef.current;
    if (!video) return;
    video.currentTime = 0;
    void video.play();
  };
  const seek = (event: React.PointerEvent<HTMLDivElement>) => {
    const video = videoRef.current;
    if (!video || !video.duration) return;
    const rect = event.currentTarget.getBoundingClientRect();
    video.currentTime = Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width)) * video.duration;
  };
  const duration = timeline?.duration || videoRef.current?.duration || 1;

  return (
    <figure className="tut-demo">
      <div className="tut-video" style={timeline ? { aspectRatio: `${timeline.width} / ${timeline.height}` } : undefined}>
        <video
          key={base}
          ref={videoRef}
          src={`${base}.mp4`}
          poster={`${base}.jpg`}
          muted
          loop
          autoPlay
          playsInline
          preload="auto"
          onPlay={() => setPlaying(true)}
          onPause={() => setPlaying(false)}
          onClick={toggle}
        />
        <div className={`tut-chip${active ? " is-on" : ""}`} aria-hidden="true">
          {active ? eventChip(active, lang) : null}
        </div>
      </div>
      <div className="tut-demo-bar">
        <button type="button" className="tut-icon-btn" onClick={toggle} aria-label={playing ? UI[lang].pause : UI[lang].play}>
          {playing ? <Pause size={15} /> : <Play size={15} />}
        </button>
        <div className="tut-progress" onPointerDown={seek} role="presentation">
          <div className="tut-progress-fill" style={{ width: `${Math.min(100, (time / duration) * 100)}%` }} />
          {timeline?.events
            .filter((event, index, all) => index === 0 || all[index - 1].step !== event.step)
            .map((event) => (
              <span key={event.t} className="tut-progress-mark" style={{ left: `${(event.t / duration) * 100}%` }} />
            ))}
        </div>
        <button type="button" className="tut-icon-btn" onClick={replay} aria-label={UI[lang].replay}>
          <RotateCcw size={15} />
        </button>
      </div>
    </figure>
  );
}

function Practice({
  lesson,
  lang,
  done,
  onTaskDone,
}: {
  lesson: TutorialLesson;
  lang: TutorialLang;
  done: number[];
  onTaskDone: (index: number) => void;
}) {
  const frameRef = useRef<HTMLIFrameElement>(null);
  const [session, setSession] = useState(0);
  const [canPractise, setCanPractise] = useState(() => !window.matchMedia(PRACTICE_QUERY).matches);

  useEffect(() => {
    const query = window.matchMedia(PRACTICE_QUERY);
    const update = () => setCanPractise(!query.matches);
    query.addEventListener("change", update);
    return () => query.removeEventListener("change", update);
  }, []);
  const doneRef = useRef(done);
  doneRef.current = done;

  useEffect(() => {
    const timer = window.setInterval(() => {
      const host = frameRef.current?.contentWindow as (Window & { visualtexTutorial?: { snapshot: () => TutorialSnapshot } }) | null;
      let snapshot: TutorialSnapshot | undefined;
      try {
        snapshot = host?.visualtexTutorial?.snapshot();
      } catch {
        return;
      }
      if (!snapshot) return;
      lesson.tasks.forEach((task, index) => {
        if (!doneRef.current.includes(index) && task.done(snapshot)) onTaskDone(index);
      });
    }, 400);
    return () => window.clearInterval(timer);
  }, [lesson, onTaskDone]);

  const complete = lesson.tasks.every((_, index) => done.includes(index));

  return (
    <section className="tut-practice" aria-labelledby="tut-practice-title">
      <div className="tut-practice-head">
        <h3 id="tut-practice-title">{UI[lang].practice}</h3>
        <ul className="tut-tasks">
          {lesson.tasks.map((task, index) => (
            <li key={index} className={done.includes(index) ? "is-done" : undefined}>
              <span className="tut-check" aria-hidden="true">{done.includes(index) ? <Check size={12} strokeWidth={3} /> : null}</span>
              {task.label[lang]}
            </li>
          ))}
        </ul>
        <button type="button" className="tut-text-btn" onClick={() => setSession((value) => value + 1)}>
          <RotateCcw size={14} />
          {UI[lang].restart}
        </button>
      </div>
      {canPractise ? (
        <>
          {complete && <p className="tut-complete">{UI[lang].allDone}</p>}
          <iframe
            key={`${lesson.id}-${lang}-${session}`}
            ref={frameRef}
            className="tut-frame"
            src={`/editor?tutorial=${lesson.id}&lang=${lang === "en" ? "en" : "cn"}`}
            title={`${lesson.title[lang]} · ${UI[lang].practice}`}
            allow="clipboard-write"
          />
          <p className="tut-footnote">{UI[lang].sandboxNote}</p>
        </>
      ) : (
        <p className="tut-phone-note">{UI[lang].phoneNote}</p>
      )}
    </section>
  );
}

export function TutorialPage() {
  const [lang, setLang] = useState<TutorialLang>(detectLang);
  const [progress, setProgress] = useState<Progress>(readProgress);
  const [currentId, setCurrentId] = useState(() => {
    const hash = window.location.hash.slice(1);
    return hash === CHEATSHEET_ID || TUTORIAL_LESSONS.some((lesson) => lesson.id === hash) ? hash : TUTORIAL_LESSONS[0].id;
  });
  const [step, setStep] = useState(-1);
  const t = UI[lang];
  const index = TUTORIAL_LESSONS.findIndex((lesson) => lesson.id === currentId);
  const lesson = index >= 0 ? TUTORIAL_LESSONS[index] : null;

  useEffect(() => {
    document.documentElement.lang = lang === "zh" ? "zh-CN" : "en";
    document.title = t.title;
  }, [lang, t.title]);

  useEffect(() => {
    const onHash = () => {
      const hash = window.location.hash.slice(1);
      if (hash === CHEATSHEET_ID || TUTORIAL_LESSONS.some((item) => item.id === hash)) setCurrentId(hash);
    };
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);

  const go = (id: string) => {
    window.history.replaceState(null, "", `#${id}`);
    setCurrentId(id);
    setStep(-1);
    window.scrollTo({ top: 0 });
  };

  const markDone = useCallback((lessonId: string, task: number) => {
    setProgress((previous) => {
      const list = previous[lessonId] ?? [];
      if (list.includes(task)) return previous;
      const next = { ...previous, [lessonId]: [...list, task] };
      writeStorage(PROGRESS_KEY, JSON.stringify(next));
      return next;
    });
  }, []);
  const onTaskDone = useCallback((task: number) => lesson && markDone(lesson.id, task), [lesson, markDone]);

  const isComplete = (item: TutorialLesson) => item.tasks.every((_, task) => progress[item.id]?.includes(task));
  const toggleLang = () => {
    const next = lang === "zh" ? "en" : "zh";
    writeStorage(LANG_KEY, next);
    setLang(next);
  };

  const nav = useMemo(
    () => [...TUTORIAL_LESSONS.map((item) => ({ id: item.id, label: item.title[lang], item })), { id: CHEATSHEET_ID, label: t.cheatsheet, item: null }],
    [lang, t.cheatsheet],
  );

  return (
    <div className="tut-page">
      <header className="tut-top">
        <a className="tut-brand" href="/">
          <img src="/favicon.svg" alt="" width={22} height={22} />
          <span>VisualTeX</span>
        </a>
        <span className="tut-top-title">{lang === "zh" ? "新手教程" : "Tutorial"}</span>
        <div className="tut-top-actions">
          <button type="button" className="tut-text-btn" onClick={toggleLang}>{t.langSwitch}</button>
          <a className="tut-primary" href="/editor">{t.openEditor}</a>
        </div>
      </header>

      <div className="tut-layout">
        <nav className="tut-nav" aria-label={t.lessons}>
          <ol>
            {nav.map((entry, position) => (
              <li key={entry.id}>
                <a
                  href={`#${entry.id}`}
                  className={entry.id === currentId ? "is-current" : undefined}
                  aria-current={entry.id === currentId ? "page" : undefined}
                  onClick={(event) => {
                    event.preventDefault();
                    go(entry.id);
                  }}
                >
                  <span className={`tut-nav-num${entry.item && isComplete(entry.item) ? " is-done" : ""}`}>
                    {entry.item && isComplete(entry.item) ? <Check size={12} strokeWidth={3} /> : entry.item ? position + 1 : "⌨"}
                  </span>
                  {entry.label}
                </a>
              </li>
            ))}
          </ol>
        </nav>

        <main className="tut-main">
          {lesson ? (
            <article key={lesson.id}>
              <p className="tut-kicker">{`${index + 1} / ${TUTORIAL_LESSONS.length}`}</p>
              <h1>{lesson.title[lang]}</h1>
              <p className="tut-summary"><RichText text={lesson.summary[lang]} /></p>

              <div className="tut-learn">
                <DemoPlayer lesson={lesson} lang={lang} onStep={setStep} />
                <ol className="tut-steps">
                  {lesson.steps.map((item, position) => (
                    <li key={position} className={position === step ? "is-active" : undefined}>
                      <RichText text={item[lang]} />
                    </li>
                  ))}
                </ol>
              </div>

              <Practice lesson={lesson} lang={lang} done={progress[lesson.id] ?? []} onTaskDone={onTaskDone} />

              <footer className="tut-pager">
                {index > 0 ? (
                  <button type="button" className="tut-text-btn" onClick={() => go(TUTORIAL_LESSONS[index - 1].id)}>
                    <ArrowLeft size={15} />
                    {t.prev}
                  </button>
                ) : <span />}
                <button
                  type="button"
                  className={isComplete(lesson) ? "tut-primary" : "tut-secondary"}
                  onClick={() => go(TUTORIAL_LESSONS[index + 1]?.id ?? CHEATSHEET_ID)}
                >
                  {TUTORIAL_LESSONS[index + 1] ? t.next : t.cheatsheet}
                  <ArrowRight size={15} />
                </button>
              </footer>
            </article>
          ) : (
            <article>
              <h1>{t.keysTitle}</h1>
              <table className="tut-cheatsheet">
                <tbody>
                  {CHEATSHEET.map((row) => (
                    <tr key={row.zh}>
                      <td><Keys keys={row.keys} /></td>
                      <td>{row[lang]}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <p className="tut-footnote">{t.more}</p>
              <footer className="tut-pager">
                <button type="button" className="tut-text-btn" onClick={() => go(TUTORIAL_LESSONS[TUTORIAL_LESSONS.length - 1].id)}>
                  <ArrowLeft size={15} />
                  {t.prev}
                </button>
                <a className="tut-primary" href="/editor">
                  {t.openEditor}
                  <ArrowRight size={15} />
                </a>
              </footer>
            </article>
          )}
        </main>
      </div>
    </div>
  );
}
