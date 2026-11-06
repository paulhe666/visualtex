// Document mode's tutorials (toolbar button and main menu): the tasks on the left, the chosen one's screen recording
// and what it shows on the right. The videos are the app's own files (public/tutorials/<zh|en>/<id>.mp4), read
// into a blob first: a media element asks for byte ranges, which the app's asset protocol does not answer everywhere.
import { useEffect, useMemo, useRef, useState } from "react";
import { ChevronLeft, ChevronRight, GraduationCap, X } from "lucide-react";
import { TUTORIAL_GROUPS, type Tutorial } from "./tutorials";

const blobs = new Map<string, Promise<string>>();

function videoBlob(path: string): Promise<string> {
  let url = blobs.get(path);
  if (!url) {
    url = fetch(path)
      .then((response) => {
        if (!response.ok) throw new Error(`${response.status}`);
        return response.blob();
      })
      .then((blob) => URL.createObjectURL(blob));
    url.catch(() => blobs.delete(path));
    blobs.set(path, url);
  }
  return url;
}

const ALL: Tutorial[] = TUTORIAL_GROUPS.flatMap((group) => group.tutorials);

export function TutorialDialog(props: { open: boolean; isEn: boolean; onClose: () => void }) {
  const { open, isEn, onClose } = props;
  const [current, setCurrent] = useState(ALL[0].id);
  const [src, setSrc] = useState<string | null>(null);
  const [failed, setFailed] = useState(false);
  const videoRef = useRef<HTMLVideoElement>(null);
  const lang = isEn ? "en" : "zh";
  const pick = (pair: [string, string]) => (isEn ? pair[1] : pair[0]);
  const index = ALL.findIndex((tutorial) => tutorial.id === current);
  const tutorial = ALL[index];
  const base = `${import.meta.env.BASE_URL}tutorials/${lang}/`;
  const poster = useMemo(() => `${base}${current}.jpg`, [base, current]);

  useEffect(() => {
    if (!open) return;
    let alive = true;
    setSrc(null);
    setFailed(false);
    videoBlob(`${base}${current}.mp4`).then(
      (url) => alive && setSrc(url),
      () => alive && setFailed(true),
    );
    // the next one too, so stepping through does not wait
    const next = ALL[index + 1];
    if (next) void videoBlob(`${base}${next.id}.mp4`).catch(() => undefined);
    return () => {
      alive = false;
    };
  }, [open, base, current, index]);

  useEffect(() => {
    if (!open) return;
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
      } else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        const step = event.key === "ArrowDown" ? 1 : -1;
        setCurrent((id) => {
          const at = ALL.findIndex((item) => item.id === id);
          return ALL[Math.min(ALL.length - 1, Math.max(0, at + step))].id;
        });
      }
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [open, onClose]);

  if (!open || !tutorial) return null;
  const go = (step: number) => setCurrent(ALL[Math.min(ALL.length - 1, Math.max(0, index + step))].id);

  return (
    <div
      className="modal-backdrop help-dialog-backdrop"
      role="presentation"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <section className="dialog help-dialog vtx-tutorial-dialog" role="dialog" aria-modal="true" aria-labelledby="vtx-tutorial-title">
        <header className="help-dialog-header">
          <div>
            <strong id="vtx-tutorial-title">
              <GraduationCap size={18} />
              {isEn ? "Document tutorials" : "文档教程"}
            </strong>
          </div>
          <button type="button" className="icon-button compact" onClick={onClose} aria-label={isEn ? "Close" : "关闭"} title={isEn ? "Close (Esc)" : "关闭 (Esc)"}>
            <X size={17} />
          </button>
        </header>
        <div className="help-dialog-body vtx-tutorial-body">
          <nav className="help-dialog-toc vtx-tutorial-list" aria-label={isEn ? "Tutorials" : "教程"}>
            {TUTORIAL_GROUPS.map((group) => (
              <div key={group.title[1]} className="vtx-tutorial-group">
                <div className="vtx-tutorial-group-title">{pick(group.title)}</div>
                {group.tutorials.map((item) => (
                  <button
                    type="button"
                    key={item.id}
                    className={item.id === current ? "is-current" : undefined}
                    aria-current={item.id === current ? "true" : undefined}
                    onClick={() => setCurrent(item.id)}
                  >
                    {pick(item.title)}
                  </button>
                ))}
              </div>
            ))}
          </nav>
          <article className="vtx-tutorial-main">
            <div className="vtx-tutorial-video">
              {failed ? (
                <div className="vtx-tutorial-missing">{isEn ? "This video could not be loaded." : "无法加载这段视频。"}</div>
              ) : (
                <video
                  key={current}
                  ref={videoRef}
                  src={src ?? undefined}
                  poster={poster}
                  autoPlay
                  muted
                  playsInline
                  controls
                  preload="auto"
                />
              )}
            </div>
            <div className="vtx-tutorial-text">
              <h3>{pick(tutorial.title)}</h3>
              <p>{pick(tutorial.summary)}</p>
            </div>
            <footer className="vtx-tutorial-steps">
              <button type="button" className="secondary-button" disabled={index <= 0} onClick={() => go(-1)}>
                <ChevronLeft size={15} />
                {isEn ? "Previous" : "上一个"}
              </button>
              <span>
                {index + 1} / {ALL.length}
              </span>
              <button type="button" className="secondary-button" disabled={index >= ALL.length - 1} onClick={() => go(1)}>
                {isEn ? "Next" : "下一个"}
                <ChevronRight size={15} />
              </button>
            </footer>
          </article>
        </div>
      </section>
    </div>
  );
}
