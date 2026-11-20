// Document mode's tutorial (the toolbar's tutorial button, the main menu, and by itself the first time document mode
// opens): a few cards, one task each, most with a short looping animation of the page made from the app itself
// (tutorial/<zh|en>/<id>.webp). An animation is read into a blob and shown from a new object URL each time its card
// comes up, so it starts from its first frame.
import { useEffect, useRef, useState } from "react";
import { FileText, X } from "lucide-react";
import { TUTORIAL_CARDS, type TutorialCard } from "@platform/tutorials";

const ANIMATIONS = import.meta.glob("./tutorial/*/*.webp", { eager: true, query: "?url", import: "default" }) as Record<string, string>;

const blobs = new Map<string, Promise<Blob>>();

function animationBlob(url: string): Promise<Blob> {
  let blob = blobs.get(url);
  if (!blob) {
    blob = fetch(url).then((response) => {
      if (!response.ok) throw new Error(`${response.status}`);
      return response.blob();
    });
    blob.catch(() => blobs.delete(url));
    blobs.set(url, blob);
  }
  return blob;
}

function animationUrl(card: TutorialCard | undefined, lang: "zh" | "en"): string | null {
  return card?.animation ? (ANIMATIONS[`./tutorial/${lang}/${card.id}.webp`] ?? null) : null;
}

/** A card's text: what is written between backticks is drawn as a key. */
function CardText({ text }: { text: string }) {
  return <p>{text.split("`").map((part, i) => (i % 2 ? <kbd key={i}>{part}</kbd> : part))}</p>;
}

export function TutorialDialog(props: { open: boolean; isEn: boolean; onClose: () => void }) {
  const { open, isEn, onClose } = props;
  const [index, setIndex] = useState(0);
  const [src, setSrc] = useState<string | null>(null);
  const [failed, setFailed] = useState(false);
  const nextRef = useRef<HTMLButtonElement>(null);
  const lang = isEn ? "en" : "zh";
  const pick = (pair: [string, string]) => (isEn ? pair[1] : pair[0]);
  const count = TUTORIAL_CARDS.length;
  const card = TUTORIAL_CARDS[Math.min(index, count - 1)];
  const last = index >= count - 1;

  // each opening starts at the first card
  useEffect(() => {
    if (open) setIndex(0);
  }, [open]);

  useEffect(() => {
    setSrc(null);
    setFailed(false);
    const url = open ? animationUrl(card, lang) : null;
    if (!url) return;
    let alive = true;
    let objectUrl: string | null = null;
    animationBlob(url).then(
      (blob) => {
        if (!alive) return;
        objectUrl = URL.createObjectURL(blob);
        setSrc(objectUrl);
      },
      () => alive && setFailed(true),
    );
    // the next card's as well, so stepping on does not wait
    const nextUrl = animationUrl(TUTORIAL_CARDS[index + 1], lang);
    if (nextUrl) void animationBlob(nextUrl).catch(() => undefined);
    return () => {
      alive = false;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [open, card, index, lang]);

  useEffect(() => {
    if (!open) return;
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        onClose();
      } else if (event.key === "ArrowRight" || event.key === "ArrowLeft") {
        event.preventDefault();
        event.stopPropagation();
        const step = event.key === "ArrowRight" ? 1 : -1;
        setIndex((at) => Math.min(count - 1, Math.max(0, at + step)));
      }
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [open, onClose, count]);

  // Enter goes on (or closes on the last card)
  useEffect(() => {
    if (open) nextRef.current?.focus({ preventScroll: true });
  }, [open, index]);

  if (!open || !card) return null;
  const hasAnimation = card.animation && !failed;

  return (
    <div
      className="modal-backdrop help-dialog-backdrop"
      role="presentation"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <section className="vtx-tutorial-dialog" role="dialog" aria-modal="true" aria-labelledby="vtx-tutorial-title">
        <button
          type="button"
          className="icon-button compact vtx-tutorial-close"
          onClick={onClose}
          aria-label={isEn ? "Close" : "关闭"}
          title={isEn ? "Close (Esc)" : "关闭 (Esc)"}
        >
          <X size={17} />
        </button>
        <div className="vtx-tutorial-stage">
          {hasAnimation ? (
            src && <img key={src} src={src} alt="" draggable={false} />
          ) : (
            <div className="vtx-tutorial-file" aria-hidden="true">
              <FileText size={46} strokeWidth={1.3} />
              {card.path && <code>{pick(card.path)}</code>}
            </div>
          )}
        </div>
        <div className="vtx-tutorial-copy">
          <h3 id="vtx-tutorial-title">{pick(card.title)}</h3>
          <CardText text={pick(card.text)} />
        </div>
        <footer className="vtx-tutorial-footer">
          <div className="vtx-tutorial-dots" role="tablist" aria-label={isEn ? "Cards" : "卡片"}>
            {TUTORIAL_CARDS.map((item, at) => (
              <button
                type="button"
                key={item.id}
                role="tab"
                className={at === index ? "is-current" : undefined}
                aria-selected={at === index}
                aria-label={`${pick(item.title)} (${at + 1}/${count})`}
                title={pick(item.title)}
                onClick={() => setIndex(at)}
              />
            ))}
          </div>
          <div className="vtx-tutorial-actions">
            {index > 0 && (
              <button type="button" className="secondary-button" onClick={() => setIndex(index - 1)}>
                {isEn ? "Back" : "上一步"}
              </button>
            )}
            <button
              ref={nextRef}
              type="button"
              className="primary-button"
              onClick={() => (last ? onClose() : setIndex(index + 1))}
            >
              {last ? (isEn ? "Start writing" : "开始写作") : isEn ? "Next" : "下一步"}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
