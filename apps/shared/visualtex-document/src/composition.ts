import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/**
 * ProseMirror's view.composing is set by a compositionstart in the text and cleared only by a compositionend it
 * handles. A composition that ends where it does not see it (the end went to an open formula's editor, whose events
 * are its own; the focus left the text with the input method's text still marked; WebKit sent no compositionend)
 * left the flag set for the rest of the session: the pages were never measured again (margins, page breaks and
 * numbers all gone), the keymaps did nothing (keys are not dispatched while composing). Here a composition the text no
 * longer has is ended: the compositionend ProseMirror missed is sent to it.
 */
export function compositionGuardPlugin(): Plugin {
  return new Plugin({
    key: new PluginKey("visualtex-document-composition-guard"),
    view(view) {
      let timer = 0;
      const end = () => {
        window.clearTimeout(timer);
        if (!view.isDestroyed && view.composing) view.dom.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true, data: "" }));
      };
      const later = (delay: number, stale: () => boolean) => {
        window.clearTimeout(timer);
        timer = window.setTimeout(() => {
          if (!view.isDestroyed && view.composing && stale()) end();
        }, delay);
      };
      // a compositionend from inside a block whose events are its own (ProseMirror skips it): the text's composition, if
      // the focus is not in the text any more, is over too
      const onEnd = (event: Event) => {
        if (event.target !== view.dom && view.composing) later(0, () => !composingInText(view));
      };
      // the focus went out of the text (to a formula's editor, another panel, another window)
      const onFocusOut = () => {
        if (view.composing) later(50, () => !composingInText(view));
      };
      // a key that is no input method's (not 229, not marked composing) while composing: the composition is long over
      const onKeyDown = (event: KeyboardEvent) => {
        if (view.composing && !event.isComposing && event.keyCode !== 229 && event.key !== "Process" && event.key !== "Dead") end();
      };
      view.dom.addEventListener("compositionend", onEnd, true);
      view.dom.addEventListener("focusout", onFocusOut);
      view.dom.addEventListener("keydown", onKeyDown, true);
      return {
        destroy() {
          window.clearTimeout(timer);
          view.dom.removeEventListener("compositionend", onEnd, true);
          view.dom.removeEventListener("focusout", onFocusOut);
          view.dom.removeEventListener("keydown", onKeyDown, true);
        },
      };
    },
  });
}

/** An input method can be composing in the text: the focus and the selection are in it (not in a block's own editor). */
export function composingInText(view: EditorView): boolean {
  return view.composing && (view.root as Document | ShadowRoot).activeElement === view.dom;
}
