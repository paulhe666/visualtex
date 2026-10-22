import { injectStylesheet, releaseStylesheet } from '../common/stylesheet';

import { parseLatex, Atom, Context } from '../core/core';

import { getCaretPoint } from '../editor-mathfield/utils';
import type { _Mathfield } from '../editor-mathfield/mathfield-private';

import { getKeybindingsForCommand } from './keybindings';

import { complete } from '../editor-mathfield/autocomplete';
import {
  getSharedElement,
  releaseSharedElement,
} from '../common/shared-element';
import { requestUpdate } from '../editor-mathfield/render';
import { visualTexLayoutEngine } from '../editor-mathfield/visualtex-hooks';
import {
  visualTexCandidateNumberMarkup,
  visualTexCompletionInfo,
} from '../latex-commands/definitions-utils';

// VisualTeX: one suggestion list for all math fields, owned by the field it
// was last shown for (`visualTexMathfield`); it is hidden rather than removed,
// so the editor can keep its own panel in step with it.
type SuggestionPanel = HTMLElement & { visualTexMathfield?: _Mathfield };

function escapeHtmlAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function latexToMarkup(mf: _Mathfield, latex: string): string {
  const context = new Context({ from: mf.context });

  const root = new Atom({
    mode: 'math',
    type: 'root',
    body: parseLatex(latex, { context }),
  });

  return visualTexLayoutEngine.renderAtoms(root.body ?? [], { display: true });
}

export function showSuggestionPopover(
  mf: _Mathfield,
  suggestions: string[]
): void {
  if (suggestions.length === 0) {
    hideSuggestionPopover(mf);
    return;
  }

  let template = '';
  for (const [i, suggestion] of suggestions.entries()) {
    // A suggestion is the LaTeX it inserts; the list shows its command,
    // label and preview
    const info = visualTexCompletionInfo(mf, suggestion);
    const command = suggestion;
    const commandMarkup = latexToMarkup(mf, info.preview);
    const keybinding = getKeybindingsForCommand(mf.keybindings, info.command).join(
      '<br>'
    );

    template += `<li role="button" data-command="${escapeHtmlAttr(command)}" ${
      i === mf.suggestionIndex ? 'class=ML__popover__current' : ''
    }>${visualTexCandidateNumberMarkup(mf, i)}<span class="ML__popover__latex">${escapeHtmlAttr(info.command)}${info.label ? ' · ' + escapeHtmlAttr(info.label) : ''}</span><span class="ML__popover__command">${commandMarkup}</span>`;

    if (keybinding)
      template += `<span class="ML__popover__keybinding">${keybinding}</span>`;

    template += '</li>';
  }
  const panel = createSuggestionPopover(mf, `<ul>${template}</ul>`);
  if (isSuggestionPopoverVisible()) {
    panel
      .querySelector('.ML__popover__current')
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  setTimeout(() => {
    if (panel?.isConnected && panel.visualTexMathfield === mf && panel.getAttribute('aria-hidden') !== 'true' && !isSuggestionPopoverVisible()) {
      panel.classList.add('is-visible');
      updateSuggestionPopoverPosition(mf);
      panel
        .querySelector('.ML__popover__current')
        ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  }, 32);
}

export function isSuggestionPopoverVisible(): boolean {
  const panel = document.getElementById('mathlive-suggestion-popover');
  if (!panel) return false;
  return panel.classList.contains('is-visible');
}

export function updateSuggestionPopoverPosition(
  mf: _Mathfield,
  options?: { deferred: boolean }
): void {
  // Check that the mathfield is still valid
  // (we're calling ourselves from requestAnimationFrame() and the mathfield
  // could have gotten destroyed
  if (!mf.element || mf.element.mathfield !== mf) return;

  if (!isSuggestionPopoverVisible() || (document.getElementById('mathlive-suggestion-popover') as SuggestionPanel | null)?.visualTexMathfield !== mf) return;

  if (mf.model.at(mf.model.position)?.type !== 'latex') {
    hideSuggestionPopover(mf);
    return;
  }

  if (options?.deferred) {
    // Call ourselves again later, typically after the
    // rendering/layout of the DOM has been completed
    // (don't do it on next frame, it might be too soon)
    setTimeout(() => updateSuggestionPopoverPosition(mf), 32);
    return;
  }

  const position = getCaretPoint(mf);
  if (!position) return;

  // Get screen width & height (browser compatibility)
  const viewportHeight =
    window.innerHeight ||
    document.documentElement.clientHeight ||
    document.body.clientHeight;
  const viewportWidth =
    window.innerWidth ||
    document.documentElement.clientWidth ||
    document.body.clientWidth;

  // Get scrollbar size. This would be 0 in mobile device (also no needed).
  const scrollbarWidth =
    window.innerWidth - document.documentElement.clientWidth;
  const scrollbarHeight =
    window.innerHeight - document.documentElement.clientHeight;
  // Prevent screen overflow horizontal.
  const panel = document.getElementById('mathlive-suggestion-popover')!;
  if (position.x + panel.offsetWidth / 2 > viewportWidth - scrollbarWidth) {
    panel.style.left = `${
      viewportWidth - panel.offsetWidth - scrollbarWidth
    }px`;
  } else if (position.x - panel.offsetWidth / 2 < 0) panel.style.left = '0';
  else panel.style.left = `${position.x - panel.offsetWidth / 2}px`;

  // VisualTeX: keep the list inside the viewport, centred on the caret, and
  // give it the height the side it opens on has room for
  const viewportPadding = 8;
  const tipGap = 15;
  const viewportRight = viewportWidth - scrollbarWidth - viewportPadding;
  const viewportBottom =
    viewportHeight - scrollbarHeight - viewportPadding;
  const centeredLeft = position.x - panel.offsetWidth / 2;
  panel.style.left = `${Math.max(
    viewportPadding,
    Math.min(centeredLeft, viewportRight - panel.offsetWidth),
  )}px`;

  const spaceAbove = Math.max(
    0,
    position.y - position.height - viewportPadding - tipGap,
  );
  const spaceBelow = Math.max(0, viewportBottom - position.y - tipGap);
  const placeAbove = spaceAbove > spaceBelow;
  const availableHeight = Math.max(
    36,
    Math.min(320, placeAbove ? spaceAbove : spaceBelow),
  );
  panel.style.setProperty(
    '--visualtex-suggestion-max-height',
    `${availableHeight}px`,
  );

  // And position the popover right below or above the caret
  if (placeAbove) {
    panel.classList.add('ML__popover--reverse-direction');
    panel.classList.remove('top-tip');
    panel.classList.add('bottom-tip');
    panel.style.top = `${Math.max(
      viewportPadding,
      position.y - position.height - panel.offsetHeight - tipGap,
    )}px`;
  } else {
    panel.classList.remove('ML__popover--reverse-direction');
    panel.classList.add('top-tip');
    panel.classList.remove('bottom-tip');
    panel.style.top = `${Math.max(
      viewportPadding,
      Math.min(position.y + tipGap, viewportBottom - panel.offsetHeight),
    )}px`;
  }
}

export function hideSuggestionPopover(mf: _Mathfield): void {
  mf.suggestionIndex = 0;
  const panel = document.getElementById('mathlive-suggestion-popover') as SuggestionPanel | null;
  if (panel?.visualTexMathfield === mf) { panel.classList.remove('is-visible'); panel.setAttribute('aria-hidden', 'true'); }
}

export function createSuggestionPopover(
  mf: _Mathfield,
  html: string
): SuggestionPanel {
  let panel = document.getElementById('mathlive-suggestion-popover') as SuggestionPanel | null;
  if (!panel) {
    injectStylesheet('suggestion-popover'); injectStylesheet('core');
    panel = getSharedElement('mathlive-suggestion-popover') as SuggestionPanel;
    panel.addEventListener('pointerdown', (ev) => ev.preventDefault());
    panel.addEventListener('click', (ev) => {
      // a click accepts the clicked candidate for the field that owns the list
      const item = (ev.target as HTMLElement).closest<HTMLElement>('li[data-command]');
      const owner = panel!.visualTexMathfield;
      if (!item || !owner) return;
      const entries: string[] = owner.visualTexCompletion?.suggestions || [];
      const insertionLatex =
        item.dataset.visualtexInsertionLatex || item.dataset.command;
      const index = entries.indexOf(insertionLatex!);
      if (index < 0) return;
      owner.suggestionIndex = index;
      complete(owner, 'accept-all'); owner.dirty = true; owner.focus();
      requestUpdate(owner);
    });
  }
  panel.visualTexMathfield = mf;
  panel.setAttribute('aria-hidden', 'false');
  panel.innerHTML = globalThis.MathfieldElement.createHTML(html);
  for (const item of panel.querySelectorAll<HTMLElement>('li[data-command]')) {
    const insertionLatex = item.dataset.command || '';
    const info = visualTexCompletionInfo(mf, insertionLatex);
    item.dataset.visualtexInsertionLatex = insertionLatex;
    item.dataset.command = info.command || insertionLatex;
  }

  return panel;
}

export function disposeSuggestionPopover(mf: _Mathfield): void {
  const panel = document.getElementById('mathlive-suggestion-popover') as SuggestionPanel | null;
  if (!panel || panel.visualTexMathfield !== mf) return;
  releaseSharedElement('mathlive-suggestion-popover');
  releaseStylesheet('suggestion-popover');
  releaseStylesheet('core');
}
