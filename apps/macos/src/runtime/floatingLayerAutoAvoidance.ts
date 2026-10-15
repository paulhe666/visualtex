const FLOATING_LAYER_SELECTOR = [
  '[role="menu"]',
  '[role="listbox"]',
  '[data-visualtex-floating-layer]',
  '#mathlive-suggestion-popover',
].join(',');

// CodeMirror places and sizes its own tooltips (and its completion list is a
// role="listbox" child of one). Fitting that list as if it floated on its
// own pushed it down inside the tooltip, leaving the top of the box empty.
const SELF_POSITIONED_OWNER_SELECTOR = '.cm-tooltip';

const VIEWPORT_PADDING = 10;
const MIN_FLOATING_LAYER_SIZE = 96;
// Sub-pixel differences are rounding noise, not a new position. Ignoring them
// keeps a re-fit from writing a style that would trigger yet another re-fit.
const POSITION_EPSILON = 0.5;

type InlineStyleSnapshot = {
  translate: string;
  maxWidth: string;
  maxHeight: string;
  overflowX: string;
  overflowY: string;
  boxSizing: string;
  // Stylesheet limits seen before this module wrote any inline style.
  cssMaxWidth: string;
  cssMaxHeight: string;
};

const originalInlineStyles = new WeakMap<HTMLElement, InlineStyleSnapshot>();

function rememberInlineStyles(layer: HTMLElement) {
  if (originalInlineStyles.has(layer)) return;
  const computed = window.getComputedStyle(layer);
  originalInlineStyles.set(layer, {
    translate: layer.style.translate,
    maxWidth: layer.style.maxWidth,
    maxHeight: layer.style.maxHeight,
    overflowX: layer.style.overflowX,
    overflowY: layer.style.overflowY,
    boxSizing: layer.style.boxSizing,
    cssMaxWidth: computed.maxWidth,
    cssMaxHeight: computed.maxHeight,
  });
}

function restoreManagedInlineStyles(layer: HTMLElement) {
  const original = originalInlineStyles.get(layer);
  if (!original) return;
  layer.style.translate = original.translate;
  layer.style.maxWidth = original.maxWidth;
  layer.style.maxHeight = original.maxHeight;
  layer.style.overflowX = original.overflowX;
  layer.style.overflowY = original.overflowY;
  layer.style.boxSizing = original.boxSizing;
  delete layer.dataset.visualtexAutoAvoidAdjusted;
}

function clipsAxis(value: string) {
  return /(?:auto|scroll|hidden|clip)/.test(value);
}

function visibleBoundaryFor(layer: HTMLElement) {
  let left = VIEWPORT_PADDING;
  let top = VIEWPORT_PADDING;
  let right = Math.max(left, window.innerWidth - VIEWPORT_PADDING);
  let bottom = Math.max(top, window.innerHeight - VIEWPORT_PADDING);

  for (
    let ancestor = layer.parentElement;
    ancestor && ancestor !== document.documentElement;
    ancestor = ancestor.parentElement
  ) {
    if (ancestor === document.body) continue;
    const style = window.getComputedStyle(ancestor);
    const clipsX = clipsAxis(style.overflowX);
    const clipsY = clipsAxis(style.overflowY);
    if (!clipsX && !clipsY) continue;

    const rect = ancestor.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) continue;
    if (clipsX) {
      left = Math.max(left, rect.left + VIEWPORT_PADDING);
      right = Math.min(right, rect.right - VIEWPORT_PADDING);
    }
    if (clipsY) {
      top = Math.max(top, rect.top + VIEWPORT_PADDING);
      bottom = Math.min(bottom, rect.bottom - VIEWPORT_PADDING);
    }
  }

  return {
    left,
    top,
    right: Math.max(left, right),
    bottom: Math.max(top, bottom),
  };
}

function parsePixelPair(value: string): [number, number] {
  if (!value || value === 'none') return [0, 0];
  const [x = '0', y = '0'] = value.trim().split(/\s+/);
  return [Number.parseFloat(x) || 0, Number.parseFloat(y) || 0];
}

/**
 * The layer's layout box in viewport coordinates, without our own `translate`
 * and without an entry animation's scale/offset. Measuring the transformed box
 * instead made every fit depend on the previous fit and on the animation frame,
 * so menus visibly slid after opening and could keep re-fitting each other.
 */
function untransformedRect(layer: HTMLElement) {
  const rect = layer.getBoundingClientRect();
  const width = layer.offsetWidth || rect.width;
  const height = layer.offsetHeight || rect.height;
  const style = window.getComputedStyle(layer);
  const [translateX, translateY] = parsePixelPair(style.translate);
  let scaleX = 1;
  let scaleY = 1;
  let offsetX = 0;
  let offsetY = 0;
  const matrix = /^matrix\(([^)]+)\)$/.exec(style.transform);
  if (matrix) {
    const [a, b, c, d, e, f] = matrix[1].split(',').map((part) => Number.parseFloat(part));
    if (Math.abs(b) < 1e-6 && Math.abs(c) < 1e-6 && a > 0 && d > 0) {
      scaleX = a;
      scaleY = d;
      offsetX = e;
      offsetY = f;
    }
  }
  const [originX, originY] = parsePixelPair(style.transformOrigin);
  const left = rect.left - translateX - offsetX - originX * (1 - scaleX);
  const top = rect.top - translateY - offsetY - originY * (1 - scaleY);
  return { left, top, right: left + width, bottom: top + height, width, height };
}

function cappedSize(cssLimit: string, available: number) {
  const px = `${Math.floor(available)}px`;
  return cssLimit && cssLimit !== 'none' ? `min(${cssLimit}, ${px})` : px;
}

function fitFloatingLayer(layer: HTMLElement) {
  if (!layer.isConnected) return;
  rememberInlineStyles(layer);
  const original = originalInlineStyles.get(layer)!;

  const box = untransformedRect(layer);
  if (box.width <= 0 || box.height <= 0) return;

  const boundary = visibleBoundaryFor(layer);
  const availableWidth = Math.max(MIN_FLOATING_LAYER_SIZE, boundary.right - boundary.left);
  const availableHeight = Math.max(MIN_FLOATING_LAYER_SIZE, boundary.bottom - boundary.top);

  // Size caps depend only on the boundary, never on the layer's current size,
  // so applying them cannot feed back into the next measurement.
  if (!original.maxWidth) {
    const cap = cappedSize(original.cssMaxWidth, availableWidth);
    if (layer.style.maxWidth !== cap) layer.style.maxWidth = cap;
  }
  if (!original.maxHeight) {
    const cap = cappedSize(original.cssMaxHeight, availableHeight);
    if (layer.style.maxHeight !== cap) layer.style.maxHeight = cap;
  }
  if (layer.scrollWidth > layer.clientWidth + 1 && layer.style.overflowX !== 'auto') {
    layer.style.boxSizing = 'border-box';
    layer.style.overflowX = 'auto';
  }
  if (layer.scrollHeight > layer.clientHeight + 1 && layer.style.overflowY !== 'auto') {
    layer.style.boxSizing = 'border-box';
    layer.style.overflowY = 'auto';
  }

  const fitted = untransformedRect(layer);
  let shiftX = 0;
  let shiftY = 0;
  if (fitted.left < boundary.left) shiftX = boundary.left - fitted.left;
  else if (fitted.right > boundary.right) {
    shiftX = Math.max(boundary.left - fitted.left, boundary.right - fitted.right);
  }
  if (fitted.top < boundary.top) shiftY = boundary.top - fitted.top;
  else if (fitted.bottom > boundary.bottom) {
    shiftY = Math.max(boundary.top - fitted.top, boundary.bottom - fitted.bottom);
  }

  const [currentX, currentY] = parsePixelPair(layer.style.translate);
  if (Math.abs(shiftX) < POSITION_EPSILON && Math.abs(shiftY) < POSITION_EPSILON) {
    if (layer.style.translate !== original.translate) layer.style.translate = original.translate;
  } else if (
    Math.abs(currentX - shiftX) >= POSITION_EPSILON ||
    Math.abs(currentY - shiftY) >= POSITION_EPSILON
  ) {
    layer.style.translate = `${Math.round(shiftX)}px ${Math.round(shiftY)}px`;
  }

  if (layer.dataset.visualtexAutoAvoidAdjusted !== 'true') {
    layer.dataset.visualtexAutoAvoidAdjusted = 'true';
  }
}

function visibleFloatingLayers() {
  return Array.from(
    document.querySelectorAll<HTMLElement>(FLOATING_LAYER_SELECTOR),
  ).filter((layer) => {
    if (layer.closest(SELF_POSITIONED_OWNER_SELECTOR)) return false;
    const style = window.getComputedStyle(layer);
    return style.display !== 'none' && style.visibility !== 'hidden';
  });
}

export function installFloatingLayerAutoAvoidance() {
  let frame = 0;
  let disposed = false;
  const observedLayers = new Set<HTMLElement>();

  const fitVisibleLayers = () => {
    const layers = new Set(visibleFloatingLayers());
    for (const layer of observedLayers) {
      if (layers.has(layer)) continue;
      resizeObserver.unobserve(layer);
      restoreManagedInlineStyles(layer);
      originalInlineStyles.delete(layer);
      observedLayers.delete(layer);
    }
    for (const layer of layers) {
      try {
        if (!observedLayers.has(layer)) {
          observedLayers.add(layer);
          resizeObserver.observe(layer);
        }
        fitFloatingLayer(layer);
      } catch {
        // One transient or host-owned floating layer must never prevent the
        // remaining VisualTeX/MathLive popovers from being fitted. Office
        // WebView2 can expose a detached node for one animation frame while a
        // native suggestion surface is being replaced.
      }
    }
  };
  const update = () => {
    if (disposed || frame) return;
    frame = window.requestAnimationFrame(() => {
      frame = 0;
      fitVisibleLayers();
    });
  };
  // A layer that has just been inserted is fitted in the same task, before
  // the browser paints it; waiting for the next frame showed it unadjusted
  // for one frame and then slid it into place.
  const fitNow = () => {
    if (disposed) return;
    if (frame) {
      window.cancelAnimationFrame(frame);
      frame = 0;
    }
    fitVisibleLayers();
  };
  // Fires when a layer's own size changes (content loaded, caps applied), not
  // on transforms, so entry animations do not re-fit it every frame.
  const resizeObserver = new ResizeObserver(update);

  const containsFloatingLayer = (node: Node) => node instanceof Element && (
    node.matches(FLOATING_LAYER_SELECTOR) || node.querySelector(FLOATING_LAYER_SELECTOR) !== null
  );
  // Only structural changes and class/hidden toggles can open, close or move a
  // layer. Hover transitions and entry animations inside a menu do not, and
  // reacting to them is what made open menus re-fit on every pointer move.
  const observer = new MutationObserver((records) => {
    if (records.some((record) => [...record.addedNodes].some(containsFloatingLayer) ||
      (record.type === 'attributes' && containsFloatingLayer(record.target)))) {
      fitNow();
    } else if (records.some((record) =>
      [...record.removedNodes].some(containsFloatingLayer) ||
      (record.type === 'childList' && record.target instanceof Element && record.target.closest(FLOATING_LAYER_SELECTOR))
    )) update();
  });
  observer.observe(document.documentElement, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['class', 'hidden'],
  });
  window.addEventListener('resize', update);
  window.addEventListener('scroll', update, true);
  update();

  return () => {
    disposed = true;
    window.cancelAnimationFrame(frame);
    observer.disconnect();
    resizeObserver.disconnect();
    window.removeEventListener('resize', update);
    window.removeEventListener('scroll', update, true);
    for (const layer of observedLayers) {
      restoreManagedInlineStyles(layer);
      originalInlineStyles.delete(layer);
    }
    observedLayers.clear();
  };
}
