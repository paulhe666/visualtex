const FLOATING_LAYER_SELECTOR = [
  '[role="menu"]',
  '[role="listbox"]',
  '[data-visualtex-floating-layer]',
  '#mathlive-suggestion-popover',
].join(',');

const VIEWPORT_PADDING = 10;
const MIN_FLOATING_LAYER_SIZE = 96;

type InlineStyleSnapshot = {
  translate: string;
  maxWidth: string;
  maxHeight: string;
  overflowX: string;
  overflowY: string;
  boxSizing: string;
};

const originalInlineStyles = new WeakMap<HTMLElement, InlineStyleSnapshot>();

function rememberInlineStyles(layer: HTMLElement) {
  if (originalInlineStyles.has(layer)) return;
  originalInlineStyles.set(layer, {
    translate: layer.style.translate,
    maxWidth: layer.style.maxWidth,
    maxHeight: layer.style.maxHeight,
    overflowX: layer.style.overflowX,
    overflowY: layer.style.overflowY,
    boxSizing: layer.style.boxSizing,
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

function fitFloatingLayer(layer: HTMLElement) {
  if (!layer.isConnected) return;
  rememberInlineStyles(layer);
  restoreManagedInlineStyles(layer);

  const initialRect = layer.getBoundingClientRect();
  if (initialRect.width <= 0 || initialRect.height <= 0) return;

  const boundary = visibleBoundaryFor(layer);
  const availableWidth = Math.max(
    MIN_FLOATING_LAYER_SIZE,
    boundary.right - boundary.left,
  );
  const availableHeight = Math.max(
    MIN_FLOATING_LAYER_SIZE,
    boundary.bottom - boundary.top,
  );
  const initialScaleX =
    layer.offsetWidth > 0 ? initialRect.width / layer.offsetWidth : 1;
  const initialScaleY =
    layer.offsetHeight > 0 ? initialRect.height / layer.offsetHeight : 1;

  if (initialRect.width > availableWidth) {
    layer.style.maxWidth = `${Math.floor(
      availableWidth / Math.max(0.1, initialScaleX),
    )}px`;
    layer.style.boxSizing = 'border-box';
    layer.style.overflowX = 'auto';
  }
  if (initialRect.height > availableHeight) {
    layer.style.maxHeight = `${Math.floor(
      availableHeight / Math.max(0.1, initialScaleY),
    )}px`;
    layer.style.boxSizing = 'border-box';
    layer.style.overflowY = 'auto';
  }

  let shiftX = 0;
  let shiftY = 0;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const rect = layer.getBoundingClientRect();
    let viewportShiftX = 0;
    let viewportShiftY = 0;

    if (rect.left < boundary.left) {
      viewportShiftX += boundary.left - rect.left;
    }
    if (rect.right + viewportShiftX > boundary.right) {
      viewportShiftX -= rect.right + viewportShiftX - boundary.right;
    }
    if (rect.top < boundary.top) {
      viewportShiftY += boundary.top - rect.top;
    }
    if (rect.bottom + viewportShiftY > boundary.bottom) {
      viewportShiftY -= rect.bottom + viewportShiftY - boundary.bottom;
    }

    if (Math.abs(viewportShiftX) < 0.25 && Math.abs(viewportShiftY) < 0.25) {
      break;
    }

    const scaleX = layer.offsetWidth > 0 ? rect.width / layer.offsetWidth : 1;
    const scaleY = layer.offsetHeight > 0 ? rect.height / layer.offsetHeight : 1;
    shiftX += viewportShiftX / Math.max(0.1, scaleX);
    shiftY += viewportShiftY / Math.max(0.1, scaleY);
    layer.style.translate = `${shiftX}px ${shiftY}px`;
  }

  layer.dataset.visualtexAutoAvoidAdjusted = 'true';
}

function visibleFloatingLayers() {
  return Array.from(
    document.querySelectorAll<HTMLElement>(FLOATING_LAYER_SELECTOR),
  ).filter((layer) => {
    const style = window.getComputedStyle(layer);
    return style.display !== 'none' && style.visibility !== 'hidden';
  });
}

export function installFloatingLayerAutoAvoidance() {
  let frame = 0;
  let disposed = false;
  const observedLayers = new Set<HTMLElement>();

  const update = () => {
    if (disposed || frame) return;
    frame = window.requestAnimationFrame(() => {
      frame = 0;
      const layers = new Set(visibleFloatingLayers());
      for (const layer of observedLayers) {
        if (layers.has(layer)) continue;
        resizeObserver.unobserve(layer);
        restoreManagedInlineStyles(layer);
        originalInlineStyles.delete(layer);
        observedLayers.delete(layer);
      }
      for (const layer of layers) {
        if (!observedLayers.has(layer)) {
          observedLayers.add(layer);
          resizeObserver.observe(layer);
        }
        fitFloatingLayer(layer);
      }
    });
  };
  const resizeObserver = new ResizeObserver(update);

  const handleMotion = (event: Event) => {
    const target = event.target;
    if (target instanceof Element && (
      target.closest(FLOATING_LAYER_SELECTOR) || target.querySelector(FLOATING_LAYER_SELECTOR)
    )) {
      update();
    }
  };

  const containsFloatingLayer = (node: Node) => node instanceof Element && (
    node.matches(FLOATING_LAYER_SELECTOR) || node.querySelector(FLOATING_LAYER_SELECTOR)
  );
  const observer = new MutationObserver((records) => {
    if (records.some((record) =>
      (record.target instanceof Element && record.target.closest(FLOATING_LAYER_SELECTOR)) ||
      (record.type === 'attributes' && containsFloatingLayer(record.target)) ||
      [...record.addedNodes, ...record.removedNodes].some(containsFloatingLayer)
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
  const motionEvents = ['transitionrun', 'transitionend', 'animationstart', 'animationend'];
  for (const event of motionEvents) document.addEventListener(event, handleMotion, true);
  update();

  return () => {
    disposed = true;
    window.cancelAnimationFrame(frame);
    observer.disconnect();
    resizeObserver.disconnect();
    window.removeEventListener('resize', update);
    window.removeEventListener('scroll', update, true);
    for (const event of motionEvents) document.removeEventListener(event, handleMotion, true);
    for (const layer of observedLayers) {
      restoreManagedInlineStyles(layer);
      originalInlineStyles.delete(layer);
    }
    observedLayers.clear();
  };
}
