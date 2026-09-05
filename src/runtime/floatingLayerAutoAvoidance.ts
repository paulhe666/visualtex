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
  const settleTimers = new Set<number>();

  const update = () => {
    if (disposed) return;
    window.cancelAnimationFrame(frame);
    frame = window.requestAnimationFrame(() => {
      for (const layer of visibleFloatingLayers()) {
        try {
          fitFloatingLayer(layer);
        } catch {
          // One transient or host-owned floating layer must never prevent the
          // remaining VisualTeX/MathLive popovers from being fitted. Office
          // WebView2 can expose a detached node for one animation frame while a
          // native suggestion surface is being replaced.
        }
      }
    });
  };

  const updateThroughOpeningAnimation = () => {
    update();
    for (const delay of [48, 160, 260]) {
      const timer = window.setTimeout(() => {
        settleTimers.delete(timer);
        update();
      }, delay);
      settleTimers.add(timer);
    }
  };

  const handleMotionSettled = (event: Event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.matches(FLOATING_LAYER_SELECTOR) || target.closest(FLOATING_LAYER_SELECTOR)) {
      update();
    }
  };

  const observer = new MutationObserver(updateThroughOpeningAnimation);
  observer.observe(document.documentElement, {
    childList: true,
    subtree: true,
  });
  window.addEventListener('resize', updateThroughOpeningAnimation);
  window.addEventListener('scroll', update, true);
  document.addEventListener('transitionend', handleMotionSettled, true);
  document.addEventListener('animationend', handleMotionSettled, true);
  updateThroughOpeningAnimation();

  return () => {
    disposed = true;
    window.cancelAnimationFrame(frame);
    for (const timer of settleTimers) window.clearTimeout(timer);
    settleTimers.clear();
    observer.disconnect();
    window.removeEventListener('resize', updateThroughOpeningAnimation);
    window.removeEventListener('scroll', update, true);
    document.removeEventListener('transitionend', handleMotionSettled, true);
    document.removeEventListener('animationend', handleMotionSettled, true);
    for (const layer of visibleFloatingLayers()) restoreManagedInlineStyles(layer);
  };
}
