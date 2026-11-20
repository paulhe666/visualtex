function fallbackUuid(): string {
  const bytes = new Uint8Array(16);
  const cryptoObject = globalThis.crypto;
  if (cryptoObject && typeof cryptoObject.getRandomValues === "function") {
    cryptoObject.getRandomValues(bytes);
  } else {
    for (let index = 0; index < bytes.length; index += 1) {
      bytes[index] = Math.floor(Math.random() * 256);
    }
  }

  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, (value) => value.toString(16).padStart(2, "0"));
  return (
    hex.slice(0, 4).join("") +
    "-" +
    hex.slice(4, 6).join("") +
    "-" +
    hex.slice(6, 8).join("") +
    "-" +
    hex.slice(8, 10).join("") +
    "-" +
    hex.slice(10, 16).join("")
  );
}

export function createUuid(): string {
  const cryptoObject = globalThis.crypto;
  if (cryptoObject && typeof cryptoObject.randomUUID === "function") {
    return cryptoObject.randomUUID();
  }
  return fallbackUuid();
}

function installArrayAt() {
  if (typeof Array.prototype.at === "function") return;
  Object.defineProperty(Array.prototype, "at", {
    configurable: true,
    writable: true,
    value<T>(this: T[], index: number): T | undefined {
      const length = this.length >>> 0;
      let offset = Number.isFinite(index) ? Math.trunc(index) : 0;
      if (offset < 0) offset += length;
      return offset < 0 || offset >= length ? undefined : this[offset];
    },
  });
}

function installStringReplaceAll() {
  if (typeof String.prototype.replaceAll === "function") return;
  Object.defineProperty(String.prototype, "replaceAll", {
    configurable: true,
    writable: true,
    value(
      this: string,
      searchValue: string | RegExp,
      replaceValue: string,
    ): string {
      const source = String(this);
      if (searchValue instanceof RegExp) {
        if (!searchValue.global) {
          throw new TypeError("replaceAll requires a global regular expression");
        }
        return source.replace(searchValue, replaceValue);
      }
      const search = String(searchValue);
      if (!search) {
        return replaceValue + Array.from(source).join(replaceValue) + replaceValue;
      }
      return source.split(search).join(replaceValue);
    },
  });
}

function installQueueMicrotask() {
  if (typeof globalThis.queueMicrotask === "function") return;
  Object.defineProperty(globalThis, "queueMicrotask", {
    configurable: true,
    writable: true,
    value(callback: VoidFunction) {
      Promise.resolve()
        .then(callback)
        .catch((error) => globalThis.setTimeout(() => {
          throw error;
        }, 0));
    },
  });
}

function installResizeObserver() {
  if (typeof globalThis.ResizeObserver === "function") return;

  class VisualTexResizeObserver implements ResizeObserver {
    private readonly targets = new Set<Element>();
    private scheduled = false;

    constructor(private readonly callback: ResizeObserverCallback) {}

    private readonly handleResize = () => this.schedule();

    private schedule() {
      if (this.scheduled) return;
      this.scheduled = true;
      globalThis.queueMicrotask(() => {
        this.scheduled = false;
        if (this.targets.size === 0) return;
        this.callback([], this);
      });
    }

    observe(target: Element) {
      const wasEmpty = this.targets.size === 0;
      this.targets.add(target);
      if (wasEmpty && typeof window !== "undefined") {
        window.addEventListener("resize", this.handleResize);
      }
      this.schedule();
    }

    unobserve(target: Element) {
      this.targets.delete(target);
      if (this.targets.size === 0 && typeof window !== "undefined") {
        window.removeEventListener("resize", this.handleResize);
      }
    }

    disconnect() {
      this.targets.clear();
      if (typeof window !== "undefined") {
        window.removeEventListener("resize", this.handleResize);
      }
    }
  }

  Object.defineProperty(globalThis, "ResizeObserver", {
    configurable: true,
    writable: true,
    value: VisualTexResizeObserver,
  });
}

export function installBrowserCompatibility() {
  installArrayAt();
  installStringReplaceAll();
  installQueueMicrotask();
  installResizeObserver();
}
