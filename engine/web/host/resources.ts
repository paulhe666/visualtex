// Hash-verified WASM/font loading for the isolated engine host. This module
// never discovers system fonts, reads Office sessions or changes browser flags.
export interface AssetDescriptor {
  readonly id: string;
  readonly kind: 'wasm' | 'font';
  readonly url: string;
  readonly sha256: string;
  readonly byteLength: number;
}
export interface AssetManifest {
  readonly schemaVersion: 1;
  readonly engineProtocol: number;
  readonly buildId: string;
  readonly assets: readonly AssetDescriptor[];
}
export interface VerifiedAsset {
  readonly id: string;
  readonly sha256: string;
  readonly bytes: Uint8Array;
}
export type ResourceErrorCode = 'invalidManifest' | 'versionMismatch' | 'unknownAsset' | 'unsafeUrl'
  | 'resourceLimit' | 'networkFailure' | 'badMime' | 'lengthMismatch' | 'hashMismatch'
  | 'cryptoUnavailable' | 'cancelled' | 'timeout' | 'disposed';
export class ResourceError extends Error {
  constructor(readonly code: ResourceErrorCode, message: string) {
    super(message); this.name = 'ResourceError';
  }
}
export interface ResourceLoaderOptions {
  readonly baseUrl: string;
  readonly expectedBuildId: string;
  readonly expectedEngineProtocol: number;
  readonly maxAssetBytes?: number;
  readonly maxResidentBytes?: number;
  readonly maxConcurrentLoads?: number;
  readonly timeoutMs?: number;
  readonly fetch?: typeof globalThis.fetch;
  readonly subtle?: Pick<SubtleCrypto, 'digest'>;
}
interface Pending {
  readonly controller: AbortController;
  promise: Promise<Uint8Array>;
  waiters: number;
  settled: boolean;
}
function fail(code: ResourceErrorCode, message: string): never { throw new ResourceError(code, message); }
const positiveInteger = (value: number): boolean => Number.isSafeInteger(value) && value > 0;
const validId = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0
  && value.length <= 256 && !/[\u0000-\u001f\u007f]/u.test(value);

/** Cache ownership stays in this loader. Callers receive separate buffers, so a
 * consumer mutating/freeing its copy cannot corrupt another document's font.
 * Resident budget covers cached and in-flight asset buffers, not caller copies. */
export class VerifiedAssetLoader {
  private readonly assets = new Map<string, AssetDescriptor>();
  private readonly urls = new Map<string, string>();
  private readonly cache = new Map<string, Uint8Array>();
  private readonly pending = new Map<string, Pending>();
  private readonly fetcher: typeof globalThis.fetch;
  private readonly subtle: Pick<SubtleCrypto, 'digest'>;
  private readonly budget: number;
  private readonly concurrency: number;
  private readonly timeoutMs: number;
  private residentBytes = 0;
  private pendingBytes = 0;
  private disposed = false;

  constructor(manifest: AssetManifest, options: ResourceLoaderOptions) {
    if (!manifest || manifest.schemaVersion !== 1 || !Array.isArray(manifest.assets)
      || !validId(manifest.buildId) || !positiveInteger(manifest.engineProtocol)
      || !validId(options.expectedBuildId) || !positiveInteger(options.expectedEngineProtocol)) {
      fail('invalidManifest', 'Invalid engine resource manifest');
    }
    if (manifest.buildId !== options.expectedBuildId || manifest.engineProtocol !== options.expectedEngineProtocol) {
      fail('versionMismatch', 'Host bindings and resource manifest belong to different engine builds');
    }
    const maxAsset = options.maxAssetBytes ?? 64 * 1024 * 1024;
    this.budget = options.maxResidentBytes ?? 128 * 1024 * 1024;
    this.concurrency = options.maxConcurrentLoads ?? 4;
    this.timeoutMs = options.timeoutMs ?? 30_000;
    if (![maxAsset, this.budget, this.concurrency, this.timeoutMs].every(positiveInteger)
      || maxAsset > this.budget || maxAsset > 64 * 1024 * 1024 || this.budget > 256 * 1024 * 1024
      || this.concurrency > 16 || this.timeoutMs > 2_147_483_647 || manifest.assets.length > 256) {
      fail('invalidManifest', 'Invalid resource limits');
    }
    const base = this.safeBase(options.baseUrl);
    for (const asset of manifest.assets) {
      if (!asset || !validId(asset.id) || this.assets.has(asset.id)
        || !['wasm', 'font'].includes(asset.kind) || typeof asset.sha256 !== 'string'
        || !/^[0-9a-f]{64}$/u.test(asset.sha256) || !positiveInteger(asset.byteLength) || asset.byteLength > maxAsset) {
        fail('invalidManifest', 'Invalid, duplicate or over-budget asset descriptor');
      }
      const url = this.safeAssetUrl(asset.url, base);
      this.assets.set(asset.id, Object.freeze({ ...asset }));
      this.urls.set(asset.id, url);
    }
    const subtle = options.subtle ?? globalThis.crypto?.subtle;
    if (!subtle || typeof subtle.digest !== 'function') fail('cryptoUnavailable', 'SHA-256 verification is unavailable in this host');
    this.subtle = subtle;
    const fetcher = options.fetch ?? globalThis.fetch;
    if (typeof fetcher !== 'function') fail('networkFailure', 'Fetch is unavailable in this host');
    this.fetcher = fetcher.bind(globalThis);
  }

  private safeBase(value: string): URL {
    let base: URL;
    try { base = new URL(value); } catch { return fail('unsafeUrl', 'Invalid resource base URL'); }
    if (!['http:', 'https:'].includes(base.protocol) || base.username || base.password || base.search || base.hash
      || !base.pathname.endsWith('/')) fail('unsafeUrl', 'Resource base must be an HTTP(S) directory URL');
    return base;
  }
  private safeAssetUrl(value: string, base: URL): string {
    if (typeof value !== 'string' || !value || /[\\\u0000-\u0020]/u.test(value)) fail('unsafeUrl', 'Invalid asset URL');
    let url: URL;
    try {
      const decoded = decodeURIComponent(value);
      if (decoded.split(/[/?#]/u).some(part => part === '..' || part === '.')
        || /%[0-9a-f]{2}/iu.test(decoded)
        || decoded.includes('\\') || /[\u0000-\u001f\u007f]/u.test(decoded)) {
        return fail('unsafeUrl', 'Asset traversal is not allowed');
      }
      url = new URL(value, base);
    } catch (error) {
      if (error instanceof ResourceError) throw error;
      return fail('unsafeUrl', 'Invalid asset URL encoding');
    }
    if (url.origin !== base.origin || !url.pathname.startsWith(base.pathname)
      || url.username || url.password || url.search || url.hash) fail('unsafeUrl', 'Asset URL escapes the fixed resource directory');
    return url.href;
  }

  get stats(): Readonly<{ cachedBytes: number; pendingBytes: number; activeLoads: number; cachedAssets: number }> {
    return { cachedBytes: this.residentBytes, pendingBytes: this.pendingBytes,
      activeLoads: this.pending.size, cachedAssets: this.cache.size };
  }

  async load(id: string, signal?: AbortSignal): Promise<VerifiedAsset> {
    if (this.disposed) fail('disposed', 'Resource loader is disposed');
    if (signal?.aborted) fail('cancelled', 'Resource request was cancelled');
    const asset = this.assets.get(id);
    if (!asset) fail('unknownAsset', 'Asset is not present in the fixed manifest');
    // Include role and URL: a font alias must not bypass the WASM MIME check.
    const key = `${asset.kind}:${this.urls.get(id)}:${asset.sha256}:${asset.byteLength}`;
    const cached = this.cache.get(key);
    if (cached) {
      this.cache.delete(key); this.cache.set(key, cached); // LRU touch
      return { id, sha256: asset.sha256, bytes: cached.slice() };
    }
    let pending = this.pending.get(key);
    if (pending?.controller.signal.aborted) {
      // A newly interested caller must not subscribe to a cancelled generation.
      // Await its cleanup so reservation/concurrency accounting remains exact.
      try { await pending.promise; } catch { /* retry belongs to a new request */ }
      return this.load(id, signal);
    }
    if (!pending) {
      if (this.pending.size >= this.concurrency) fail('resourceLimit', 'Concurrent resource budget exhausted');
      while (this.residentBytes + this.pendingBytes + asset.byteLength > this.budget && this.cache.size) {
        const oldest = this.cache.entries().next().value as [string, Uint8Array];
        this.cache.delete(oldest[0]); this.residentBytes -= oldest[1].byteLength;
      }
      if (this.pendingBytes + asset.byteLength > this.budget) fail('resourceLimit', 'In-flight resource byte budget exhausted');
      const item: Pending = { controller: new AbortController(), promise: Promise.resolve(new Uint8Array()), waiters: 0, settled: false };
      this.pendingBytes += asset.byteLength;
      this.pending.set(key, item);
      item.promise = this.read(asset, this.urls.get(id)!, item.controller)
        .then(bytes => {
          if (this.disposed || item.controller.signal.aborted) fail('cancelled', 'Resource request was cancelled');
          this.cache.set(key, bytes); this.residentBytes += bytes.byteLength;
          return bytes;
        })
        .finally(() => {
          item.settled = true; this.pendingBytes -= asset.byteLength;
          if (this.pending.get(key) === item) this.pending.delete(key);
        });
      pending = item;
    }
    const bytes = await this.subscribe(pending, signal);
    if (this.disposed) fail('disposed', 'Resource loader is disposed');
    if (signal?.aborted) fail('cancelled', 'Resource request was cancelled');
    return { id, sha256: asset.sha256, bytes: bytes.slice() };
  }

  private subscribe(pending: Pending, signal?: AbortSignal): Promise<Uint8Array> {
    pending.waiters += 1;
    return new Promise<Uint8Array>((resolve, reject) => {
      let done = false;
      const finish = (action: () => void): void => {
        if (done) return;
        done = true; signal?.removeEventListener('abort', cancelled);
        pending.waiters -= 1;
        if (!pending.waiters && !pending.settled) pending.controller.abort();
        action();
      };
      const cancelled = (): void => finish(() => reject(new ResourceError('cancelled', 'Resource request was cancelled')));
      signal?.addEventListener('abort', cancelled, { once: true });
      pending.promise.then(bytes => finish(() => resolve(bytes)), error => finish(() => reject(error)));
      if (signal?.aborted) cancelled();
    });
  }

  private async read(asset: AssetDescriptor, url: string, controller: AbortController): Promise<Uint8Array> {
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; controller.abort(); }, this.timeoutMs);
    let reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
    const cancelReader = (): void => { void reader?.cancel().catch(() => {}); };
    controller.signal.addEventListener('abort', cancelReader, { once: true });
    try {
      const response = await this.fetcher(url, { credentials: 'omit', mode: 'same-origin', redirect: 'error',
        cache: 'no-cache', signal: controller.signal });
      // Acquire the body before validation so every header/status failure also
      // cancels the stream rather than leaving a rejected download running.
      reader = response.body?.getReader();
      if (!response.ok || response.status === 206 || response.redirected) fail('networkFailure', 'Resource request did not return a complete successful response');
      const mime = (response.headers.get('content-type') ?? '').split(';')[0].trim().toLowerCase();
      if (asset.kind === 'wasm' ? mime !== 'application/wasm'
        : !['font/ttf', 'font/otf', 'font/woff', 'font/woff2', 'application/font-sfnt', 'application/octet-stream'].includes(mime)) {
        fail('badMime', 'Resource response has an unexpected MIME type');
      }
      // Content-Length may describe compressed transfer bytes. The streamed body
      // below is authoritative for the uncompressed, hash-checked asset length.
      const length = response.headers.get('content-length');
      if (length !== null && !response.headers.get('content-encoding') && Number(length) !== asset.byteLength) {
        fail('lengthMismatch', 'Resource response length differs from the manifest');
      }
      if (!reader) fail('lengthMismatch', 'Resource body is absent');
      const bytes = new Uint8Array(asset.byteLength);
      let offset = 0;
      for (;;) {
        if (controller.signal.aborted) fail('cancelled', 'Resource request was cancelled');
        const chunk = await reader.read();
        if (controller.signal.aborted) fail('cancelled', 'Resource request was cancelled');
        if (chunk.done) break;
        if (!(chunk.value instanceof Uint8Array) || chunk.value.byteLength > bytes.length - offset) {
          fail('lengthMismatch', 'Resource body exceeds its declared byte budget');
        }
        bytes.set(chunk.value, offset); offset += chunk.value.byteLength;
      }
      if (offset !== bytes.byteLength) fail('lengthMismatch', 'Resource body is truncated');
      const hash = new Uint8Array(await this.subtle.digest('SHA-256', bytes.buffer));
      if (controller.signal.aborted) fail('cancelled', 'Resource request was cancelled');
      const hex = Array.from(hash, byte => byte.toString(16).padStart(2, '0')).join('');
      if (hex !== asset.sha256) fail('hashMismatch', 'Resource SHA-256 differs from the fixed manifest');
      return bytes;
    } catch (error) {
      if (timedOut) fail('timeout', 'Resource load exceeded its time budget');
      if (error instanceof ResourceError) throw error;
      if (controller.signal.aborted) fail('cancelled', 'Resource request was cancelled');
      throw new ResourceError('networkFailure', error instanceof Error ? error.message : 'Resource load failed');
    } finally {
      clearTimeout(timer);
      controller.signal.removeEventListener('abort', cancelReader);
      if (reader) {
        try { await reader.cancel(); } catch { /* preserve the original failure */ }
        reader.releaseLock();
      }
    }
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const item of this.pending.values()) item.controller.abort();
    this.cache.clear(); this.residentBytes = 0;
  }
}
