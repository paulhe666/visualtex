import { ChildProcessWithoutNullStreams, spawn } from "node:child_process";

interface RpcResponse {
  jsonrpc: "2.0";
  id: number;
  result?: unknown;
  error?: { code: number; message: string; data?: unknown };
}

interface PendingRequest {
  resolve(value: unknown): void;
  reject(error: Error): void;
  timer: NodeJS.Timeout;
}

class CoreTransportError extends Error {}
class CoreRpcError extends Error {}

export type CoreConnectionStatus =
  | { state: "connecting"; attempt: number }
  | { state: "connected" }
  | { state: "disconnected"; detail: string }
  | { state: "failed"; detail: string };

export interface CoreRequestOptions {
  retryAfterReconnect?: boolean;
  timeoutMs?: number;
}

const RECONNECT_DELAYS_MS = [0, 250, 750, 1_500];
const DEFAULT_REQUEST_TIMEOUT_MS = 120_000;

export class CoreClient {
  private process: ChildProcessWithoutNullStreams | null = null;
  private pending = new Map<number, PendingRequest>();
  private nextId = 1;
  private stdoutBuffer = "";
  private connecting: Promise<void> | null = null;
  private reconnectTimer: NodeJS.Timeout | null = null;
  private disposed = false;
  private everConnected = false;
  private statusListeners = new Set<(status: CoreConnectionStatus) => void>();

  constructor(
    private readonly executable: string,
    private readonly projectRoot: string,
    private readonly globalArguments: string[] = [],
  ) {}

  onStatus(listener: (status: CoreConnectionStatus) => void): { dispose(): void } {
    this.statusListeners.add(listener);
    return { dispose: () => this.statusListeners.delete(listener) };
  }

  async connect(): Promise<void> {
    await this.ensureConnected();
  }

  async request<T>(
    method: string,
    params: unknown = {},
    options: CoreRequestOptions = {},
  ): Promise<T> {
    await this.ensureConnected();
    try {
      return await this.send<T>(method, params, options.timeoutMs);
    } catch (error) {
      if (
        !options.retryAfterReconnect
        || this.disposed
        || !(error instanceof CoreTransportError)
      ) {
        throw error;
      }
      await this.restart();
      return this.send<T>(method, params, options.timeoutMs);
    }
  }

  async restart(): Promise<void> {
    this.stopCurrentProcess(new Error("VisualTeX Core restart requested"));
    await this.ensureConnected();
  }

  dispose(): void {
    this.disposed = true;
    if (this.reconnectTimer) clearTimeout(this.reconnectTimer);
    this.reconnectTimer = null;
    this.stopCurrentProcess(new Error("VisualTeX Core session disposed"));
    this.statusListeners.clear();
  }

  private async ensureConnected(): Promise<void> {
    if (this.disposed) throw new Error("VisualTeX Core session is disposed");
    if (this.isRunning()) return;
    if (this.connecting) return this.connecting;
    this.connecting = this.connectWithRetries().finally(() => {
      this.connecting = null;
    });
    return this.connecting;
  }

  private async connectWithRetries(): Promise<void> {
    let lastError: Error | null = null;
    for (let attempt = 0; attempt < RECONNECT_DELAYS_MS.length; attempt += 1) {
      if (this.disposed) throw new Error("VisualTeX Core session is disposed");
      const delay = RECONNECT_DELAYS_MS[attempt] ?? 0;
      if (delay > 0) await wait(delay);
      this.emitStatus({ state: "connecting", attempt: attempt + 1 });
      try {
        await this.startProcess();
        this.everConnected = true;
        this.emitStatus({ state: "connected" });
        return;
      } catch (error) {
        lastError = asError(error);
        this.stopCurrentProcess(lastError);
      }
    }
    const detail = lastError?.message ?? "unknown connection error";
    this.emitStatus({ state: "failed", detail });
    throw lastError ?? new Error(detail);
  }

  private async startProcess(): Promise<void> {
    const child = spawn(
      this.executable,
      [...this.globalArguments, "rpc", this.projectRoot],
      {
      cwd: this.projectRoot,
      stdio: "pipe",
        env: { ...process.env, RUST_LOG: "warn" },
      },
    );
    this.process = child;
    this.stdoutBuffer = "";
    child.stdout.setEncoding("utf8");
    child.stdout.on("data", (chunk: string) => this.acceptStdout(chunk));
    child.stderr.setEncoding("utf8");
    child.stderr.on("data", (chunk: string) => console.error(`[VisualTeX Core] ${chunk}`));
    child.on("error", (error) => this.handleTermination(child, error));
    child.on("exit", (code, signal) => {
      this.handleTermination(
        child,
        new Error(`VisualTeX Core exited (code=${code}, signal=${signal})`),
      );
    });

    const initialized = await this.send<{ protocolVersion: number }>(
      "initialize",
      {},
      10_000,
    );
    if (!Number.isInteger(initialized.protocolVersion)) {
      throw new Error("VisualTeX Core initialize response has no protocolVersion");
    }
  }

  private send<T>(method: string, params: unknown, timeoutMs = DEFAULT_REQUEST_TIMEOUT_MS): Promise<T> {
    const child = this.process;
    if (!child || !this.isRunning()) {
      return Promise.reject(new CoreTransportError("VisualTeX Core is not connected"));
    }
    const id = this.nextId++;
    const payload = JSON.stringify({ jsonrpc: "2.0", id, method, params });
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new CoreTransportError(`VisualTeX Core request timed out: ${method}`));
      }, timeoutMs);
      this.pending.set(id, {
        resolve: resolve as (value: unknown) => void,
        reject,
        timer,
      });
      child.stdin.write(`${payload}\n`, (error) => {
        if (!error) return;
        const pending = this.pending.get(id);
        if (!pending) return;
        clearTimeout(pending.timer);
        this.pending.delete(id);
        pending.reject(new CoreTransportError(error.message));
      });
    });
  }

  private acceptStdout(chunk: string): void {
    this.stdoutBuffer += chunk;
    for (;;) {
      const newline = this.stdoutBuffer.indexOf("\n");
      if (newline < 0) break;
      const line = this.stdoutBuffer.slice(0, newline).trim();
      this.stdoutBuffer = this.stdoutBuffer.slice(newline + 1);
      if (!line) continue;
      try {
        const response = JSON.parse(line) as RpcResponse;
        const pending = this.pending.get(response.id);
        if (!pending) continue;
        clearTimeout(pending.timer);
        this.pending.delete(response.id);
        if (response.error) {
          const detail = response.error.data === undefined
            ? response.error.message
            : `${response.error.message}: ${JSON.stringify(response.error.data)}`;
          pending.reject(new CoreRpcError(detail));
        } else {
          pending.resolve(response.result);
        }
      } catch (error) {
        console.error("Invalid VisualTeX Core response", line, error);
      }
    }
  }

  private handleTermination(child: ChildProcessWithoutNullStreams, error: Error): void {
    if (child !== this.process) return;
    this.process = null;
    this.rejectAll(new CoreTransportError(error.message));
    if (this.disposed) return;
    this.emitStatus({ state: "disconnected", detail: error.message });
    if (this.everConnected) this.scheduleReconnect();
  }

  private scheduleReconnect(): void {
    if (this.reconnectTimer || this.disposed) return;
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      void this.ensureConnected().catch(() => undefined);
    }, 400);
  }

  private stopCurrentProcess(error: Error): void {
    const child = this.process;
    this.process = null;
    this.rejectAll(new CoreTransportError(error.message));
    if (child && child.exitCode === null && !child.killed) child.kill();
  }

  private rejectAll(error: Error): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer);
      pending.reject(error);
    }
    this.pending.clear();
  }

  private isRunning(): boolean {
    return Boolean(this.process && this.process.exitCode === null && !this.process.killed);
  }

  private emitStatus(status: CoreConnectionStatus): void {
    for (const listener of this.statusListeners) listener(status);
  }
}

function wait(delayMs: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, delayMs));
}

function asError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}
