import { randomUUID } from "node:crypto";
import { readFile, realpath } from "node:fs/promises";
import { createConnection } from "node:net";
import * as path from "node:path";

export const BRIDGE_PROTOCOL_VERSION = 1;
export const MAX_BRIDGE_LINE_BYTES = 1024 * 1024;

export interface BridgeDiscovery {
  bridgeProtocolVersion: number;
  projectRoot: string;
  endpoint: string;
  tokenFile: string;
  pid: number;
  startedUnixMs: number;
}

export interface CoreCapabilities {
  incrementalEdits: boolean;
  incrementalSyntaxTree?: boolean;
  projectDependencyGraph?: boolean;
  visualEdits: boolean;
  compile: boolean;
  toolchainDetection: boolean;
  undoRedo: boolean;
  pdfiumRendering: boolean;
  pdfTiles: boolean;
  shadowLayoutMap: boolean;
  formulaOcr?: boolean;
  documentOcr?: boolean;
  offlineOnly?: boolean;
}

export interface InitializeResult {
  protocolVersion: number;
  capabilities: CoreCapabilities;
}

export interface RpcErrorShape {
  code: number;
  message: string;
  data?: unknown;
}

interface RpcResponse<T> {
  jsonrpc: "2.0";
  id: string | number | null;
  result?: T;
  error?: RpcErrorShape;
}

interface BridgeResponse<T> {
  bridgeProtocolVersion: number;
  response: RpcResponse<T>;
}

export interface AdapterClientOptions {
  connectTimeoutMs?: number;
  requestTimeoutMs?: number;
}

export interface ForwardSearchResult {
  pdfPath: string;
  boxes: Array<{
    page: number;
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
}

export interface InverseSearchResult {
  sourcePath: string;
  line: number;
  column: number | null;
}

export type DependencyKind = "input" | "include" | "subfile" | "subfile_include";

export interface ProjectDependencyGraph {
  edges: Array<{
    sourceFile: string;
    targetFile: string;
    rawPath: string;
    kind: DependencyKind;
    startByte: number;
    endByte: number;
    resolved: boolean;
  }>;
  cycles: string[][];
}

export class VisualTexRpcError extends Error {
  constructor(
    public readonly code: number,
    message: string,
    public readonly data?: unknown,
  ) {
    super(message);
    this.name = "VisualTexRpcError";
  }
}

export class VisualTexAdapterClient {
  private constructor(
    public readonly discovery: BridgeDiscovery,
    private readonly token: string,
    private readonly connectTimeoutMs: number,
    private readonly requestTimeoutMs: number,
  ) {}

  static async connect(
    projectRoot: string,
    options: AdapterClientOptions = {},
  ): Promise<VisualTexAdapterClient> {
    const canonicalProject = await realpath(projectRoot);
    const bridgeDirectory = await realpath(path.join(canonicalProject, ".visualtex", "bridge"));
    const discoveryPath = path.join(bridgeDirectory, "session.json");
    const discovery = parseDiscovery(JSON.parse(await readFile(discoveryPath, "utf8")));
    if (discovery.bridgeProtocolVersion !== BRIDGE_PROTOCOL_VERSION) {
      throw new Error(
        `Unsupported VisualTeX bridge protocol ${discovery.bridgeProtocolVersion}; expected ${BRIDGE_PROTOCOL_VERSION}`,
      );
    }
    const discoveredProject = await realpath(discovery.projectRoot);
    if (discoveredProject !== canonicalProject) {
      throw new Error("VisualTeX bridge discovery belongs to another project");
    }
    validateEndpoint(discovery.endpoint);
    const tokenPath = await realpath(discovery.tokenFile);
    if (!isPathInside(bridgeDirectory, tokenPath)) {
      throw new Error("VisualTeX bridge token escapes the project bridge directory");
    }
    const token = (await readFile(tokenPath, "utf8")).trim();
    if (!/^[0-9a-f]{64}$/i.test(token)) {
      throw new Error("VisualTeX bridge token has an invalid format");
    }
    const client = new VisualTexAdapterClient(
      discovery,
      token,
      options.connectTimeoutMs ?? 5_000,
      options.requestTimeoutMs ?? 180_000,
    );
    const initialized = await client.initialize();
    if (initialized.protocolVersion !== BRIDGE_PROTOCOL_VERSION) {
      throw new Error(
        `Unsupported VisualTeX Core protocol ${initialized.protocolVersion}; expected ${BRIDGE_PROTOCOL_VERSION}`,
      );
    }
    return client;
  }

  request<T>(method: string, params: unknown = {}): Promise<T> {
    if (!method || method.length > 256) throw new Error("Invalid VisualTeX RPC method");
    const id = randomUUID();
    const line = JSON.stringify({
      bridgeProtocolVersion: BRIDGE_PROTOCOL_VERSION,
      token: this.token,
      request: { jsonrpc: "2.0", id, method, params },
    }) + "\n";
    if (Buffer.byteLength(line, "utf8") > MAX_BRIDGE_LINE_BYTES) {
      throw new Error("VisualTeX bridge request exceeds the 1 MiB limit");
    }
    const { port } = validateEndpoint(this.discovery.endpoint);
    return new Promise<T>((resolve, reject) => {
      const socket = createConnection({ host: "127.0.0.1", port });
      let settled = false;
      let responseBuffer = Buffer.alloc(0);
      const connectTimer = setTimeout(() => {
        fail(new Error("VisualTeX bridge connection timed out"));
      }, this.connectTimeoutMs);
      const requestTimer = setTimeout(() => {
        fail(new Error(`VisualTeX bridge request timed out: ${method}`));
      }, this.requestTimeoutMs);

      const cleanup = () => {
        clearTimeout(connectTimer);
        clearTimeout(requestTimer);
        socket.removeAllListeners();
        socket.destroy();
      };
      const fail = (error: Error) => {
        if (settled) return;
        settled = true;
        cleanup();
        reject(error);
      };
      const succeed = (value: T) => {
        if (settled) return;
        settled = true;
        cleanup();
        resolve(value);
      };

      socket.once("connect", () => {
        clearTimeout(connectTimer);
        socket.write(line, "utf8", (error) => {
          if (error) fail(error);
        });
      });
      socket.on("data", (chunk: Buffer) => {
        responseBuffer = Buffer.concat([responseBuffer, chunk]);
        if (responseBuffer.length > MAX_BRIDGE_LINE_BYTES) {
          fail(new Error("VisualTeX bridge response exceeds the 1 MiB limit"));
          return;
        }
        const newline = responseBuffer.indexOf(0x0a);
        if (newline < 0) return;
        try {
          const envelope = JSON.parse(responseBuffer.subarray(0, newline).toString("utf8")) as BridgeResponse<T>;
          if (envelope.bridgeProtocolVersion !== BRIDGE_PROTOCOL_VERSION) {
            throw new Error("VisualTeX bridge response protocol version mismatch");
          }
          if (!envelope.response || envelope.response.id !== id) {
            throw new Error("VisualTeX bridge response id mismatch");
          }
          if (envelope.response.error) {
            const error = envelope.response.error;
            throw new VisualTexRpcError(error.code, error.message, error.data);
          }
          succeed(envelope.response.result as T);
        } catch (error) {
          fail(error instanceof Error ? error : new Error(String(error)));
        }
      });
      socket.once("error", fail);
      socket.once("end", () => {
        if (!settled) fail(new Error("VisualTeX bridge closed without a response"));
      });
    });
  }

  initialize(): Promise<InitializeResult> {
    return this.request("initialize", { protocolVersion: BRIDGE_PROTOCOL_VERSION });
  }

  rootSnapshot<T = unknown>(): Promise<T> {
    return this.request("project.rootSnapshot");
  }

  projectDependencies(): Promise<ProjectDependencyGraph> {
    return this.request("project.dependencies");
  }

  refreshFromDisk<T = unknown>(): Promise<T> {
    return this.request("project.refreshFromDisk");
  }

  async compile<T = unknown>(): Promise<T> {
    await this.refreshFromDisk();
    return this.request("project.compile");
  }

  async forwardSearch(
    sourceFile: string,
    line: number,
    column: number,
    pdfPath: string,
  ): Promise<ForwardSearchResult> {
    await this.refreshFromDisk();
    return this.request("synctex.forwardSearch", {
      sourceFile,
      line,
      column,
      pdfPath,
    });
  }

  inverseSearch(
    pdfPath: string,
    page: number,
    x: number,
    y: number,
  ): Promise<InverseSearchResult> {
    return this.request("synctex.inverseSearch", { pdfPath, page, x, y });
  }

  shutdown(): Promise<{ ok: boolean }> {
    return this.request("bridge.shutdown");
  }
}

export type VisualTexUriAction =
  | { kind: "open"; project: string }
  | {
      kind: "forwardSearch";
      project: string;
      sourceFile: string;
      line: number;
      column: number;
      pdfPath: string;
    }
  | {
      kind: "inverseSearch";
      project: string;
      pdfPath: string;
      page: number;
      x: number;
      y: number;
    };

export function buildVisualTexUri(action: VisualTexUriAction): string {
  const host = action.kind === "open"
    ? "open"
    : action.kind === "forwardSearch"
      ? "forward-search"
      : "inverse-search";
  const uri = new URL(`visualtex://${host}`);
  uri.searchParams.set("v", String(BRIDGE_PROTOCOL_VERSION));
  uri.searchParams.set("project", action.project);
  if (action.kind === "forwardSearch") {
    uri.searchParams.set("source", action.sourceFile);
    uri.searchParams.set("line", String(action.line));
    uri.searchParams.set("column", String(action.column));
    uri.searchParams.set("pdf", action.pdfPath);
  } else if (action.kind === "inverseSearch") {
    uri.searchParams.set("pdf", action.pdfPath);
    uri.searchParams.set("page", String(action.page));
    uri.searchParams.set("x", String(action.x));
    uri.searchParams.set("y", String(action.y));
  }
  return uri.toString();
}

export function parseVisualTexUri(value: string): VisualTexUriAction {
  const uri = new URL(value);
  if (uri.protocol !== "visualtex:") throw new Error("Not a visualtex URI");
  const version = requiredNumber(uri, "v", true);
  if (version !== BRIDGE_PROTOCOL_VERSION) {
    throw new Error(`Unsupported visualtex URI version ${version}`);
  }
  const project = requiredString(uri, "project");
  if (uri.hostname === "open") return { kind: "open", project };
  if (uri.hostname === "forward-search") {
    return {
      kind: "forwardSearch",
      project,
      sourceFile: requiredString(uri, "source"),
      line: requiredNumber(uri, "line", true),
      column: requiredNumber(uri, "column", true),
      pdfPath: requiredString(uri, "pdf"),
    };
  }
  if (uri.hostname === "inverse-search") {
    return {
      kind: "inverseSearch",
      project,
      pdfPath: requiredString(uri, "pdf"),
      page: requiredNumber(uri, "page", true),
      x: requiredNumber(uri, "x", false),
      y: requiredNumber(uri, "y", false),
    };
  }
  throw new Error(`Unsupported visualtex URI action: ${uri.hostname}`);
}

function parseDiscovery(value: unknown): BridgeDiscovery {
  if (!value || typeof value !== "object") throw new Error("Invalid VisualTeX bridge discovery");
  const record = value as Record<string, unknown>;
  return {
    bridgeProtocolVersion: integer(record.bridgeProtocolVersion, "bridgeProtocolVersion"),
    projectRoot: string(record.projectRoot, "projectRoot"),
    endpoint: string(record.endpoint, "endpoint"),
    tokenFile: string(record.tokenFile, "tokenFile"),
    pid: integer(record.pid, "pid"),
    startedUnixMs: number(record.startedUnixMs, "startedUnixMs"),
  };
}

function validateEndpoint(endpoint: string): { port: number } {
  const match = /^127\.0\.0\.1:(\d{1,5})$/.exec(endpoint);
  if (!match) throw new Error("VisualTeX bridge endpoint is not loopback-only");
  const port = Number(match[1]);
  if (!Number.isInteger(port) || port < 1 || port > 65_535) {
    throw new Error("VisualTeX bridge endpoint has an invalid port");
  }
  return { port };
}

function isPathInside(parent: string, child: string): boolean {
  const relative = path.relative(parent, child);
  return relative.length > 0 && !relative.startsWith("..") && !path.isAbsolute(relative);
}

function requiredString(uri: URL, name: string): string {
  const value = uri.searchParams.get(name);
  if (!value) throw new Error(`Missing visualtex URI parameter: ${name}`);
  return value;
}

function requiredNumber(uri: URL, name: string, integerOnly: boolean): number {
  const value = Number(requiredString(uri, name));
  if (!Number.isFinite(value) || (integerOnly && !Number.isInteger(value))) {
    throw new Error(`Invalid visualtex URI number: ${name}`);
  }
  return value;
}

function string(value: unknown, name: string): string {
  if (typeof value !== "string" || value.length === 0 || value.length > 32_768) {
    throw new Error(`Invalid ${name}`);
  }
  return value;
}

function number(value: unknown, name: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(`Invalid ${name}`);
  return value;
}

function integer(value: unknown, name: string): number {
  const parsed = number(value, name);
  if (!Number.isInteger(parsed)) throw new Error(`Invalid ${name}`);
  return parsed;
}
