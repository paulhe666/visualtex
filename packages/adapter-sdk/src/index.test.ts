import { mkdtemp, mkdir, realpath, rm, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:net";
import { tmpdir } from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import {
  BRIDGE_PROTOCOL_VERSION,
  VisualTexAdapterClient,
  VisualTexRpcError,
  buildVisualTexUri,
  parseVisualTexUri,
} from "./index.js";

const temporaryDirectories: string[] = [];
const servers: Server[] = [];

afterEach(async () => {
  await Promise.all(servers.splice(0).map((server) => new Promise<void>((resolve) => {
    server.close(() => resolve());
  })));
  await Promise.all(temporaryDirectories.splice(0).map((directory) =>
    rm(directory, { recursive: true, force: true })));
});

async function createProject(endpoint: string, tokenFileOverride?: string): Promise<{
  projectRoot: string;
  token: string;
}> {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "visualtex-adapter-论文-"));
  temporaryDirectories.push(projectRoot);
  const bridgeDirectory = path.join(projectRoot, ".visualtex", "bridge");
  await mkdir(bridgeDirectory, { recursive: true });
  const token = "a".repeat(64);
  const tokenFile = tokenFileOverride ?? path.join(bridgeDirectory, "token.txt");
  await writeFile(tokenFile, token, "utf8");
  await writeFile(path.join(bridgeDirectory, "session.json"), JSON.stringify({
    bridgeProtocolVersion: BRIDGE_PROTOCOL_VERSION,
    projectRoot: await realpath(projectRoot),
    endpoint,
    tokenFile,
    pid: process.pid,
    startedUnixMs: Date.now(),
  }), "utf8");
  return { projectRoot, token };
}

async function startFakeBridge(token = "a".repeat(64)): Promise<{
  endpoint: string;
  methods: string[];
}> {
  const methods: string[] = [];
  const server = createServer((socket) => {
    let buffer = "";
    socket.setEncoding("utf8");
    socket.on("data", (chunk: string) => {
      buffer += chunk;
      const newline = buffer.indexOf("\n");
      if (newline < 0) return;
      const envelope = JSON.parse(buffer.slice(0, newline)) as {
        bridgeProtocolVersion: number;
        token: string;
        request: { id: string; method: string; params: unknown };
      };
      methods.push(envelope.request.method);
      const error = envelope.token !== token
        ? { code: -32010, message: "bad token" }
        : envelope.request.method === "fail"
          ? { code: -32000, message: "expected failure", data: { field: "value" } }
          : undefined;
      const result = envelope.request.method === "initialize"
        ? { protocolVersion: 1, capabilities: { compile: true } }
        : envelope.request.method === "project.dependencies"
          ? {
              edges: [{
                sourceFile: "main.tex",
                targetFile: "chapter.tex",
                rawPath: "chapter",
                kind: "input",
                startByte: 10,
                endByte: 25,
                resolved: true,
              }],
              cycles: [],
            }
          : envelope.request.method === "project.compile"
            ? { status: "succeeded" }
          : envelope.request.method === "synctex.forwardSearch"
            ? { pdfPath: "build/main.pdf", boxes: [{ page: 1, x: 1, y: 2, width: 3, height: 4 }] }
            : envelope.request.method === "synctex.inverseSearch"
              ? { sourcePath: "main.tex", line: 8, column: 1 }
              : { ok: true, params: envelope.request.params };
      socket.end(JSON.stringify({
        bridgeProtocolVersion: 1,
        response: {
          jsonrpc: "2.0",
          id: envelope.request.id,
          ...(error ? { error } : { result }),
        },
      }) + "\n");
    });
  });
  servers.push(server);
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("missing fake bridge address");
  return { endpoint: `127.0.0.1:${address.port}`, methods };
}

describe("VisualTexAdapterClient", () => {
  it("discovers an authenticated bridge and wraps compile/SyncTeX", async () => {
    const bridge = await startFakeBridge();
    const project = await createProject(bridge.endpoint);
    const client = await VisualTexAdapterClient.connect(project.projectRoot);
    await expect(client.projectDependencies()).resolves.toMatchObject({
      edges: [{ sourceFile: "main.tex", targetFile: "chapter.tex", resolved: true }],
      cycles: [],
    });
    await expect(client.compile<{ status: string }>()).resolves.toEqual({ status: "succeeded" });
    await expect(client.forwardSearch("main.tex", 8, 1, "build/main.pdf"))
      .resolves.toMatchObject({ boxes: [{ page: 1 }] });
    await expect(client.inverseSearch("build/main.pdf", 1, 10, 20))
      .resolves.toEqual({ sourcePath: "main.tex", line: 8, column: 1 });
    expect(bridge.methods).toEqual([
      "initialize",
      "project.dependencies",
      "project.refreshFromDisk",
      "project.compile",
      "project.refreshFromDisk",
      "synctex.forwardSearch",
      "synctex.inverseSearch",
    ]);
  });

  it("preserves typed JSON-RPC errors", async () => {
    const bridge = await startFakeBridge();
    const project = await createProject(bridge.endpoint);
    const client = await VisualTexAdapterClient.connect(project.projectRoot);
    await expect(client.request("fail")).rejects.toMatchObject({
      name: "VisualTexRpcError",
      code: -32000,
      data: { field: "value" },
    } satisfies Partial<VisualTexRpcError>);
  });

  it("rejects non-loopback discovery and token paths outside the project", async () => {
    const nonLoopback = await createProject("192.0.2.1:1234");
    await expect(VisualTexAdapterClient.connect(nonLoopback.projectRoot))
      .rejects.toThrow("not loopback-only");

    const bridge = await startFakeBridge();
    const outside = await mkdtemp(path.join(tmpdir(), "visualtex-token-outside-"));
    temporaryDirectories.push(outside);
    const outsideToken = path.join(outside, "token.txt");
    const escaped = await createProject(bridge.endpoint, outsideToken);
    await expect(VisualTexAdapterClient.connect(escaped.projectRoot))
      .rejects.toThrow("token escapes");
  });
});

describe("visualtex URI", () => {
  it("round-trips Unicode open and SyncTeX actions", () => {
    const open = { kind: "open" as const, project: "/tmp/论文 项目" };
    expect(parseVisualTexUri(buildVisualTexUri(open))).toEqual(open);

    const forward = {
      kind: "forwardSearch" as const,
      project: "C:\\论文 项目",
      sourceFile: "章节 一.tex",
      line: 18,
      column: 3,
      pdfPath: ".visualtex\\build\\main.pdf",
    };
    expect(parseVisualTexUri(buildVisualTexUri(forward))).toEqual(forward);
  });

  it("rejects unknown actions and protocol versions", () => {
    expect(() => parseVisualTexUri("visualtex://unknown?v=1&project=x"))
      .toThrow("Unsupported visualtex URI action");
    expect(() => parseVisualTexUri("visualtex://open?v=2&project=x"))
      .toThrow("Unsupported visualtex URI version");
  });
});
