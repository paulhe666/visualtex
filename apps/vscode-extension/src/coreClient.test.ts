import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { CoreClient, type CoreConnectionStatus } from "./coreClient";

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map((directory) =>
    rm(directory, {
      recursive: true,
      force: true,
      maxRetries: 10,
      retryDelay: 50,
    })));
});

async function fakeCoreExecutable(): Promise<{
  executable: string;
  globalArguments: string[];
  root: string;
}> {
  const root = await mkdtemp(path.join(tmpdir(), "visualtex-core-client-"));
  temporaryDirectories.push(root);
  const marker = path.join(root, "unstable-marker");
  const executable = path.join(root, "fake-core.js");
  const script = `#!/usr/bin/env node
const fs = require("node:fs");
const readline = require("node:readline");
const marker = ${JSON.stringify(marker)};
const input = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
function reply(id, result) { process.stdout.write(JSON.stringify({ jsonrpc: "2.0", id, result }) + "\\n"); }
input.on("line", (line) => {
  const request = JSON.parse(line);
  if (request.method === "initialize") {
    reply(request.id, { protocolVersion: 1, capabilities: {} });
  } else if (request.method === "echo") {
    reply(request.id, request.params);
  } else if (request.method === "unstable") {
    if (!fs.existsSync(marker)) {
      fs.writeFileSync(marker, "exited once");
      process.exit(17);
    }
    reply(request.id, { recovered: true });
  } else if (request.method === "fail") {
    process.stdout.write(JSON.stringify({ jsonrpc: "2.0", id: request.id, error: { code: -32000, message: "expected failure" } }) + "\\n");
  } else {
    reply(request.id, null);
  }
});
`;
  await writeFile(executable, script, "utf8");
  return {
    executable: process.execPath,
    globalArguments: [executable],
    root,
  };
}

describe("CoreClient", () => {
  it("connects, performs the initialize handshake, and returns RPC results", async () => {
    const fake = await fakeCoreExecutable();
    const statuses: CoreConnectionStatus[] = [];
    const client = new CoreClient(fake.executable, fake.root, fake.globalArguments);
    const subscription = client.onStatus((status) => statuses.push(status));
    try {
      await client.connect();
      await expect(client.request("echo", { value: "中文😀" })).resolves.toEqual({
        value: "中文😀",
      });
      expect(statuses.some((status) => status.state === "connected")).toBe(true);
    } finally {
      subscription.dispose();
      client.dispose();
    }
  });

  it("restarts and retries only when the caller explicitly allows it", async () => {
    const fake = await fakeCoreExecutable();
    const client = new CoreClient(fake.executable, fake.root, fake.globalArguments);
    try {
      await client.connect();
      await expect(client.request(
        "unstable",
        {},
        { retryAfterReconnect: true, timeoutMs: 5_000 },
      )).resolves.toEqual({ recovered: true });
    } finally {
      client.dispose();
    }
  });

  it("surfaces JSON-RPC errors without retrying them", async () => {
    const fake = await fakeCoreExecutable();
    const client = new CoreClient(fake.executable, fake.root, fake.globalArguments);
    try {
      await client.connect();
      await expect(client.request("fail", {}, { retryAfterReconnect: true }))
        .rejects.toThrow("expected failure");
    } finally {
      client.dispose();
    }
  });
});
