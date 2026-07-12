import { describe, expect, it } from "vitest";
import { createWebviewHtml } from "./webview";

function fakeWebview(): Parameters<typeof createWebviewHtml>[0] {
  return { cspSource: "vscode-webview://test" } as Parameters<typeof createWebviewHtml>[0];
}

describe("createWebviewHtml", () => {
  it("contains the structured editor, PDF and OCR panels with a strict CSP", () => {
    const html = createWebviewHtml(fakeWebview());
    expect(html).toContain("id=\"editorPanel\"");
    expect(html).toContain("id=\"pdfPanel\"");
    expect(html).toContain("id=\"ocrPanel\"");
    expect(html).toContain("default-src 'none'");
    expect(html).toContain("vscode-webview://test");
  });

  it("emits syntactically valid JavaScript", () => {
    const html = createWebviewHtml(fakeWebview());
    const script = html.match(/<script nonce="[^"]+">([\s\S]*?)<\/script>/)?.[1];
    expect(script).toBeTruthy();
    expect(() => new Function(script!)).not.toThrow();
  });
});
