import { Component, type ErrorInfo, type ReactNode } from "react";

interface Props {
  children: ReactNode;
}

interface State {
  error: Error | null;
}

export class VisualTexErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("VisualTeX web editor crashed", error, info);
  }

  render() {
    const error = this.state.error;
    if (!error) return this.props.children;

    return (
      <main
        role="alert"
        style={{
          minHeight: "100vh",
          boxSizing: "border-box",
          display: "grid",
          placeContent: "center",
          gap: 14,
          padding: 32,
          background: "#f5f6f8",
          color: "#881337",
          fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        }}
      >
        <strong style={{ fontSize: 18 }}>VisualTeX 网页编辑器加载失败</strong>
        <p style={{ margin: 0, maxWidth: 760, color: "#4b5563" }}>
          当前浏览器在加载编辑器时遇到异常。重新加载后若仍出现此页面，请把下面的错误信息发送给开发者。
        </p>
        <pre
          style={{
            boxSizing: "border-box",
            maxWidth: "min(900px, calc(100vw - 64px))",
            maxHeight: "45vh",
            margin: 0,
            padding: 14,
            overflow: "auto",
            border: "1px solid #fecdd3",
            borderRadius: 8,
            background: "#fff1f2",
            color: "#9f1239",
            whiteSpace: "pre-wrap",
            overflowWrap: "anywhere",
            font: "12px/1.55 ui-monospace, SFMono-Regular, Menlo, monospace",
          }}
        >
          {error.stack || error.message || String(error)}
        </pre>
        <button
          type="button"
          onClick={() => window.location.reload()}
          style={{
            justifySelf: "start",
            padding: "8px 14px",
            border: "1px solid #be123c",
            borderRadius: 7,
            background: "#be123c",
            color: "white",
            font: "600 13px -apple-system, BlinkMacSystemFont, sans-serif",
            cursor: "pointer",
          }}
        >
          重新加载
        </button>
      </main>
    );
  }
}
