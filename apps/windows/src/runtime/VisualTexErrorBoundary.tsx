import { Component, type ErrorInfo, type ReactNode } from "react";
import { prepareEditorCrashSafeRestart } from "./editorCrashRecovery";

interface Props {
  children: ReactNode;
  title?: ReactNode;
  description?: ReactNode;
  reloadLabel?: ReactNode;
  logContext?: string;
  enableSafeRestart?: boolean;
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
    console.error(
      this.props.logContext ?? "VisualTeX main window crashed",
      error,
      info,
    );
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
          fontFamily:
            "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        }}
      >
        <strong style={{ fontSize: 18 }}>
          {this.props.title ?? "VisualTeX 界面加载失败"}
        </strong>
        <p style={{ margin: 0, maxWidth: 760, color: "#4b5563" }}>
          {this.props.description ??
            "应用已经启动，但当前 WebView 在加载界面时遇到了异常。重新打开后若仍出现此页面，请把下面的错误信息发送给开发者。"}
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
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button
            type="button"
            onClick={() => window.location.reload()}
            style={{
              padding: "8px 14px",
              border: "1px solid #be123c",
              borderRadius: 7,
              background: "#be123c",
              color: "white",
              font: "600 13px -apple-system, BlinkMacSystemFont, sans-serif",
              cursor: "pointer",
            }}
          >
            {this.props.reloadLabel ?? "重新加载"}
          </button>
          {this.props.enableSafeRestart !== false ? (
            <button
              type="button"
              onClick={() => {
                try {
                  prepareEditorCrashSafeRestart();
                  window.location.reload();
                } catch (reason) {
                  console.error(
                    "VisualTeX safe restart preparation failed",
                    reason,
                  );
                  window.alert(
                    `无法准备安全启动，原有数据未被清除。${reason instanceof Error ? reason.message : String(reason)}`,
                  );
                }
              }}
              style={{
                padding: "8px 14px",
                border: "1px solid #9f1239",
                borderRadius: 7,
                background: "white",
                color: "#9f1239",
                font: "600 13px -apple-system, BlinkMacSystemFont, sans-serif",
                cursor: "pointer",
              }}
            >
              安全启动（保留设置）
            </button>
          ) : null}
        </div>
      </main>
    );
  }
}
