function describeBootError(value: unknown): string {
  if (value instanceof Error) return value.stack || value.message;
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function renderBootError(value: unknown) {
  const detail = describeBootError(value) || "Unknown startup error";
  console.error("VisualTeX frontend startup failed", value);
  const root = document.getElementById("root");
  if (!root) return;
  root.innerHTML = "";
  const panel = document.createElement("main");
  panel.setAttribute("role", "alert");
  panel.style.cssText =
    "min-height:100vh;display:grid;place-content:center;gap:12px;padding:32px;" +
    "background:#f5f6f8;color:#9f1239;font:14px/1.6 -apple-system,BlinkMacSystemFont,sans-serif;" +
    "white-space:pre-wrap;overflow:auto";
  const heading = document.createElement("strong");
  heading.textContent = "VisualTeX 前端启动失败";
  const message = document.createElement("code");
  message.textContent = detail;
  panel.append(heading, message);
  root.append(panel);
}

void import("./desktop/main").catch(renderBootError);
