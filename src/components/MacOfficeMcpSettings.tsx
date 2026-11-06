import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { Copy, History, KeyRound, RefreshCw, ShieldAlert, ToggleLeft, ToggleRight } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { mcpClientGuides } from "./mcpClientGuides";

// Settings → Office integration → AI agents (MCP): docs/office-mcp/MACOS.md. A client starts the VisualTeX
// executable with --mcp-stdio, which relays to the companion's /mcp.

type McpMode = "readOnly" | "allow";

type McpStatus = {
  settings: { enabled: boolean; mode: McpMode };
  tokenPath: string;
  launcherPath: string;
  launcherArgs: string[];
  hosts: { hostId: string; app: string; productVersion: string }[];
  clients: { sessionId: string; name: string; version: string; connectedAt: string }[];
};

type AuditEntry = {
  time: string;
  client: string;
  tool: string;
  document?: string;
  ok: boolean;
  error?: string;
  durationMs: number;
};

function errorMessage(error: unknown) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return "";
}

function localTime(value: string) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

export function MacOfficeMcpSettings() {
  const isEn = useEditorStore((state) => state.language) === "en";
  const [status, setStatus] = useState<McpStatus | null>(null);
  const [audit, setAudit] = useState<AuditEntry[] | null>(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [copied, setCopied] = useState<string | null>(null);
  const [clientKey, setClientKey] = useState("agent");

  const refresh = useCallback(async () => {
    try {
      setStatus(await invoke<McpStatus>("get_mcp_status"));
      setMessage("");
    } catch (error) {
      setMessage(errorMessage(error));
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const run = async (command: string, args?: Record<string, unknown>) => {
    setBusy(true);
    setMessage("");
    try {
      setStatus(await invoke<McpStatus>(command, args));
    } catch (error) {
      setMessage(errorMessage(error));
    } finally {
      setBusy(false);
    }
  };

  const save = (enabled: boolean, mode: McpMode) => run("set_mcp_settings", { settings: { enabled, mode } });

  const copy = async (key: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopied(key);
      window.setTimeout(() => setCopied((value) => (value === key ? null : value)), 1600);
    } catch (error) {
      setMessage(errorMessage(error));
    }
  };

  const toggleAudit = async () => {
    if (audit) {
      setAudit(null);
      return;
    }
    try {
      setAudit(await invoke<AuditEntry[]>("get_mcp_audit", { limit: 100 }));
    } catch (error) {
      setMessage(errorMessage(error));
    }
  };

  const enabled = Boolean(status?.settings.enabled);
  const mode = status?.settings.mode ?? "allow";
  const launcher = status?.launcherPath ?? "";
  const args = status?.launcherArgs ?? ["--mcp-stdio"];
  const guides = mcpClientGuides(launcher, args, "mac", isEn);
  const guide = guides.find((item) => item.key === clientKey) ?? guides[0];

  return (
    <section className="settings-section office-integration-section office-mcp-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "AI agents (MCP)" : "AI 智能体（MCP）"}</strong>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy}
          aria-label={isEn ? "Refresh" : "刷新"}
          title={isEn ? "Refresh" : "刷新"}
        >
          <RefreshCw size={15} />
        </button>
      </div>

      <div className="office-mcp-row">
        <strong>{isEn ? "Agent access" : "智能体访问"}</strong>
        <button
          type="button"
          className="secondary-button office-mcp-toggle"
          disabled={busy || !status}
          aria-pressed={enabled}
          onClick={() => void save(!enabled, mode)}
        >
          {enabled ? <ToggleRight size={16} /> : <ToggleLeft size={16} />}
          {enabled ? (isEn ? "Enabled" : "已开启") : isEn ? "Disabled" : "已关闭"}
        </button>
      </div>

      {enabled && (
        <>
          <div className="office-mcp-row">
            <strong>{isEn ? "Permission" : "权限"}</strong>
            <select
              className="office-mcp-select"
              value={mode}
              disabled={busy}
              onChange={(event) => void save(true, event.target.value as McpMode)}
            >
              <option value="allow">{isEn ? "Read and change" : "读取和修改"}</option>
              <option value="readOnly">{isEn ? "Read only" : "只读"}</option>
            </select>
          </div>

          <div className="office-mcp-clients">
            <div className="office-mcp-client-head">
              <strong>{isEn ? "Connect a client" : "连接客户端"}</strong>
              <select
                className="office-mcp-select"
                value={guide.key}
                aria-label={isEn ? "Client" : "客户端"}
                onChange={(event) => setClientKey(event.target.value)}
              >
                {guides.map((item) => (
                  <option key={item.key} value={item.key}>
                    {item.label}
                  </option>
                ))}
              </select>
            </div>
            <div className="office-mcp-client-where">
              <span>{guide.where}</span>
              <button
                type="button"
                className="secondary-button"
                disabled={!launcher}
                onClick={() => void copy(guide.key, guide.text)}
              >
                <Copy size={15} />
                {copied === guide.key ? (isEn ? "Copied" : "已复制") : isEn ? "Copy" : "复制"}
              </button>
            </div>
            <pre className={guide.key === "agent" ? "office-mcp-snippet is-prose" : "office-mcp-snippet"}>
              {launcher ? guide.text : "—"}
            </pre>
          </div>

          <dl className="office-mcp-connections">
            <dt>{isEn ? "Office" : "Office 应用"}</dt>
            <dd>
              {status?.hosts.length
                ? status.hosts
                    .map((host) => `${host.app === "word" ? "Word" : "PowerPoint"} ${host.productVersion}`.trim())
                    .join(isEn ? ", " : "、")
                : "—"}
            </dd>
            <dt>{isEn ? "Agents" : "智能体"}</dt>
            <dd>
              {status?.clients.length
                ? status.clients.map((client) => `${client.name} ${client.version}`.trim()).join(isEn ? ", " : "、")
                : "—"}
            </dd>
          </dl>

          <div className="office-mcp-actions">
            <button type="button" className="secondary-button" onClick={() => void toggleAudit()}>
              <History size={15} />
              {isEn ? "Activity" : "活动记录"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy}
              title={status?.tokenPath}
              onClick={() => void run("reset_mcp_access")}
            >
              <KeyRound size={15} />
              {isEn ? "Reset access" : "重置访问"}
            </button>
          </div>

          {audit && (
            <ol className="office-mcp-audit">
              {audit.length === 0 && <li>—</li>}
              {audit.map((entry, index) => (
                <li key={`${entry.time}-${index}`} className={entry.ok ? undefined : "is-error"}>
                  <span>{localTime(entry.time)}</span>
                  <span>{entry.client}</span>
                  <code>{entry.tool}</code>
                  <span>{entry.ok ? `${entry.durationMs} ms` : entry.error}</span>
                </li>
              ))}
            </ol>
          )}
        </>
      )}

      {message && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>{message}</span>
        </div>
      )}
    </section>
  );
}
