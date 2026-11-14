// Settings → AI agents (MCP): how each MCP client is told to start VisualTeX (a stdio server: the launcher and its
// arguments), with where the text goes. Agent writes can take minutes, so the clients that have a tool timeout get
// 600 seconds.

export type McpClientGuide = { key: string; label: string; where: string; text: string };

export type McpPlatform = "mac" | "windows";

const TOOL_TIMEOUT_SECONDS = 600;

export function mcpClientGuides(launcher: string, args: string[], platform: McpPlatform, isEn: boolean): McpClientGuide[] {
  const t = (en: string, zh: string) => (isEn ? en : zh);
  const windows = platform === "windows";
  const home = (path: string) => (windows ? `%USERPROFILE%\\${path.replace(/\//g, "\\")}` : `~/${path}`);
  const quote = (value: string) => (windows ? `"${value}"` : `'${value.replace(/'/g, `'\\''`)}'`);
  const commandLine = [quote(launcher), ...args].join(" ");
  const json = (value: unknown) => JSON.stringify(value, null, 2);
  // a TOML literal string keeps a Windows path's backslashes as they are
  const toml = (value: string) => (value.includes("'") ? JSON.stringify(value) : `'${value}'`);
  const server = { command: launcher, args };
  return [
    {
      key: "agent",
      label: t("Ask your agent", "让智能体配置"),
      where: t("Send to the agent you use", "发给你使用的智能体"),
      text: t(
        `Add VisualTeX as a stdio MCP server named visualtex: command "${launcher}"${args.length ? `, arguments ${args.join(" ")}` : ""}, tool timeout ${TOOL_TIMEOUT_SECONDS} seconds. Then reload the MCP servers.`,
        `把 VisualTeX 加为 stdio MCP 服务器，名称 visualtex，命令 "${launcher}"${args.length ? `，参数 ${args.join(" ")}` : ""}，工具超时 ${TOOL_TIMEOUT_SECONDS} 秒，然后重新加载 MCP 服务器。`,
      ),
    },
    {
      key: "claude-code",
      label: "Claude Code",
      where: t("Terminal", "终端"),
      text: `claude mcp add --scope user visualtex -- ${commandLine}`,
    },
    {
      key: "claude-desktop",
      label: "Claude Desktop",
      where: t(
        "Settings → Developer → Edit Config (claude_desktop_config.json) → mcpServers",
        "设置 → 开发者 → 编辑配置（claude_desktop_config.json）→ mcpServers",
      ),
      text: json({ mcpServers: { visualtex: server } }),
    },
    {
      key: "codex",
      label: "Codex",
      where: home(".codex/config.toml"),
      text: [
        "[mcp_servers.visualtex]",
        `command = ${toml(launcher)}`,
        `args = [${args.map(toml).join(", ")}]`,
        "startup_timeout_sec = 30",
        `tool_timeout_sec = ${TOOL_TIMEOUT_SECONDS}`,
        "",
      ].join("\n"),
    },
    {
      key: "cursor",
      label: "Cursor",
      where: `${home(".cursor/mcp.json")} → mcpServers`,
      text: json({ mcpServers: { visualtex: server } }),
    },
    {
      key: "vscode",
      label: "VS Code (Copilot)",
      where: t("Command Palette → MCP: Open User Configuration → servers", "命令面板 → MCP: Open User Configuration → servers"),
      text: json({ servers: { visualtex: { type: "stdio", ...server } } }),
    },
    {
      key: "cline",
      label: "Cline",
      where: t("Cline → MCP Servers → Configure → Configure MCP Servers → mcpServers", "Cline → MCP Servers → Configure → Configure MCP Servers → mcpServers"),
      text: json({ mcpServers: { visualtex: { ...server, timeout: TOOL_TIMEOUT_SECONDS, disabled: false } } }),
    },
    {
      key: "windsurf",
      label: "Windsurf",
      where: `${home(".codeium/windsurf/mcp_config.json")} → mcpServers`,
      text: json({ mcpServers: { visualtex: server } }),
    },
    {
      key: "gemini",
      label: "Gemini CLI",
      where: `${home(".gemini/settings.json")} → mcpServers`,
      text: json({ mcpServers: { visualtex: { ...server, timeout: TOOL_TIMEOUT_SECONDS * 1000 } } }),
    },
    {
      key: "opencode",
      label: "opencode",
      where: `${home(".config/opencode/opencode.json")} → mcp`,
      text: json({
        $schema: "https://opencode.ai/config.json",
        mcp: { visualtex: { type: "local", command: [launcher, ...args], enabled: true, timeout: TOOL_TIMEOUT_SECONDS * 1000 } },
      }),
    },
    {
      key: "cherry",
      label: "Cherry Studio",
      where: t("Settings → MCP Servers → Add → Import from JSON", "设置 → MCP 服务器 → 添加 → 从 JSON 导入"),
      // Cherry Studio wants forward slashes in a Windows path
      text: json({
        mcpServers: { visualtex: { type: "stdio", command: launcher.replace(/\\/g, "/"), args, timeout: TOOL_TIMEOUT_SECONDS } },
      }),
    },
    {
      key: "other",
      label: t("Other clients", "其他客户端"),
      where: t(`stdio; tool timeout ${TOOL_TIMEOUT_SECONDS} s or more`, `stdio；工具超时不少于 ${TOOL_TIMEOUT_SECONDS} 秒`),
      text: commandLine,
    },
  ];
}
