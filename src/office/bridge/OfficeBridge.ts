import { ensureCompanionReady } from "../api/companionClient";
import {
  createOfficeSession,
  getOfficeSession,
  updateOfficeSession,
  type OfficeSessionMode,
} from "../api/sessionClient";
import type { OfficeHostAdapter } from "../adapters/OfficeHostAdapter";
import { DialogController } from "./DialogController";
import type { VisualTeXDialogMessage } from "./bridgeMessages";

function sessionHasFormula(lines: Array<{ latex: string }>) {
  return lines.some((line) => line.latex.trim().length > 0);
}

export class OfficeBridge {
  private readonly dialog = new DialogController();
  private activeSessionId: string | null = null;
  private commandRunning = false;

  constructor(private readonly adapter: OfficeHostAdapter) {}

  async run(mode: OfficeSessionMode) {
    if (this.commandRunning || this.dialog.isOpen) {
      this.adapter.showMessage("VisualTeX 编辑窗口已经打开。");
      return;
    }

    this.commandRunning = true;
    try {
      this.adapter.showMessage("正在连接 VisualTeX 本地伴侣服务…");
      await ensureCompanionReady();
      const selection = await this.adapter.readSelection(mode);
      const session = await createOfficeSession({
        mode,
        host: this.adapter.host,
        sourceDocumentId: selection.sourceDocumentId,
        sourceObjectId: selection.sourceObjectId,
        autoCommitOnClose: true,
        ...selection.sessionSeed,
      });
      this.activeSessionId = session.id;
      this.adapter.showMessage("正在打开 VisualTeX 编辑器…");
      await this.dialog.open(session.id, {
        onMessage: (message) => this.handleDialogMessage(message),
        onClosed: () => this.handleDialogClosed(session.id),
      });
      this.adapter.showMessage("VisualTeX 编辑器已打开。");
    } catch (error) {
      this.adapter.showMessage(
        error instanceof Error ? error.message : "无法启动 VisualTeX Office 编辑器。",
      );
      this.activeSessionId = null;
    } finally {
      this.commandRunning = false;
    }
  }

  async openDesktopApp() {
    try {
      await this.adapter.openDesktopApp();
    } catch (error) {
      this.adapter.showMessage(
        error instanceof Error ? error.message : "无法打开 VisualTeX.app。",
      );
    }
  }

  private async handleDialogMessage(message: VisualTeXDialogMessage) {
    if (message.sessionId !== this.activeSessionId) return;

    if (message.type === "visualtex-ready") {
      await updateOfficeSession(message.sessionId, { status: "editing" });
      return;
    }

    if (message.type === "visualtex-cancel") {
      await updateOfficeSession(message.sessionId, {
        status: "cancelled",
        explicitCancel: true,
      });
      this.dialog.close();
      this.activeSessionId = null;
      this.adapter.showMessage("已取消，Office 文档未修改。");
      return;
    }

    if (message.type === "visualtex-error") {
      await updateOfficeSession(message.sessionId, {
        status: "failed",
        error: message.message,
      });
      this.adapter.showMessage(message.message);
      return;
    }

    if (message.type === "visualtex-commit") {
      await this.commitSession(message.sessionId, true);
    }
  }

  private async handleDialogClosed(sessionId: string) {
    if (sessionId !== this.activeSessionId) return;
    this.activeSessionId = null;

    try {
      const session = await getOfficeSession(sessionId);
      if (
        session.status === "completed" ||
        session.status === "cancelled" ||
        session.explicitCancel
      ) {
        return;
      }

      const shouldAutoCommit =
        session.autoCommitOnClose &&
        sessionHasFormula(session.lines) &&
        Boolean(session.exportResult) &&
        (session.mode === "create" || session.dirty);

      if (shouldAutoCommit) {
        await this.commitSession(sessionId, false);
        return;
      }

      if (!sessionHasFormula(session.lines)) {
        await updateOfficeSession(sessionId, {
          status: "cancelled",
          explicitCancel: false,
        });
        this.adapter.showMessage("空公式已取消，Office 文档未修改。");
      } else if (!session.exportResult) {
        await updateOfficeSession(sessionId, {
          status: "failed",
          error: "公式导出尚未成功，已保留恢复记录。",
        });
        this.adapter.showMessage("公式未插入：导出失败，Session 已保留以便恢复。");
      }
    } catch (error) {
      this.adapter.showMessage(
        error instanceof Error
          ? `无法处理关闭事件：${error.message}`
          : "无法处理 VisualTeX 编辑窗口关闭事件。",
      );
    }
  }

  private async commitSession(sessionId: string, closeAfterSuccess: boolean) {
    try {
      const session = await getOfficeSession(sessionId);
      if (session.status === "cancelled" || session.explicitCancel) return;
      if (!sessionHasFormula(session.lines)) {
        await updateOfficeSession(sessionId, { status: "cancelled" });
        this.adapter.showMessage("空公式没有插入 Office 文档。");
        if (closeAfterSuccess) this.dialog.close();
        return;
      }
      if (!session.exportResult) {
        throw new Error("公式 SVG 尚未生成，无法写入 Office 文档。");
      }
      if (session.mode === "edit" && !session.dirty) {
        await updateOfficeSession(sessionId, { status: "completed" });
        if (closeAfterSuccess) this.dialog.close();
        this.activeSessionId = null;
        this.adapter.showMessage("公式内容未变化，无需更新。");
        return;
      }

      await updateOfficeSession(sessionId, { status: "committing", error: null });
      await this.adapter.applySession(session);
      await updateOfficeSession(sessionId, { status: "completed", error: null });
      if (closeAfterSuccess) this.dialog.close();
      this.activeSessionId = null;
      this.adapter.showMessage(
        session.mode === "edit" ? "VisualTeX 公式已更新。" : "VisualTeX 公式已插入。",
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : "Office 公式写入失败。";
      await updateOfficeSession(sessionId, {
        status: "failed",
        error: message,
      }).catch(() => undefined);
      this.adapter.showMessage(`${message} Session 已保留以便恢复。`);
    }
  }
}
