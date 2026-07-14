import { OfficeBridge } from "../bridge/OfficeBridge";
import { WindowsOleAdapter } from "./WindowsOleAdapter";
import type { OfficeHost } from "../shared/sessionClient";

/** Windows-only Office.js bridge. It never imports the macOS adapters. */
export class WindowsOleBridge extends OfficeBridge {
  private readonly windowsAdapter: WindowsOleAdapter;

  constructor(host: OfficeHost) {
    const adapter = new WindowsOleAdapter(host);
    super(adapter);
    this.windowsAdapter = adapter;
  }

  async updateEquationNumbers() {
    return this.windowsAdapter.updateEquationNumbers();
  }
}
