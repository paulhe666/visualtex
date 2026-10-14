import "mathlive";
import type { CommandUsage } from "../../src/types/command";

export interface MathLivePersistentTypingStyle {
  bold: boolean;
  // null follows MathLive's normal math alphabet; false forces upright.
  italic: boolean | null;
  color: string | null;
  backgroundColor: string | null;
}

declare module "mathlive" {
  interface MathfieldElement {
    visualTexCompletionPreferences?: {
      usage: Record<string, CommandUsage>;
      personalize: boolean;
      count: number;
    };
    visualTexPersistentTypingStyle?: MathLivePersistentTypingStyle;
    visualTexInsertToolbarTemplate(latex: string, options?: Parameters<MathfieldElement["insert"]>[1]): boolean;
    visualTexSelectRange(anchor: number, focus: number): void;
    visualTexInsertDocumentEdges(prefix: string, suffix: string): void;
    visualTexStopPointerTracking(): void;
    visualTexStartPointerSelection(event: PointerEvent): void;
    readonly visualTexParentEnvironment: string | null;
  }
}
