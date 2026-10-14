import "mathlive";

declare module "mathlive" {
  /**
   * Added by vite.mathliveSemanticCompletion.ts to VisualTeX's MathLive kernel:
   * every command the kernel understands as [name, required args, optional args].
   */
  export function visualTexLatexCommandSignatures(): Array<[string, number, number]>;
}
