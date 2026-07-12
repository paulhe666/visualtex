import { describe, expect, it } from "vitest";
import {
  commandRegistry,
  createMatrixCommand,
  searchCommands,
  templateForSelection,
} from "./commands";

describe("VisualTeX 1.0.6 formula command behavior", () => {
  it("wraps the current selection instead of discarding it", () => {
    const fraction = commandRegistry.find((command) => command.id === "frac")!;
    const root = commandRegistry.find((command) => command.id === "sqrt")!;
    const scripts = commandRegistry.find((command) => command.id === "scripts")!;

    expect(templateForSelection(fraction, "a+b")).toBe("\\frac{a+b}{\\placeholder{}}");
    expect(templateForSelection(root, "x^2+y^2")).toBe("\\sqrt{x^2+y^2}");
    expect(templateForSelection(scripts, "X")).toBe("X_{\\placeholder{}}^{\\placeholder{}}");
  });

  it("searches commands by LaTeX, English aliases and Chinese labels", () => {
    expect(searchCommands("frac", 3)[0]?.id).toBe("frac");
    expect(searchCommands("matrix", 5).some((command) => command.category === "matrix")).toBe(true);
    expect(searchCommands("偏导", 5).some((command) => command.id === "partial")).toBe(true);
  });

  it("creates bounded custom matrices with editable placeholders", () => {
    const matrix = createMatrixCommand(3, 4, "pmatrix");
    expect(matrix.id).toBe("custom-pmatrix-3x4");
    expect(matrix.insertTemplate).toContain("\\begin{pmatrix}");
    expect(matrix.insertTemplate.match(/\\placeholder\{\}/g)).toHaveLength(12);
    expect(matrix.insertTemplate).toContain("\\end{pmatrix}");
  });
});
