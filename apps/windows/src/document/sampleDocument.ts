// A small document for the dev page and the tests.
export const sampleDocument = {
  type: "doc",
  content: [
    { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Harmonic oscillator" }] },
    {
      type: "paragraph",
      content: [
        { type: "text", text: "The energy of a mass " },
        { type: "math_inline", attrs: { latex: "m" } },
        { type: "text", text: " on a spring of stiffness " },
        { type: "math_inline", attrs: { latex: "k" } },
        { type: "text", text: " is " },
        { type: "math_inline", attrs: { latex: "E=\\frac{1}{2}mv^2+\\frac{1}{2}kx^2" } },
        { type: "text", text: ", and its frequency is" },
      ],
    },
    { type: "math_display", attrs: { latex: "\\omega=\\sqrt{\\frac{k}{m}}", numbered: true } },
    { type: "paragraph", content: [{ type: "text", text: "A table of values:" }] },
    {
      type: "table",
      content: [
        {
          type: "table_row",
          content: [
            { type: "table_header", content: [{ type: "paragraph", content: [{ type: "text", text: "Quantity" }] }] },
            { type: "table_header", content: [{ type: "paragraph", content: [{ type: "text", text: "Formula" }] }] },
          ],
        },
        {
          type: "table_row",
          content: [
            { type: "table_cell", content: [{ type: "paragraph", content: [{ type: "text", text: "Period" }] }] },
            {
              type: "table_cell",
              content: [{ type: "paragraph", content: [{ type: "math_inline", attrs: { latex: "T=2\\pi\\sqrt{m/k}" } }] }],
            },
          ],
        },
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "End." }] },
  ],
};
