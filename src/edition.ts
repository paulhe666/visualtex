// The build's edition. "personal": the author's own Mac and Windows builds, where document mode has LaTeX (input,
// source, import, export, the TeX-rendered LaTeX documents). "public": releases given to others and the web
// version, where document mode is Markdown only. Set at build time: VITE_VISUALTEX_EDITION=personal.
export type Edition = "personal" | "public";

export const EDITION: Edition = import.meta.env.VITE_VISUALTEX_EDITION === "personal" ? "personal" : "public";

/** Document mode's LaTeX (personal edition only). */
export const DOCUMENT_LATEX = EDITION === "personal";
