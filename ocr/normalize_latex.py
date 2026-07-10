from __future__ import annotations

import re

_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def normalize_formula_latex(value: str) -> str:
    """Conservatively normalize OCR LaTeX without changing math semantics."""
    if not isinstance(value, str):
        raise TypeError("Formula LaTeX must be a string")

    latex = _CONTROL_CHARS.sub("", value).strip()

    wrappers = (
        ("$$", "$$"),
        (r"\[", r"\]"),
        (r"\(", r"\)"),
    )
    for start, end in wrappers:
        if latex.startswith(start) and latex.endswith(end):
            latex = latex[len(start) : len(latex) - len(end)].strip()
            break

    # Preserve LaTeX matrix/array row separators while flattening physical lines.
    latex = re.sub(r"(?<!\\)\r?\n", " ", latex)
    latex = re.sub(r"[ \t]+", " ", latex).strip()
    return latex
