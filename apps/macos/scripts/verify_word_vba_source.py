"""Check the VBA source embedded in a compiled DOTM against its build profile.

Read decompressed module source: stale names can survive in VBA's binary symbol
table after a procedure is removed, so a byte-string search is insufficient.
"""

import argparse
from pathlib import Path
import re

from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
PROCEDURE = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.MULTILINE | re.IGNORECASE,
)


def procedure_names(source):
    return {name.lower() for name in PROCEDURE.findall(source)}


def verify(path, diagnostics=False):
    expected = (ROOT / "office/macos-offline/word/VTWordAdapter.bas").read_text()
    if diagnostics:
        expected += (ROOT / "tests/office/word/VTWordAdapterDiagnostics.bas.inc").read_text()
    parser = VBA_Parser(str(path))
    try:
        modules = {name.lower(): source for _, _, name, source in parser.extract_macros()}
    finally:
        parser.close()
    actual = modules.get("vtwordadapter.bas")
    if actual is None:
        raise ValueError("Compiled DOTM does not contain VTWordAdapter.bas")

    missing = procedure_names(expected) - procedure_names(actual)
    extra = procedure_names(actual) - procedure_names(expected)
    if missing or extra:
        raise ValueError(f"Word VBA profile mismatch: missing={sorted(missing)}, extra={sorted(extra)}")
    flag = re.search(r"^#Const\s+VT_WORD_DIAGNOSTICS\s*=\s*(\S+)", actual, re.MULTILINE | re.IGNORECASE)
    expected_values = {"true", "-1"} if diagnostics else {"false", "0"}
    if not flag or flag[1].lower() not in expected_values:
        raise ValueError("Compiled Word VBA has the wrong diagnostic build flag")
    print(f"Word VBA {'diagnostic' if diagnostics else 'production'} profile: {len(procedure_names(actual))} procedures verified")


if __name__ == "__main__":
    args = argparse.ArgumentParser(description=__doc__)
    args.add_argument("artifact", type=Path)
    args.add_argument("--diagnostics", action="store_true")
    options = args.parse_args()
    verify(options.artifact, options.diagnostics)
