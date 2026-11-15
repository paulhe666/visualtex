"""Check the VBA source embedded in a compiled DOTM against its build profile.

Read decompressed module source: stale names can survive in VBA's binary symbol
table after a procedure is removed, so a byte-string search is insufficient.
Every module of the Word add-in (office/macos-offline/shared and word) must be
in the project with the procedures of its source, no other standard or class
module may be there, and every module with a diagnostic flag must have the
profile's value (a diagnostic build composes the tests into VTWordAdapter).
"""

import argparse
from pathlib import Path
import re

from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
OFFICE = ROOT / "office/macos-offline"
DIAGNOSTICS = ROOT / "tests/office/word/VTWordAdapterDiagnostics.bas.inc"
PROCEDURE = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.MULTILINE | re.IGNORECASE,
)
FLAG = re.compile(r"^#Const\s+VT_WORD_DIAGNOSTICS\s*=\s*(\S+)", re.MULTILINE | re.IGNORECASE)


def procedure_names(source):
    return {name.lower() for name in PROCEDURE.findall(source)}


def expected_modules(diagnostics):
    modules = {}
    for folder in (OFFICE / "shared", OFFICE / "word"):
        for path in sorted(folder.glob("*.bas")) + sorted(folder.glob("*.cls")):
            source = path.read_text()
            if diagnostics and path.name == "VTWordAdapter.bas":
                source += DIAGNOSTICS.read_text()
            modules[path.name.lower()] = source
    return modules


def verify(path, diagnostics=False):
    expected = expected_modules(diagnostics)
    parser = VBA_Parser(str(path))
    try:
        actual = {name.lower(): source for _, _, name, source in parser.extract_macros()}
    finally:
        parser.close()
    # Word's document module (ThisDocument) belongs to every template
    actual = {name: source for name, source in actual.items() if name != "thisdocument.cls"}

    missing_modules = sorted(set(expected) - set(actual))
    extra_modules = sorted(set(actual) - set(expected))
    if missing_modules or extra_modules:
        raise ValueError(f"Word VBA modules differ: missing={missing_modules}, extra={extra_modules}")
    for name, source in expected.items():
        missing = procedure_names(source) - procedure_names(actual[name])
        extra = procedure_names(actual[name]) - procedure_names(source)
        if missing or extra:
            raise ValueError(f"Word VBA profile mismatch in {name}: missing={sorted(missing)}, extra={sorted(extra)}")
        if FLAG.search(source):
            flag = FLAG.search(actual[name])
            expected_values = {"true", "-1"} if diagnostics else {"false", "0"}
            if not flag or flag[1].lower() not in expected_values:
                raise ValueError(f"Compiled Word VBA module {name} has the wrong diagnostic build flag")
    if not FLAG.search(actual["vtwordadapter.bas"]):
        raise ValueError("Compiled VTWordAdapter has no diagnostic build flag")
    procedures = sum(len(procedure_names(source)) for source in actual.values())
    print(f"Word VBA {'diagnostic' if diagnostics else 'production'} profile: {len(actual)} modules, "
          f"{procedures} procedures verified")


if __name__ == "__main__":
    args = argparse.ArgumentParser(description=__doc__)
    args.add_argument("artifact", type=Path)
    args.add_argument("--diagnostics", action="store_true")
    options = args.parse_args()
    verify(options.artifact, options.diagnostics)
