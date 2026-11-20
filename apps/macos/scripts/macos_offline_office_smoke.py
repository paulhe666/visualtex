"""Verify the Office artifacts that ship, without depending on implementation text.

Formula/session behavior is covered by the TypeScript and Rust regressions;
Word/PowerPoint interactions still require the real-host acceptance scripts.
"""

import hashlib
import json
import plistlib
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from verify_word_vba_source import verify as verify_word_vba_source


ROOT = Path(__file__).resolve().parents[1]
OFFICE = ROOT / "office/macos-offline"
RESOURCES = OFFICE / "resources"
CALLBACK_ATTRIBUTES = {
    "onLoad", "onAction", "getEnabled", "getVisible", "getLabel",
    "getImage", "getPressed", "getText", "getItemCount", "getItemLabel",
    "getItemID", "getSelectedItemIndex", "getSelectedItemID", "getContent",
}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def xml_structure(element):
    return (
        element.tag,
        sorted(element.attrib.items()),
        (element.text or "").strip(),
        [xml_structure(child) for child in element],
    )


def verify_addin(host, filename, vba_part, document_part, content_type, manifest):
    artifact = RESOURCES / filename
    expected = manifest["files"][filename]["sha256"]
    require(hashlib.sha256(artifact.read_bytes()).hexdigest() == expected,
            f"{filename}: packaged bytes differ from addins.json")

    source_ribbon = ET.parse(OFFICE / host / "customUI14.xml").getroot()
    callbacks = {
        value for element in source_ribbon.iter()
        for name, value in element.attrib.items() if name in CALLBACK_ATTRIBUTES
    }
    public_procedures = set()
    for folder in (OFFICE / "shared", OFFICE / host):
        for source in folder.glob("*.bas"):
            public_procedures.update(re.findall(
                r"^Public\s+(?:Sub|Function)\s+(\w+)",
                source.read_text(), re.MULTILINE | re.IGNORECASE,
            ))
    require(callbacks <= public_procedures,
            f"{host}: Ribbon callbacks missing from VBA: {callbacks - public_procedures}")

    with zipfile.ZipFile(artifact) as package:
        require(package.testzip() is None, f"{filename}: corrupt ZIP entry")
        types = ET.fromstring(package.read("[Content_Types].xml"))
        require(any(entry.get("PartName") == document_part and
                    entry.get("ContentType") == content_type for entry in types),
                f"{filename}: incorrect Office document type")
        ribbon = ET.fromstring(package.read("customUI/customUI14.xml"))
        require(xml_structure(ribbon) == xml_structure(source_ribbon),
                f"{filename}: packaged Ribbon differs from source")
        relationships = ET.fromstring(package.read("_rels/.rels"))
        require(any(entry.get("Target", "").lstrip("/") == "customUI/customUI14.xml"
                    and entry.get("Type", "").endswith("/ui/extensibility")
                    for entry in relationships), f"{filename}: Ribbon is not connected")
        vba = package.read(vba_part)
        for callback in callbacks:
            require(callback.encode() in vba or callback.encode("utf-16le") in vba,
                    f"{filename}: compiled VBA is missing Ribbon callback {callback}")

    print(f"PASS {filename}: checksum, document type, Ribbon and {len(callbacks)} callbacks")


def main():
    manifest = json.loads((RESOURCES / "addins.json").read_text())
    version = json.loads((ROOT / "package.json").read_text())["version"]
    require(manifest["schemaVersion"] == 1, "Unsupported add-in manifest schema")
    require(manifest["pluginVersion"] == version, "Add-in and application versions differ")
    for args in [
        ("word", "VisualTeX.dotm", "word/vbaProject.bin", "/word/document.xml",
         "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"),
        ("powerpoint", "VisualTeX.ppam", "ppt/vbaProject.bin", "/ppt/presentation.xml",
         "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"),
    ]:
        verify_addin(*args, manifest)

    verify_word_vba_source(RESOURCES / "VisualTeX.dotm")

    config = json.loads((ROOT / "src-tauri/tauri.macos.conf.json").read_text())
    resources = config["bundle"]["resources"]
    for relative in resources:
        require((ROOT / "src-tauri" / relative).exists(), f"Missing bundle resource: {relative}")
    for name in ("VisualTeX.dotm", "VisualTeX.ppam", "addins.json"):
        require(f"office/macos-offline/resources/{name}" in resources.values(),
                f"{name} is not bundled")

    with (ROOT / "src-tauri/Info.macos.plist").open("rb") as file:
        plist = plistlib.load(file)
    require(any("visualtex" in entry.get("CFBundleURLSchemes", [])
                for entry in plist.get("CFBundleURLTypes", [])), "Missing visualtex URL scheme")

    if sys.platform == "darwin":
        with tempfile.TemporaryDirectory(prefix="visualtex-office-smoke-") as temp:
            for host, name in [("word", "VisualTeXWord"), ("powerpoint", "VisualTeXPowerPoint")]:
                subprocess.run([
                    "/usr/bin/osacompile", "-l", "AppleScript", "-o",
                    str(Path(temp) / f"{name}.scpt"), str(OFFICE / host / f"{name}.scpt"),
                ], check=True, capture_output=True, timeout=30)
        print("PASS both AppleScriptTask sources compile")
    else:
        print("SKIP AppleScript compilation: requires macOS")
    print("VisualTeX macOS Office artifact smoke: PASS")


if __name__ == "__main__":
    main()
