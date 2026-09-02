#!/usr/bin/env python3
"""Verify the final Office package, not just the separately bundled Ribbon source."""

from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
UI_REL = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
UI_NS = "http://schemas.microsoft.com/office/2009/07/customui"
UI_PART = "customUI/customUI14.xml"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def target_part(relationship, base):
    require(relationship.get("TargetMode") != "External", "Ribbon relationships must be internal")
    target = relationship.get("Target", "")
    require(bool(target), "Ribbon relationship has no target")
    path = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(base, target))
    require(path != ".." and not path.startswith("../"), "Ribbon target leaves the package")
    return path


def verify(root, host, filename, main_part, main_type):
    path = root / "resources" / filename
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        require(len(names) == len(set(names)), "Duplicate OOXML entries")
        for part in ["[Content_Types].xml", "_rels/.rels", f"{host}/vbaProject.bin", main_part, UI_PART]:
            require(part in names, f"Missing {part}; package the compiled VBA with the reviewed Ribbon before distributing")
        content_types = ET.fromstring(package.read("[Content_Types].xml"))
        defaults = {entry.get("Extension"): entry.get("ContentType") for entry in content_types.findall(f"{{{CT_NS}}}Default")}
        overrides = {entry.get("PartName"): entry.get("ContentType") for entry in content_types.findall(f"{{{CT_NS}}}Override")}
        require(overrides.get("/" + main_part) == main_type, "Wrong Office add-in main content type")
        require(overrides.get("/" + UI_PART, defaults.get("xml")) == "application/xml", "Ribbon XML content type is missing")

        relationships = ET.fromstring(package.read("_rels/.rels"))
        ui_links = [entry for entry in relationships.findall(f"{{{REL_NS}}}Relationship") if entry.get("Type") == UI_REL]
        require(len(ui_links) == 1 and target_part(ui_links[0], "") == UI_PART, "Package root does not link to the reviewed Ribbon")
        ribbon_bytes = package.read(UI_PART)
        source_host = "word" if host == "word" else "powerpoint"
        require(ribbon_bytes.strip() == (root / source_host / "customUI14.xml").read_bytes().strip(), "Packaged Ribbon differs from the reviewed source")
        ribbon = ET.fromstring(ribbon_bytes)
        require(ribbon.tag == f"{{{UI_NS}}}customUI", "Wrong Ribbon namespace")
        require(any(tab.get("label") == "VisualTeX" for tab in ribbon.iter(f"{{{UI_NS}}}tab")), "VisualTeX tab is missing")

        project = package.read(f"{host}/vbaProject.bin")
        for element in ribbon.iter():
            for attribute in ["onLoad", "onAction"]:
                callback = element.get(attribute)
                if callback:
                    require(callback.encode() in project or callback.encode("utf-16le") in project, f"Missing compiled Ribbon callback {callback}")

        images = {element.get("image") for element in ribbon.iter() if element.get("image")}
        if images:
            rel_part = "customUI/_rels/customUI14.xml.rels"
            require(rel_part in names, "Ribbon image relationships are missing")
            image_rels = ET.fromstring(package.read(rel_part))
            by_id = {entry.get("Id"): entry for entry in image_rels.findall(f"{{{REL_NS}}}Relationship")}
            for image in images:
                require(image in by_id, f"Ribbon image {image} has no relationship")
                relation = by_id[image]
                require(relation.get("Type") == IMAGE_REL, f"Wrong relationship type for {image}")
                part = target_part(relation, "customUI")
                require(part in names, f"Ribbon image is missing: {part}")
                require(package.read(part).startswith(b"\x89PNG\r\n\x1a\n"), f"Ribbon image is not PNG: {part}")
                require(overrides.get("/" + part, defaults.get("png")) == "image/png", f"PNG content type is missing: {part}")
    print(f"{filename}: Ribbon, callbacks and {len(images)} images PASS")


def main():
    root = Path(sys.argv[1])
    for args in [
        ("word", "VisualTeX.dotm", "word/document.xml", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"),
        ("ppt", "VisualTeX.ppam", "ppt/presentation.xml", "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"),
    ]:
        try:
            verify(root, *args)
        except (ValueError, KeyError, OSError, zipfile.BadZipFile, ET.ParseError) as error:
            raise SystemExit(f"{args[1]}: {error}") from error


if __name__ == "__main__":
    main()
