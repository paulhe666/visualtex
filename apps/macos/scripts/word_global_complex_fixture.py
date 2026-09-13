"""Reuse verified Word scaffold topology with freshly rendered formula content."""
import base64
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from inspect_word_conversion_docx import NS, W

template, assets, output = map(Path, sys.argv[1:])
fixtures = json.loads((assets / "fixtures.json").read_text())
with zipfile.ZipFile(template) as archive:
    files = {name: archive.read(name) for name in archive.namelist()}
root = ET.fromstring(files["word/document.xml"])
old_ids = [d.get("title").split(":")[3] for d in root.findall(".//wp:docPr", NS)]
for old, fixture in zip(old_ids, fixtures):
    new = fixture["metadata"]["formulaId"]
    for name, data in files.items():
        if name.endswith(".xml"):
            for a, b in [(old, new), (old.replace("-", ""), new.replace("-", "")),
                         (old.replace("-", "_"), new.replace("-", "_"))]:
                data = data.replace(a.encode(), b.encode())
            files[name] = data
root = ET.fromstring(files["word/document.xml"])
settings = ET.fromstring(files["word/settings.xml"])
rels = ET.fromstring(files["word/_rels/document.xml.rels"])
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P = "{http://schemas.openxmlformats.org/package/2006/relationships}"
docvars = settings.find(W + "docVars")
for variable in list(docvars):
    if variable.get(W + "name", "").startswith(("VT_Latex_", "VT_Metadata_", "VT_OMML_", "VT_NativeSignature_", "VT_ImageScale_", "VT_Format_")):
        docvars.remove(variable)
for index, (drawing, fixture) in enumerate(zip(root.findall(".//w:drawing", NS), fixtures)):
    meta = fixture["metadata"]
    formula_id = meta["formulaId"]
    props = drawing.find(".//wp:docPr", NS)
    props.set("descr", fixture["encodedMetadata"])
    props.set("title", f"visualtex:formula-ref:v1:{formula_id}:{meta['displayMode']}:{int(meta['numbered'])}")
    for element in drawing.iter():
        if element.tag.endswith("}extent") or element.tag.endswith("}ext") and "cx" in element.attrib:
            element.set("cx", str(round(fixture["width"] * 12700)))
            element.set("cy", str(round(fixture["height"] * 12700)))
        if R + "embed" in element.attrib:
            extension = "svg" if element.tag.endswith("}svgBlip") else "png"
            rel_id = f"rIdComplex{index}{extension}"
            element.set(R + "embed", rel_id)
            target = f"media/complex-{index}.{extension}"
            ET.SubElement(rels, P + "Relationship", Id=rel_id, Type=R[1:-1] + "/image", Target=target)
            files["word/" + target] = (assets / f"{index}.{extension}").read_bytes()
    # Size only the visible formula run. Hidden SEQ helpers retain their geometry.
    for run in root.findall(".//w:r", NS):
        if drawing in list(run):
            props_run = run.find(W + "rPr")
            if props_run is None:
                props_run = ET.SubElement(run, W + "rPr")
            size = props_run.find(W + "sz")
            if size is None:
                size = ET.SubElement(props_run, W + "sz")
            size.set(W + "val", str(round(meta["fontSizePt"] * 2)))
    stem = formula_id.replace("-", "_")
    values = {f"VT_Format_{stem}": f"{meta['displayMode']}|{int(meta['numbered'])}",
              f"VT_ImageScale_{stem}": "|".join(str(meta[k]) for k in ["fontSizePt", "referenceWidthPt", "referenceHeightPt", "referenceBaselinePt", "fontSizePt"]),
              f"VT_Latex_{stem}_Count": "1", f"VT_Metadata_{stem}_Count": "1",
              f"VT_Latex_{stem}_001": base64.urlsafe_b64encode(meta["latex"].encode()).decode().rstrip("="),
              f"VT_Metadata_{stem}_001": fixture["encodedMetadata"]}
    for name, value in values.items():
        ET.SubElement(docvars, W + "docVar", {W + "name": name, W + "val": value})
for name, element in [("word/document.xml", root), ("word/settings.xml", settings), ("word/_rels/document.xml.rels", rels)]:
    # ElementTree changes namespace prefixes; mc:Ignorable contains lexical
    # prefixes from the seed. Remove this optional list so Word never sees
    # undeclared prefixes after serialization.
    for node in element.iter():
        for attr in list(node.attrib):
            if attr.endswith("}Ignorable"):
                del node.attrib[attr]
    files[name] = ET.tostring(element, encoding="utf-8", xml_declaration=True)
with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
    for name, data in files.items():
        archive.writestr(name, data)
print(output)
