"""Inspect snapshots saved by real Word, including image geometry and native fields."""
import base64
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
import zlib

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
W = "{" + NS["w"] + "}"


def decode_metadata(value):
    payload = value.rsplit(":", 1)[-1]
    raw = base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4))
    if ":deflate:" in value:
        raw = zlib.decompress(raw, -15)
    return json.loads(raw)


def inspect(path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        settings = ET.fromstring(archive.read("word/settings.xml"))
    variables = {v.get(W + "name"): v.get(W + "val")
                 for v in settings.iter(W + "docVar")}
    metadata = {}
    for name, count in variables.items():
        if name.startswith("VT_Metadata_") and name.endswith("_Count"):
            stem = name[:-6]
            value = "".join(variables[f"{stem}_{i:03d}"] for i in range(1, int(count) + 1))
            decoded = decode_metadata(value)
            metadata[decoded["formulaId"]] = decoded

    elements = list(root.iter())
    position = {element: i for i, element in enumerate(elements)}
    ends = {e.get(W + "id"): position[e] for e in root.iter(W + "bookmarkEnd")}
    bookmarks = {e.get(W + "name"): (position[e], ends.get(e.get(W + "id"), position[e]))
                 for e in root.iter(W + "bookmarkStart")}
    maths = root.findall(".//m:oMath", NS)
    images = []
    for drawing in root.findall(".//w:drawing", NS):
        props = drawing.find(".//wp:docPr", NS)
        extent = drawing.find(".//wp:extent", NS)
        if props is None or extent is None:
            continue
        title = props.get("title", "")
        if not title.startswith("visualtex:formula-ref:v1:"):
            continue
        parts = title.split(":")
        images.append({"formulaId": parts[3], "width": int(extent.get("cx")) / 12700,
                       "height": int(extent.get("cy")) / 12700})

    fields, stack = [], []
    for element in elements:
        if element.tag == W + "fldChar":
            kind = element.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"code": "", "result": "", "start": position[element], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                field = stack.pop()
                field["end"] = position[element]
                fields.append(field)
        elif stack and element.tag == W + "instrText" and not stack[-1]["separated"]:
            stack[-1]["code"] += element.text or ""
        elif element.tag in (W + "t", "{" + NS["m"] + "}t"):
            for field in stack:
                if field["separated"]:
                    field["result"] += element.text or ""
    for element in root.iter(W + "fldSimple"):
        fields.append({"code": element.get(W + "instr", ""),
                       "result": "".join(element.itertext()),
                       "start": position[element],
                       "end": max(position[n] for n in element.iter())})
    sequences = sorted([f for f in fields if re.match(r"\s*SEQ\s", f["code"], re.I)], key=lambda f: f["start"])
    for field in sequences:
        field["owner"] = next((name for name, (start, end) in bookmarks.items()
                               if name.startswith("VT_N_") and start <= field["end"] and end >= field["start"]), None)
    formulas = []
    for formula_id, meta in metadata.items():
        image = next((i for i in images if i["formulaId"] == formula_id), None)
        native_range = bookmarks.get("VT_F_" + formula_id.replace("-", ""))
        native = None
        if native_range:
            native = next((m for m in maths if position[m] <= native_range[1] and
                           max(position[e] for e in m.iter()) >= native_range[0]), None)
        number_range = bookmarks.get("VT_R_" + formula_id.replace("-", ""))
        number = ""
        if number_range:
            number = "".join(e.text or "" for e in elements[number_range[0]:number_range[1]]
                             if e.tag in (W + "t", "{" + NS["m"] + "}t"))
        formulas.append({"formulaId": formula_id, "kind": "image" if image else "omml" if native is not None else "missing",
                         "latex": meta.get("latex"), "fontSizePt": meta.get("fontSizePt"),
                         "numbered": meta.get("numbered"), "number": number,
                         "width": image["width"] if image else None, "height": image["height"] if image else None,
                         "referenceWidthPt": meta.get("referenceWidthPt"), "referenceHeightPt": meta.get("referenceHeightPt"),
                         "mathText": "".join(n.text or "" for n in native.findall(".//m:t", NS)) if native is not None else None})
    return {"path": path, "images": len(images), "omml": len(maths),
            "tables": len(root.findall(".//w:tbl", NS)), "formulas": formulas,
            "sequences": sequences, "references": [f for f in fields if re.match(r"\s*REF\s", f["code"], re.I)]}


if __name__ == "__main__":
    print(json.dumps(inspect(sys.argv[1]), ensure_ascii=False, indent=2))
