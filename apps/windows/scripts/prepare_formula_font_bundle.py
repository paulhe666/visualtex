"""Build/verify the app-private, redistributable formula-font bundle.

Only this generated resource directory is written. No Windows font is copied,
installed globally, or redistributed. Runtime and normal packaging are offline:
--prepare is an explicit maintainer operation; the default --check only verifies
pinned bytes. KaTeX's WOFF2 container is decoded without changing glyph tables.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path
import urllib.request
import urllib.error
import time
from concurrent.futures import ThreadPoolExecutor

APP = Path(__file__).resolve().parents[1]
REPO = APP.parents[1]
OUT = APP / "src-tauri/resources/formula-fonts"
MANIFEST = OUT / "manifest.json"
DOWNLOADS = REPO / "engine/target/formula-font-downloads-v1"
MAX_DOWNLOAD = 48 * 1024 * 1024
MIT_MIRROR = "https://mirrors.mit.edu/CTAN"


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def write_exact(path: Path, data: bytes) -> None:
    if path.exists():
        if path.read_bytes() != data:
            raise RuntimeError(f"Refusing to overwrite different font resource: {path}")
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def download(url: str) -> bytes:
    key = DOWNLOADS / (sha(url.encode()) + ".download")
    if key.exists():
        return key.read_bytes()
    request = urllib.request.Request(url, headers={"User-Agent": "VisualTeX-font-bundle-builder/1"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                if response.status != 200:
                    raise RuntimeError(f"Font source HTTP {response.status}: {url}")
                data = response.read(MAX_DOWNLOAD + 1)
            break
        except (urllib.error.URLError, OSError) as error:
            if attempt == 2 or isinstance(error, urllib.error.HTTPError) and error.code == 404:
                raise RuntimeError(f"Could not fetch licensed font resource {url}: {error}") from error
            time.sleep(attempt + 1)
    if not data or len(data) > MAX_DOWNLOAD:
        raise RuntimeError(f"Font source exceeds download budget: {url}")
    write_exact(key, data)
    return data


def check() -> dict:
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    if manifest.get("schema") != "visualtex-private-formula-fonts-v1":
        raise RuntimeError("Unknown formula-font manifest schema")
    records = manifest["fonts"] + manifest["licenses"]
    seen = set()
    for item in records:
        name = item["file"]
        if name in seen or Path(name).name != name:
            raise RuntimeError(f"Unsafe/duplicate font resource name: {name}")
        seen.add(name)
        data = (OUT / name).read_bytes()
        if len(data) != item["bytes"] or sha(data) != item["sha256"]:
            raise RuntimeError(f"Formula font integrity check failed: {name}")
    license_files = {x["file"] for x in manifest["licenses"]}
    for font in manifest["fonts"]:
        if font["licenseFile"] not in license_files:
            raise RuntimeError(f"Missing redistribution license: {font['file']}")
        if not font["family"] or not font["glyphs"] or not font["codepoints"]:
            raise RuntimeError(f"Font has no genuine glyph repertoire: {font['file']}")
    if not any(x["id"] == "KaTeX_Main-Regular" for x in manifest["fonts"]):
        raise RuntimeError("The real KaTeX default font is missing")
    return {"status": "PASS", "fonts": len(manifest["fonts"]),
            "licenses": len(license_files), "bytes": sum(x["bytes"] for x in manifest["fonts"]),
            "manifestSha256": sha(MANIFEST.read_bytes()), "runtimeNetworkRequired": False,
            "systemFontInstallationRequired": False, "windowsFontsRedistributed": False}


def prepare() -> dict:
    if MANIFEST.exists():
        return check()
    from fontTools.ttLib import TTFont
    import fontTools
    fonts = []
    licenses = []

    def add_license(name: str, data: bytes, source: str) -> None:
        text = data.decode("utf-8", errors="strict")
        if not any(marker in text for marker in ["MIT License", "Permission is hereby granted",
                                                  "GUST Font License", "SIL OPEN FONT LICENSE"]):
            raise RuntimeError(f"Unrecognized font license: {name}")
        write_exact(OUT / name, data)
        licenses.append({"file": name, "bytes": len(data), "sha256": sha(data), "source": source})

    def add_font(name: str, data: bytes, source: str, license_name: str,
                 source_hash: str | None = None, transformation: str = "none") -> None:
        font = TTFont(io.BytesIO(data), recalcTimestamp=False)
        family = font["name"].getDebugName(16) or font["name"].getDebugName(1)
        subfamily = font["name"].getDebugName(17) or font["name"].getDebugName(2)
        if not family or not font.getBestCmap():
            raise RuntimeError(f"Not a usable OpenType font: {name}")
        # The bundled catalog contains open fonts, never copied OS fonts.
        if family in {"Times New Roman", "Cambria Math", "Arial", "Palatino Linotype", "SimSun", "KaiTi", "SimHei"}:
            raise RuntimeError(f"Proprietary system face must not enter the redistribution bundle: {family}")
        write_exact(OUT / name, data)
        fonts.append({"id": Path(name).stem, "file": name, "family": family,
                      "subfamily": subfamily, "faceIndex": 0,
                      "weight": font["OS/2"].usWeightClass,
                      "italic": bool(font["head"].macStyle & 2),
                      "hasMathTable": "MATH" in font,
                      "bytes": len(data), "sha256": sha(data),
                      "glyphs": len(font.getGlyphOrder()), "codepoints": len(font.getBestCmap()),
                      "source": source, "sourceSha256": source_hash or sha(data),
                      "transformation": transformation, "licenseFile": license_name,
                      "version": font["name"].getDebugName(5)})
        font.close()

    # Use the exact KaTeX faces shipped by the project's pinned MathLive oracle,
    # not a look-alike font and not any of its JavaScript/HTML/CSS renderer.
    oracle = APP / "node_modules/mathlive"
    package = json.loads((oracle / "package.json").read_text(encoding="utf-8"))
    if package["version"] != "0.109.2":
        raise RuntimeError("Pinned KaTeX source requires the already-installed MathLive 0.109.2 oracle")
    katex_license_url = "https://raw.githubusercontent.com/KaTeX/katex-fonts/master/LICENSE"
    add_license("KaTeX-LICENSE.txt", download(katex_license_url), katex_license_url)
    add_license("MathLive-font-distribution-LICENSE.txt", (oracle / "LICENSE.txt").read_bytes(),
                "npm:mathlive@0.109.2/LICENSE.txt")
    originals = sorted((oracle / "fonts").glob("KaTeX_*.woff2"))
    if len(originals) != 20:
        raise RuntimeError(f"Expected the complete 20-face KaTeX font set, got {len(originals)}")
    for path in originals:
        source = path.read_bytes()
        font = TTFont(io.BytesIO(source), recalcTimestamp=False)
        original_tables = {tag: font.getTableData(tag) for tag in ["cmap", "glyf", "hmtx"] if tag in font}
        font.flavor = None
        output = io.BytesIO()
        font.save(output, reorderTables=True)
        font.close()
        data = output.getvalue()
        decoded = TTFont(io.BytesIO(data), recalcTimestamp=False)
        for tag, original in original_tables.items():
            if decoded.getTableData(tag) != original:
                raise RuntimeError(f"WOFF2 decoding changed the {tag} glyph table of {path.name}")
        decoded.close()
        add_font(path.stem + ".ttf", data, f"npm:mathlive@0.109.2/fonts/{path.name}",
                 "KaTeX-LICENSE.txt", sha(source), f"WOFF2 container decoding; fontTools {fontTools.__version__}; cmap/glyf/hmtx byte-identical")

    lm = APP / "src-windows/artifacts/office-fonts"
    add_license("GUST-FONT-LICENSE.txt", (lm / "GUST-FONT-LICENSE.txt").read_bytes(),
                "https://tug.org/fonts/licenses/GUST-FONT-LICENSE.txt")
    add_font("latinmodern-math.otf", (lm / "latinmodern-math.otf").read_bytes(),
             "https://www.gust.org.pl/projects/e-foundry/lm-math", "GUST-FONT-LICENSE.txt")
    stix_root = "https://raw.githubusercontent.com/stipub/stixfonts/v2.13/"
    add_license("STIXTwoMath-OFL.txt", download(stix_root + "OFL.txt"), stix_root + "OFL.txt")
    stix = download(stix_root + "fonts/static_otf/STIXTwoMath-Regular.otf")
    if sha(stix) != "f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72":
        raise RuntimeError("The pinned upstream STIX Two Math 2.13 hash changed")
    add_font("STIXTwoMath-Regular.otf", stix, stix_root + "fonts/static_otf/STIXTwoMath-Regular.otf", "STIXTwoMath-OFL.txt")

    requests = []
    for family in ["termes", "pagella", "heros"]:
        for face in ["regular", "italic", "bold", "bolditalic"]:
            name = f"texgyre{family}-{face}.otf"
            requests.append((name, f"{MIT_MIRROR}/fonts/tex-gyre/opentype/{name}", "GUST-FONT-LICENSE.txt"))
    for name, url, license_name in requests:
        add_font(name, download(url), url, license_name)

    cjk = [
        ("NotoSansCJKsc-Regular.otf", "https://raw.githubusercontent.com/notofonts/noto-cjk/Sans2.004/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf", "NotoSansCJK-OFL.txt", "https://raw.githubusercontent.com/notofonts/noto-cjk/Sans2.004/LICENSE"),
        ("NotoSerifCJKsc-Regular.otf", "https://raw.githubusercontent.com/notofonts/noto-cjk/Serif2.003/Serif/OTF/SimplifiedChinese/NotoSerifCJKsc-Regular.otf", "NotoSerifCJK-OFL.txt", "https://raw.githubusercontent.com/notofonts/noto-cjk/Serif2.003/Serif/LICENSE"),
        ("LXGWWenKai-Regular.ttf", "https://raw.githubusercontent.com/lxgw/LxgwWenKai/v1.520/fonts/TTF/LXGWWenKai-Regular.ttf", "LXGWWenKai-OFL.txt", "https://raw.githubusercontent.com/lxgw/LxgwWenKai/v1.520/OFL.txt"),
    ]
    with ThreadPoolExecutor(max_workers=3) as executor:
        pending = [(spec, executor.submit(download, spec[1]), executor.submit(download, spec[3])) for spec in cjk]
        for (name, url, license_name, license_url), font_future, license_future in pending:
            add_license(license_name, license_future.result(), license_url)
            add_font(name, font_future.result(), url, license_name)
    manifest = {"schema": "visualtex-private-formula-fonts-v1", "defaultProfile": "katex",
                "fonts": sorted(fonts, key=lambda x: x["id"]),
                "licenses": sorted(licenses, key=lambda x: x["file"]),
                "runtimeNetworkRequired": False,
                "proprietaryFonts": "System-local optional access only; redistribution requires a separate grant.",
                "fontSetPolicy": "Faces and glyph outlines are genuine. Explicit mathematical construction/fallback fonts retain their own identities."}
    write_exact(MANIFEST, (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))
    return check()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--prepare", action="store_true", help="Explicitly build the pinned resource bundle; may download upstream fonts")
    parser.add_argument("--check", action="store_true", help="Verify existing resources offline (default)")
    args = parser.parse_args()
    print(json.dumps(prepare() if args.prepare else check(), ensure_ascii=False, indent=2))
