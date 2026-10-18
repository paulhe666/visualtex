"""Verify Word-saved inline image formula positions against the Windows rule.

The rule (WindowsOffice.Contracts WordInlineAlignment.CalculateFontPositionForHost,
mirrored by VTWordInlineImagePosition in the Word add-in):
  auto / baseline paragraphs: Position = -(descent below the TeX baseline),
      rounded to whole points; the baseline then sits on the text baseline.
  center / top paragraphs:    Position = 0.
  bottom (FarEast50):         Position = -round(max(0, height - font size) / 2).

usage: python3 scripts/word_inline_baseline_docx_regression.py SAVED.docx
"""
import json, sys, zipfile, xml.etree.ElementTree as E
from inspect_word_conversion_docx import decode_metadata, NS, W
P = '{' + NS['wp'] + '}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def round_half_up(value):
    return int(value + 0.5)


def semantic_size(value):
    size = value if value and value > 0 else 14.0
    return round_half_up(min(200.0, max(5.0, size)) * 100) / 100


def paragraph_alignment(paragraph, styles, default_style_id, default_alignment):
    style_id = default_style_id
    ppr = paragraph.find(W + 'pPr')
    if ppr is not None:
        own = ppr.find(W + 'textAlignment')
        if own is not None:
            return own.get(W + 'val')
        style = ppr.find(W + 'pStyle')
        if style is not None:
            style_id = style.get(W + 'val')
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        style = styles.get(style_id)
        if style is None:
            break
        own = style.find(W + 'pPr/' + W + 'textAlignment')
        if own is not None:
            return own.get(W + 'val')
        based = style.find(W + 'basedOn')
        style_id = based.get(W + 'val') if based is not None else None
    return default_alignment


rows = []
with zipfile.ZipFile(sys.argv[1]) as z:
    root = E.fromstring(z.read('word/document.xml'))
    styles, default_style_id, default_alignment = {}, None, 'auto'
    if 'word/styles.xml' in z.namelist():
        style_root = E.fromstring(z.read('word/styles.xml'))
        for style in style_root.iter(W + 'style'):
            if style.get(W + 'type') == 'paragraph':
                styles[style.get(W + 'styleId')] = style
                if style.get(W + 'default') == '1' and default_style_id is None:
                    default_style_id = style.get(W + 'styleId')
        default = style_root.find(W + 'docDefaults/' + W + 'pPrDefault/' + W + 'pPr/' + W + 'textAlignment')
        if default is not None:
            default_alignment = default.get(W + 'val')
    parents = {c: p for p in root.iter() for c in p}
    for drawing in root.iter(W + 'drawing'):
        props = drawing.find('.//' + P + 'docPr'); ext = drawing.find('.//' + P + 'extent')
        if props is None or not props.get('title', '').startswith('visualtex:formula-ref:v1:'):
            continue
        meta = decode_metadata(props.get('descr', ''))
        if meta['displayMode'] != 'inline':
            continue
        crop = drawing.find('.//' + A + 'srcRect')
        if crop is not None and any(int(crop.get(k, '0')) < 0 for k in ('t', 'b')):
            raise AssertionError(f"{meta['lines']}: picture still carries a transparent crop canvas")
        run = drawing
        while run.tag != W + 'r':
            run = parents[run]
        paragraph = run
        while paragraph.tag != W + 'p':
            paragraph = parents[paragraph]
        alignment = paragraph_alignment(paragraph, styles, default_style_id, default_alignment)
        pos = run.find('./' + W + 'rPr/' + W + 'position')
        position = 0 if pos is None else float(pos.get(W + 'val')) / 2
        height = int(ext.get('cy')) / 12700
        descent = -meta['referenceBaselinePt'] * height / meta['referenceHeightPt']
        if alignment in ('top', 'center'):
            expected = 0
        elif alignment == 'bottom':
            expected = -round_half_up(max(0.0, height - semantic_size(meta.get('fontSizePt'))) * 0.5)
        else:
            expected = -round_half_up(descent)
        assert position == expected, (meta['lines'], alignment, position, expected, height, descent)
        error = position + descent if alignment in ('auto', 'baseline') else None
        if error is not None:
            assert abs(error) <= 0.5 + 0.06, (meta['lines'], error)
        rows.append({'latex': '\n'.join(x['latex'] for x in meta['lines']), 'alignment': alignment,
                     'fontSizePt': meta['fontSizePt'], 'positionPt': position, 'descentPt': round(descent, 3),
                     'baselineErrorPt': None if error is None else round(error, 3)})
assert rows, 'No real Word inline formulas found'
errors = [abs(x['baselineErrorPt']) for x in rows if x['baselineErrorPt'] is not None]
report = {'count': len(rows), 'maxBaselineErrorPt': max(errors) if errors else None, 'formulas': rows}
print(json.dumps(report, ensure_ascii=False, indent=2))
