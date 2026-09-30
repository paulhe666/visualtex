"""Compare actual 144-dpi Word pages with the user's reference ink centres.

Usage: python3 scripts/word_chinese_image_alignment_pixels.py RENDER_DIRECTORY
The generator supplies stable ROIs; this reads painted pixels independently of
VBA Position values. Black-only thresholding excludes spelling/Return markers.
"""
import json
import sys
from pathlib import Path
import numpy as np
from PIL import Image, ImageDraw, ImageFont

root = Path(sys.argv[1])
# At 144 dpi Word's integral Position rounds by up to 1 actual pixel. The
# reference/body bounding boxes add raster quantization and use different
# fonts/scales (39 / 32 here). Screen/PDF Asian leading also differs: the Top
# sum has no integral Position that fits both within 2.5 reference pixels.
# Use 3 for both renderers and report the measured errors independently;
# primary letters must still share a baseline within 1 actual pixel.
TOLERANCE = 3
references = [
    ('display', Path('/var/folders/m5/2320d79x2gqg2f_vz_tc5znc0000gn/T/codex-clipboard-ae168f13-1ab3-455a-9056-bf39537d8101.png'), (222, 60, 292, 115), (335, 78, 364, 99)),
    ('text', Path('/var/folders/m5/2320d79x2gqg2f_vz_tc5znc0000gn/T/codex-clipboard-4717ae54-10f7-4224-8c56-c07b776ad9d6.png'), (40, 30, 206, 88), (254, 48, 287, 76)),
    ('sum', Path('/var/folders/m5/2320d79x2gqg2f_vz_tc5znc0000gn/T/codex-clipboard-6ad13082-5b3c-4fe5-baf8-1f4e56588ecf.png'), (26, 45, 234, 100), (420, 61, 454, 85)),
    ('scripts', Path('/var/folders/m5/2320d79x2gqg2f_vz_tc5znc0000gn/T/codex-clipboard-9446716b-b222-42ac-bd19-53521a278e07.png'), (142, 25, 265, 77), (62, 34, 85, 65)),
]

def mask(path):
    return (np.asarray(Image.open(path).convert('RGB')) < 130).all(axis=2)

def bounds(ink, roi):
    x, y, X, Y = roi
    ys, xs = np.where(ink[y:Y, x:X])
    if not len(ys):
        raise AssertionError(f'No painted ink inside {roi}')
    return [int(x + xs.min()), int(y + ys.min()), int(x + xs.max() + 1), int(y + ys.max() + 1)]

def center(box):
    return (box[1] + box[3]) / 2

reference = {}
for kind, path, body_roi, axis_roi in references:
    ink = mask(path)
    body, axis = bounds(ink, body_roi), bounds(ink, axis_roi)
    reference[kind] = {'body': body, 'anchor': axis, 'height': body[3] - body[1], 'delta': center(axis) - center(body)}
reference['scripts']['superscript'] = bounds(mask(references[3][1]), (100, 34, 123, 65))

if '--native' in sys.argv:
    # Screen glyph boxes have their own antialias/leading rounding. One Word
    # point is two screenshot pixels here; reference normalization plus glyph
    # edge rounding gives a three-reference-pixel bound. Keep this separate
    # from the PDF measurements and report both actual maxima.
    native_rows = []
    for mode in ['baseline', 'top', 'center', 'bottom', 'auto']:
        ink = mask(root / f'word-{mode}-r103-native.png')
        groups = [
            ('display', (344, 385, 415, 490), [('fraction', (464, 385, 494, 490)), ('integral', (1196, 385, 1229, 490))]),
            ('text', (144, 500, 290, 595), [('fraction', (333, 500, 363, 595)), ('integral', (927, 510, 954, 590))]),
            ('sum', (144, 590, 325, 695), [('sum', (477, 600, 504, 695))]),
            ('scripts', (327, 690, 435, 780), [('subscript', (227, 690, 250, 780)), ('superscript', (273, 690, 299, 780))]),
        ]
        for kind, body_roi, anchors in groups:
            body = bounds(ink, body_roi)
            target = reference[kind]
            scale = target['height'] / (body[3] - body[1])
            boxes = []
            for symbol, roi in anchors:
                anchor = bounds(ink, roi)
                boxes.append(anchor)
                reference_delta = target['delta']
                if symbol == 'superscript':
                    reference_delta = center(target['superscript']) - center(target['body'])
                error = (center(anchor) - center(body)) * scale - reference_delta
                native_rows.append({'mode': mode, 'kind': kind, 'symbol': symbol, 'body': body, 'anchor': anchor, 'errorPx': round(error, 3)})
                assert abs(error) <= 3, f'Native {mode}/{symbol}: reference error {error:.3f}px'
            if kind == 'scripts':
                assert abs(boxes[0][3] - boxes[1][3]) <= 1, f'Native {mode}: main L baselines differ'
    report = {'result': 'PASS', 'method': 'Native Word window, black ink (<130 RGB), normalize by CJK painted height', 'tolerancePx': 3, 'rows': native_rows, 'maxReferenceErrorPx': max(abs(row['errorPx']) for row in native_rows)}
    (root / 'native-pixel-acceptance.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    font = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf', 24)
    sheet = Image.new('RGB', (1340, 5 * 455), 'white')
    draw = ImageDraw.Draw(sheet)
    for i, mode in enumerate(['baseline', 'top', 'center', 'bottom', 'auto']):
        draw.text((15, i * 455 + 10), mode, font=font, fill='black')
        shot = Image.open(root / f'word-{mode}-r103-native.png').convert('RGB')
        sheet.paste(shot.crop((140, 385, 1320, 785)), (140, i * 455 + 40))
    sheet.save(root / 'word-five-modes-native-comparison.png')
    print(json.dumps({'result': 'PASS', 'measurements': len(native_rows), 'maxReferenceErrorPx': report['maxReferenceErrorPx']}, ensure_ascii=False))
    sys.exit(0)

rows = []
for page, mode in enumerate(['baseline', 'top', 'center', 'bottom', 'auto'], 1):
    ink = mask(root / f'after-page-{page:02}.png')
    groups = [
        ('display', (710, 140, 887, 265), [('fraction', (380, 170, 407, 245)), ('integral', (1109, 170, 1144, 245))]),
        ('text', (60, 270, 205, 380), [('fraction', (248, 295, 276, 375)), ('integral', (838, 295, 865, 375))]),
        ('sum', (60, 370, 237, 465), [('sum', (397, 390, 425, 470))]),
        ('scripts', (243, 455, 350, 595), [('subscript', (143, 495, 165, 544)), ('superscript', (187, 495, 209, 544))]),
    ]
    for kind, body_roi, anchors in groups:
        body = bounds(ink, body_roi)
        target = reference[kind]
        scale = target['height'] / (body[3] - body[1])
        # Word can change line leading after uncropping without changing the
        # painted formula/body relationship. Locate anchors around the observed
        # body centre; do not clip glyphs to a previous line's absolute y.
        anchor_y = int(center(body))
        span = body[3] - body[1]
        anchors = [(symbol, (roi[0], anchor_y - span, roi[2], anchor_y + span)) for symbol, roi in anchors]
        for symbol, axis_roi in anchors:
            anchor = bounds(ink, axis_roi)
            reference_delta = target['delta']
            if symbol == 'superscript':
                reference_delta = center(target['superscript']) - center(target['body'])
            normalized_delta = (center(anchor) - center(body)) * scale
            error = normalized_delta - reference_delta
            rows.append({'mode': mode, 'kind': kind, 'symbol': symbol, 'body': body, 'anchor': anchor, 'normalizedDeltaPx': round(normalized_delta, 3), 'referenceDeltaPx': reference_delta, 'errorPx': round(error, 3)})
            assert abs(error) <= TOLERANCE, f'{mode}/{kind}/{symbol}: reference error {error:.3f} px'
        if kind == 'scripts':
            left, right = (bounds(ink, roi) for _, roi in anchors)
            assert abs(left[3] - right[3]) <= 1, f'{mode}: primary L baselines differ'
    block = mask(root / f'after-page-{page + 5:02}.png')
    number = bounds(block, (1570, 180, 1740, 400))
    # The tall integral provides a continuous vertical ink component. Select
    # the component beside the observed number, excluding the preceding row.
    ys = np.where(block[100:400, 740:1090].any(axis=1))[0] + 100
    components = np.split(ys, np.where(np.diff(ys) > 3)[0] + 1)
    component = min(components, key=lambda c: abs((c[0] + c[-1] + 1) / 2 - center(number)))
    formula = bounds(block, (740, int(component[0]), 1090, int(component[-1]) + 1))
    assert formula[3] - formula[1] >= 60, 'The complete display integral must be measured'
    error = center(formula) - center(number)
    assert abs(error) <= 1.5, f'{mode}: display/number painted centres differ {error}px'
    rows.append({'mode': mode, 'kind': 'numbered-block', 'formula': formula, 'number': number, 'errorPx': error})

report = {'result': 'PASS', 'method': 'Word 144 dpi black ink (<130 RGB), normalize by CJK painted height', 'tolerancePx': TOLERANCE, 'references': reference, 'rows': rows, 'maxReferenceErrorPx': max(abs(row['errorPx']) for row in rows)}
(root / 'pixel-acceptance.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
# Contact sheet crops retain actual painted pixels; each row is one Word mode.
font = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf', 24)
sheet = Image.new('RGB', (1340, 5 * 455), 'white')
draw = ImageDraw.Draw(sheet)
for page, mode in enumerate(['baseline', 'top', 'center', 'bottom', 'auto'], 1):
    draw.text((15, (page - 1) * 455 + 10), mode, font=font, fill='black')
    image = Image.open(root / f'after-page-{page:02}.png').convert('RGB')
    crop = image.crop((50, 145, 1230, 555))
    sheet.paste(crop, (140, (page - 1) * 455 + 40))
sheet.save(root / 'word-five-modes-comparison.png')
print(json.dumps({'result': 'PASS', 'measurements': len(rows), 'maxReferenceErrorPx': report['maxReferenceErrorPx']}, ensure_ascii=False))
