from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from PIL import Image, ImageDraw


WORKER_PATH = Path(__file__).resolve().parents[2] / "src-tauri" / "ocr" / "worker.py"
SPEC = importlib.util.spec_from_file_location("visualtex_ocr_worker", WORKER_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("Unable to load VisualTeX OCR worker")
WORKER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(WORKER)


def make_formula_like_image(background: int, foreground: int) -> Image.Image:
    image = Image.new("L", (360, 100), background)
    draw = ImageDraw.Draw(image)
    draw.line((40, 54, 320, 54), fill=foreground, width=4)
    draw.ellipse((80, 18, 118, 46), outline=foreground, width=4)
    draw.line((160, 22, 190, 80), fill=foreground, width=4)
    draw.line((190, 22, 160, 80), fill=foreground, width=4)
    draw.rectangle((240, 20, 292, 82), outline=foreground, width=4)
    return image.convert("RGB")


class WorkerPreprocessTests(unittest.TestCase):
    def preprocess(self, image: Image.Image):
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "source.png"
            target = Path(temporary_directory) / "processed.png"
            image.save(source)
            _, size, metadata = WORKER._preprocess_image(str(source), str(target))
            processed = Image.open(target).convert("L").copy()
            return processed, size, metadata

    def test_white_background_is_not_inverted(self):
        processed, _, metadata = self.preprocess(make_formula_like_image(255, 0))
        self.assertFalse(metadata["background_inverted"])
        self.assertGreater(processed.getpixel((0, 0)), 245)
        self.assertLess(min(processed.getdata()), 20)

    def test_black_background_is_inverted_to_white(self):
        processed, _, metadata = self.preprocess(make_formula_like_image(0, 255))
        self.assertTrue(metadata["background_inverted"])
        self.assertGreater(processed.getpixel((0, 0)), 245)
        self.assertLess(min(processed.getdata()), 20)

    def test_transparent_white_formula_survives_compositing(self):
        image = Image.new("RGBA", (240, 80), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.line((20, 40, 220, 40), fill=(255, 255, 255, 255), width=4)
        draw.ellipse((80, 10, 120, 70), outline=(255, 255, 255, 255), width=4)
        processed, _, metadata = self.preprocess(image)
        self.assertTrue(metadata["background_inverted"])
        self.assertGreater(processed.getpixel((0, 0)), 245)
        self.assertLess(min(processed.getdata()), 20)

    def test_warmup_loads_the_requested_model_without_recognition(self):
        with mock.patch.object(WORKER, "_load_model", return_value=object()) as load_model:
            response = WORKER._handle(
                {
                    "id": "warmup-test",
                    "action": "warmup",
                    "model": "PP-FormulaNet_plus-M",
                    "device": "cpu",
                }
            )

        load_model.assert_called_once_with("PP-FormulaNet_plus-M", "cpu")
        self.assertTrue(response["ok"])
        self.assertEqual(response["event"], "warmed")
        self.assertEqual(response["model"], "PP-FormulaNet_plus-M")


if __name__ == "__main__":
    unittest.main()
