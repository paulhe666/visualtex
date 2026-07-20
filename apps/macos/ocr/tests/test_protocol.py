import unittest
from pathlib import Path

from protocol import ProtocolError, parse_request


class ProtocolTests(unittest.TestCase):
    def test_ping(self) -> None:
        request = parse_request({"id": "1", "action": "ping"})
        self.assertEqual(request.action, "ping")
        self.assertEqual(request.model, "PP-FormulaNet_plus-M")

    def test_recognize_requires_absolute_path(self) -> None:
        with self.assertRaises(ProtocolError) as context:
            parse_request(
                {
                    "id": "1",
                    "action": "recognize",
                    "image_path": "relative.png",
                }
            )
        self.assertEqual(context.exception.code, "INVALID_REQUEST")

    def test_recognize_parses_path(self) -> None:
        request = parse_request(
            {
                "id": "1",
                "action": "recognize",
                "image_path": "/tmp/formula.png",
                "model": "PP-FormulaNet-S",
                "device": "cpu",
            }
        )
        self.assertEqual(request.image_path, Path("/tmp/formula.png"))
        self.assertEqual(request.model, "PP-FormulaNet-S")

    def test_rejects_unknown_model(self) -> None:
        with self.assertRaises(ProtocolError):
            parse_request({"id": "1", "action": "warmup", "model": "unknown"})


if __name__ == "__main__":
    unittest.main()
