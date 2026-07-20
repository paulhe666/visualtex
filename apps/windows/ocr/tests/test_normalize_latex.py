import unittest

from normalize_latex import normalize_formula_latex


class NormalizeFormulaLatexTests(unittest.TestCase):
    def test_strips_display_math_wrapper(self) -> None:
        self.assertEqual(
            normalize_formula_latex(r"$$  \frac{a}{b}  $$"),
            r"\frac{a}{b}",
        )

    def test_strips_bracket_wrapper(self) -> None:
        self.assertEqual(normalize_formula_latex(r"\[x^2+y^2\]"), r"x^2+y^2")

    def test_preserves_matrix_row_separator(self) -> None:
        value = "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}"
        self.assertEqual(normalize_formula_latex(value), value)

    def test_flattens_physical_newlines(self) -> None:
        self.assertEqual(normalize_formula_latex("a +\nb"), "a + b")


if __name__ == "__main__":
    unittest.main()
