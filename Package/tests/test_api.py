"""Tests of the package's public names and of the examples in README.md."""

import contextlib
import importlib.metadata
import io
import re
import unittest

import eiopa_rfr
from support import DATA_DIR, README_FILE, run_python

# the modules of Code, which the package must only ever import as eiopa_rfr.<module>
CODE_MODULES: tuple[str, ...] = ("calculation", "data_io", "interest_rate_risk", "main", "validation",
                                 "volatility_adjustment")


class TestPublicNames(unittest.TestCase):

    def test_every_public_name_exists(self) -> None:
        for name in eiopa_rfr.__all__:
            self.assertTrue(hasattr(eiopa_rfr, name), name)

    def test_public_names_are_those_of_the_modules(self) -> None:
        self.assertIs(eiopa_rfr.calculate_sheet, eiopa_rfr.calculation.calculate_sheet)
        self.assertIs(eiopa_rfr.calculate_va_sheet, eiopa_rfr.volatility_adjustment.calculate_va_sheet)
        self.assertIs(eiopa_rfr.calculate_irr_up_sheet, eiopa_rfr.interest_rate_risk.calculate_irr_up_sheet)

    def test_modules_share_one_copy_of_each_other(self) -> None:
        # relative imports: every module uses the objects of eiopa_rfr.calculation, not a second copy
        self.assertIs(eiopa_rfr.main.calculate_sheet, eiopa_rfr.calculation.calculate_sheet)
        self.assertIs(eiopa_rfr.data_io.INSTRUMENTS, eiopa_rfr.calculation.INSTRUMENTS)

    def test_no_module_is_imported_under_its_name_in_code(self) -> None:
        process = run_python("-c", "import sys, eiopa_rfr, eiopa_rfr.cli; "
                                   f"print(sorted(set(sys.modules) & {set(CODE_MODULES)!r}))", cwd=DATA_DIR)
        self.assertEqual((process.returncode, process.stdout.strip()), (0, "[]"), process.stderr)

    def test_version_is_that_of_the_installed_package(self) -> None:
        try:
            installed: str = importlib.metadata.version("eiopa-rfr")
        except importlib.metadata.PackageNotFoundError:
            self.skipTest("the package is not installed")
        self.assertEqual(eiopa_rfr.__version__, installed)


@unittest.skipUnless(README_FILE.exists(), "needs README.md of the folder Package")
class TestReadme(unittest.TestCase):

    def test_examples_print_what_their_comments_say(self) -> None:
        # every print(...) in the examples is followed by a comment with what it prints; an explanation
        # in brackets after two spaces is not part of it
        blocks: list[str] = re.findall(r"```python\n(.*?)```", README_FILE.read_text(encoding="utf-8"), re.DOTALL)
        expected: list[str] = [comment.split("  (")[0].strip() for block in blocks
                               for comment in re.findall(r"^print\(.*\)\s+#(.*)$", block, re.MULTILINE)]
        namespace: dict[str, object] = {}
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            for block in blocks:
                exec(block, namespace)
        self.assertGreater(len(expected), 0)
        self.assertEqual(buffer.getvalue().splitlines(), expected)


if __name__ == "__main__":
    unittest.main()
