"""Tests of sync_code.py: the package must be generated from the current Code folder.

These tests need the whole repository; they are skipped where only the package is available, for
example when testing from a source distribution.
"""

import contextlib
import importlib.util
import io
import unittest
from pathlib import Path
from types import ModuleType

SYNC_SCRIPT: Path = Path(__file__).resolve().parents[1] / "sync_code.py"
CODE_DIR: Path = SYNC_SCRIPT.parents[1] / "Code"


def load_sync_code() -> ModuleType:
    """Import sync_code.py from the folder Package."""
    spec = importlib.util.spec_from_file_location("sync_code", SYNC_SCRIPT)
    assert spec is not None and spec.loader is not None
    module: ModuleType = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@unittest.skipUnless(SYNC_SCRIPT.exists() and CODE_DIR.is_dir(), "needs the repository's Code folder")
class TestSync(unittest.TestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.sync: ModuleType = load_sync_code()

    def test_package_is_up_to_date_with_code(self) -> None:
        self.assertEqual(self.sync.differences(self.sync.expected_files()), [],
                         "run python Package/sync_code.py")

    def test_check_exit_code(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(self.sync.main(["--check"]), 0)

    def test_imports_become_relative(self) -> None:
        source: str = ('"""Docstring that mentions the module calculation.\n'
                       '\n'
                       'from the year-14 forward\n'
                       '"""\n'
                       'import math\n'
                       'from calculation import (Curve,\n'
                       '                         Flags)\n'
                       'from data_io import read_csv\n')
        self.assertEqual(self.sync.rewrite_imports(source, ["calculation", "data_io"]),
                         source.replace("from calculation import", "from .calculation import")
                               .replace("from data_io import", "from .data_io import"))

    def test_imports_that_cannot_be_made_relative(self) -> None:
        for source in ("import calculation\n", "def f():\n    from calculation import Curve\n"):
            with self.subTest(source=source):
                with self.assertRaisesRegex(ValueError, "cannot make this import relative"):
                    self.sync.rewrite_imports(source, ["calculation"])


if __name__ == "__main__":
    unittest.main()
