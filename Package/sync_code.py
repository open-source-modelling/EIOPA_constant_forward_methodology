"""Generate the package eiopa_rfr from the standalone code in the folder Code.

Code is the only copy anyone edits: it stays a set of plain scripts that can be read and run with
python Code/main.py. This script copies every module of Code into src/eiopa_rfr and changes only
the imports between them, which become relative imports ("from calculation import" becomes
"from .calculation import"), and adds a first line saying where the file comes from. It also copies
the files the package needs from the rest of the repository: the license, and the EIOPA example and
the workbook's shocked curves that the package's tests run on (into tests/data).

Usage (from any folder):
    python Package/sync_code.py           write the copies that are missing or out of date, and
                                          delete generated files whose source no longer exists
    python Package/sync_code.py --check   change nothing; list the copies that are missing, out
                                          of date or no longer needed, and exit with code 1 if
                                          there are any (for CI)

Line endings are ignored when comparing, so the check gives the same answer on Windows, where Git
may check the files out with CRLF, as elsewhere. A copy is written with the line endings of its source.
"""

import argparse
import re
import sys
from pathlib import Path

PACKAGE_DIR: Path = Path(__file__).resolve().parent
REPO_DIR: Path = PACKAGE_DIR.parent
CODE_DIR: Path = REPO_DIR / "Code"
MODULE_DIR: Path = PACKAGE_DIR / "src" / "eiopa_rfr"
TEST_DATA_DIR: Path = PACKAGE_DIR / "tests" / "data"
EXAMPLE_INPUT_DIR: Path = REPO_DIR / "Examples" / "EIOPA_Example" / "Input"

# first line of every generated module; it also tells generated modules from the package's own
HEADER: str = "# Generated from Code/{name} by Package/sync_code.py: edit that file and run the script again."
HEADER_START: str = "# Generated from Code/"


def rewrite_imports(source: str, modules: list[str]) -> str:
    """Turn the imports of the modules of Code into relative imports.

    Only imports at the start of a line are changed ("from calculation import ..." becomes
    "from .calculation import ..."), also when the names continue on the next lines in brackets.

    Args:
        source: the text of a module of Code.
        modules: the names of the modules of Code, without ".py".

    Returns:
        The text with relative imports.

    Raises:
        ValueError: if an import of a module of Code is left that this function cannot rewrite,
            such as "import calculation" or an import inside a function.
    """
    names: str = "|".join(re.escape(module) for module in modules)
    text: str = re.sub(rf"^from ({names}) import ", r"from .\1 import ", source, flags=re.MULTILINE)
    left: re.Match[str] | None = re.search(rf"^[ \t]*(import ({names})\b|from ({names}) import )", text, re.MULTILINE)
    if left:
        raise ValueError(f"cannot make this import relative: {left.group(0).strip()!r}")
    return text


def read_text(path: Path) -> str:
    """The text of a UTF-8 file with its line endings as they are."""
    return path.read_bytes().decode("utf-8")


def normalised(text: str) -> str:
    """The text with every line ending as \\n, for comparing."""
    return text.replace("\r\n", "\n")


def expected_files() -> dict[Path, str]:
    """Every file this script generates, with the text it should have.

    Returns:
        {path in Package: text}: the modules of Code with relative imports and the header line, the
        license, and the EIOPA example inputs and the workbook's shocked curves for the tests.
    """
    code_modules: list[Path] = sorted(CODE_DIR.glob("*.py"))
    modules: list[str] = [path.stem for path in code_modules]
    files: dict[Path, str] = {}
    for path in code_modules:
        source: str = read_text(path)
        newline: str = "\r\n" if "\r\n" in source else "\n"
        files[MODULE_DIR / path.name] = HEADER.format(name=path.name) + newline + rewrite_imports(source, modules)
    copies: dict[Path, Path] = {REPO_DIR / "LICENSE": PACKAGE_DIR / "LICENSE",
                                CODE_DIR / "tests" / "irr_workbook_values.csv": TEST_DATA_DIR / "irr_workbook_values.csv"}
    copies.update({path: TEST_DATA_DIR / "eiopa_example" / path.name for path in sorted(EXAMPLE_INPUT_DIR.glob("*.csv"))})
    files.update({target: read_text(source) for source, target in copies.items()})
    return files


def stale_files(expected: dict[Path, str]) -> list[Path]:
    """Generated files whose source no longer exists.

    These are the modules in src/eiopa_rfr that start with the header line, and any file in
    tests/data, that expected_files does not generate.

    Args:
        expected: the result of expected_files.

    Returns:
        The paths of the files, sorted.
    """
    generated: list[Path] = [path for path in MODULE_DIR.glob("*.py") if read_text(path).startswith(HEADER_START)]
    if TEST_DATA_DIR.exists():
        generated += [path for path in TEST_DATA_DIR.rglob("*") if path.is_file()]
    return sorted(path for path in generated if path not in expected)


def differences(expected: dict[Path, str]) -> list[str]:
    """Describe every generated file that is missing, out of date or no longer needed.

    Args:
        expected: the result of expected_files.

    Returns:
        One line per file, such as "out of date: Package/src/eiopa_rfr/calculation.py"; empty if
        the package is up to date.
    """
    lines: list[str] = []
    for path, text in expected.items():
        if not path.exists():
            lines.append(f"missing: {relative(path)}")
        elif normalised(read_text(path)) != normalised(text):
            lines.append(f"out of date: {relative(path)}")
    lines += [f"no longer generated: {relative(path)}" for path in stale_files(expected)]
    return lines


def relative(path: Path) -> str:
    """The path relative to the repository folder, with forward slashes."""
    return path.relative_to(REPO_DIR).as_posix()


def sync(expected: dict[Path, str]) -> list[str]:
    """Write the missing and out-of-date files and delete those no longer generated.

    Files that are up to date are not touched, also when only their line endings differ.

    Args:
        expected: the result of expected_files.

    Returns:
        One line per file written or deleted, such as "updated: Package/src/eiopa_rfr/calculation.py".
    """
    lines: list[str] = []
    for path, text in expected.items():
        if path.exists() and normalised(read_text(path)) == normalised(text):
            continue
        lines.append(f"{'updated' if path.exists() else 'created'}: {relative(path)}")
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode("utf-8"))
    for path in stale_files(expected):
        path.unlink()
        lines.append(f"deleted: {relative(path)}")
    return lines


def main(argv: list[str] | None = None) -> int:
    """Run the script; see the module docstring.

    Args:
        argv: the command line arguments; sys.argv[1:] by default.

    Returns:
        The exit code: 1 if --check found differences, otherwise 0.
    """
    parser = argparse.ArgumentParser(description="Generate the package eiopa_rfr from the folder Code.")
    parser.add_argument("--check", action="store_true",
                        help="change nothing; exit with code 1 if a generated file is missing, out of date or stale")
    args = parser.parse_args(argv)
    expected: dict[Path, str] = expected_files()
    if args.check:
        lines: list[str] = differences(expected)
        print("\n".join(lines) if lines else "The package is up to date with Code.")
        if lines:
            print("Run python Package/sync_code.py to update it.")
        return 1 if lines else 0
    lines = sync(expected)
    print("\n".join(lines) if lines else "The package was already up to date with Code.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
