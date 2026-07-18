"""Create certificate PNGs from students.csv using Adobe Photoshop on Windows.

Install once:
    pip install pandas pywin32

Run:
    python generate.py
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Any, Optional

import pandas as pd

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError:
    print("Error: pywin32 is not installed. Run: pip install pywin32")
    sys.exit(1)


# Keep these files together so the folder can be moved anywhere on the computer.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PSD_FILE = os.path.join(BASE_DIR, "Certificate.psd")
CSV_FILE = os.path.join(BASE_DIR, "students.csv")
OUTPUT = os.path.join(BASE_DIR, "Output")

REQUIRED_COLUMNS = ("NAME", "POSITION", "TENURE")
PS_DONOTSAVECHANGES = 2


def stop(message: str) -> None:
    print(f"Error: {message}")
    sys.exit(1)


def check_files() -> None:
    if not os.path.isfile(PSD_FILE):
        stop(f"PSD template not found: {PSD_FILE}")
    if not os.path.isfile(CSV_FILE):
        stop(f"CSV file not found: {CSV_FILE}")
    os.makedirs(OUTPUT, exist_ok=True)


def load_students() -> pd.DataFrame:
    """Load a UTF-8 CSV and make sure every certificate field is present."""
    try:
        students = pd.read_csv(CSV_FILE, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError, pd.errors.EmptyDataError, pd.errors.ParserError) as exc:
        stop(f"Could not read students.csv: {exc}")

    missing = [column for column in REQUIRED_COLUMNS if column not in students.columns]
    if missing:
        stop("CSV is missing required column(s): " + ", ".join(missing))
    if students.empty:
        stop("students.csv has no certificate rows.")

    students = students.loc[:, list(REQUIRED_COLUMNS)].copy()
    for column in REQUIRED_COLUMNS:
        students[column] = students[column].astype(str).str.strip()

    blank_rows = students.eq("").any(axis=1)
    if blank_rows.any():
        rows = ", ".join(str(row + 2) for row in students.index[blank_rows])
        stop(f"Blank NAME, POSITION, or TENURE values found on CSV row(s): {rows}")
    return students


def connect_to_photoshop() -> Any:
    """Connect to Photoshop without relying on its unavailable type library."""
    try:
        app = win32com.client.Dispatch("Photoshop.Application")
        app.Visible = True
        return app
    except pywintypes.com_error as exc:
        stop(
            "Could not start Adobe Photoshop. Open Photoshop once manually, "
            f"then try again. Details: {exc}"
        )


def call_com_method(com_object: Any, method_name: str, *arguments: Any) -> Any:
    """Call a COM method explicitly.

    Some Photoshop installations do not register a type library. pywin32's
    dynamic interface then mistakes methods such as Open for properties. This
    helper invokes the method directly by DISPID, without makepy/gencache.
    """
    try:
        dispid = com_object._oleobj_.GetIDsOfNames(0, method_name)
        if isinstance(dispid, tuple):
            dispid = dispid[0]
        result = com_object._oleobj_.Invoke(
            dispid, 0, pythoncom.DISPATCH_METHOD, True, *arguments
        )
        return win32com.client.Dispatch(result) if result is not None else None
    except pywintypes.com_error as exc:
        raise RuntimeError(f"Photoshop could not run {method_name}(): {exc}") from exc


def find_layer_by_name(container: Any, layer_name: str) -> Optional[Any]:
    """Recursively find a named layer, including layers inside folders."""
    try:
        layers = container.Layers
        layer_count = layers.Count
    except (AttributeError, pywintypes.com_error):
        return None

    for index in range(1, layer_count + 1):
        layer = call_com_method(layers, "Item", index)
        if str(layer.Name) == layer_name:
            return layer
        nested_layer = find_layer_by_name(layer, layer_name)
        if nested_layer is not None:
            return nested_layer
    return None


def replace_text(document: Any, layer_name: str, value: str) -> None:
    layer = find_layer_by_name(document, layer_name)
    if layer is None:
        raise RuntimeError(f"Required layer '{layer_name}' was not found in Certificate.psd.")
    try:
        layer.TextItem.Contents = str(value)
    except pywintypes.com_error as exc:
        raise RuntimeError(f"Layer '{layer_name}' must be an editable text layer. Details: {exc}") from exc


def safe_filename(name: str) -> str:
    """Keep normal names intact while removing Windows-invalid filename characters."""
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(".")
    return filename or "unnamed_certificate"


def export_png(document: Any, output_file: Path) -> None:
    """Save a PNG copy. The open PSD remains unchanged and is never overwritten."""
    options = win32com.client.Dispatch("Photoshop.PNGSaveOptions")
    options.Interlaced = False
    call_com_method(document, "SaveAs", str(output_file), options, True)


def generate_certificates(app: Any, students: pd.DataFrame) -> None:
    total = len(students)
    created_filenames: set[str] = set()

    for number, (_, student) in enumerate(students.iterrows(), start=1):
        name = str(student["NAME"])
        filename = safe_filename(name)
        if filename.casefold() in created_filenames:
            raise RuntimeError(f"Duplicate output filename for '{name}'. Rename one student in the CSV.")
        created_filenames.add(filename.casefold())

        print(f"Generating {number}/{total}")
        document: Optional[Any] = None
        step = "opening Certificate.psd"
        try:
            document = call_com_method(app, "Open", PSD_FILE)
            step = "updating NAME"
            replace_text(document, "NAME", name)
            step = "updating POSITION"
            replace_text(document, "POSITION", str(student["POSITION"]))
            step = "updating TENURE"
            replace_text(document, "TENURE", str(student["TENURE"]))
            step = "saving PNG"
            export_png(document, Path(OUTPUT) / f"{filename}.png")
        except (pywintypes.com_error, RuntimeError, OSError) as exc:
            raise RuntimeError(f"Could not generate '{name}' while {step}: {exc}") from exc
        finally:
            if document is not None:
                try:
                    call_com_method(document, "Close", PS_DONOTSAVECHANGES)
                except RuntimeError:
                    print(f"Warning: Photoshop could not close '{name}' cleanly without saving.")


def main() -> None:
    check_files()
    students = load_students()
    photoshop = connect_to_photoshop()
    try:
        generate_certificates(photoshop, students)
    except RuntimeError as exc:
        stop(str(exc))
    print("Done!")


if __name__ == "__main__":
    main()
