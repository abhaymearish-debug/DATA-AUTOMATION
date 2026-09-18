"""
Excel Formula Recalculation Script
Recalculates all formulas in an Excel file using LibreOffice
"""

import json
import os
import platform
import subprocess
import sys
from pathlib import Path

from office.soffice import get_soffice_env

from openpyxl import load_workbook

MACRO_DIR_MACOS = "~/Library/Application Support/LibreOffice/4/user/basic/Standard"
MACRO_DIR_LINUX = "~/.config/libreoffice/4/user/basic/Standard"
MACRO_FILENAME = "Module1.xba"

RECALCULATE_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def has_gtimeout():
    try:
        subprocess.run(
            ["gtimeout", "--version"], capture_output=True, timeout=1, check=False
        )
        return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def setup_libreoffice_macro():
    macro_dir = os.path.expanduser(
        MACRO_DIR_MACOS if platform.system() == "Darwin" else MACRO_DIR_LINUX
    )
    macro_file = os.path.join(macro_dir, MACRO_FILENAME)

    if (
        os.path.exists(macro_file)
        and "RecalculateAndSave" in Path(macro_file).read_text()
    ):
        return True

    if not os.path.exists(macro_dir):
        subprocess.run(
            ["soffice", "--headless", "--terminate_after_init"],
            capture_output=True,
            timeout=10,
            env=get_soffice_env(),
        )
        os.makedirs(macro_dir, exist_ok=True)

    try:
        Path(macro_file).write_text(RECALCULATE_MACRO)
        return True
    except Exception:
        return False


def recalc(filename, timeout=30):
    if not Path(filename).exists():
        return {"error": f"File {filename} does not exist"}

    abs_path = str(Path(filename).absolute())

    if not setup_libreoffice_macro():
        return {"error": "Failed to setup LibreOffice macro"}

    cmd = [
        "soffice",
        "--headless",
        "--norestore",
        "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
        abs_path,
    ]

    if platform.system() == "Linux":
        cmd = ["timeout", str(timeout)] + cmd
    elif platform.system() == "Darwin" and has_gtimeout():
        cmd = ["gtimeout", str(timeout)] + cmd

    # Record the file's mtime before the run so we can tell whether the macro's
    # store() actually rewrote it (used for the timeout check below).
    try:
        pre_mtime = Path(abs_path).stat().st_mtime
    except OSError:
        pre_mtime = 0.0

    result = subprocess.run(cmd, capture_output=True, text=True, env=get_soffice_env())

    if result.returncode != 0 and result.returncode != 124:
        error_msg = result.stderr or "Unknown error during recalculation"
        if "Module1" in error_msg or "RecalculateAndSave" not in error_msg:
            return {"error": "LibreOffice macro not configured properly"}
        return {"error": error_msg}

    # Fix 1 Jun 2026: exit code 124 means timeout/gtimeout KILLED soffice. The
    # macro runs calculateAll(); store(); close(True) — if it was killed before
    # store() finished, the workbook was NOT recalculated/saved, yet the
    # data_only scan below would read STALE cached values and falsely report
    # "success". Only trust a 124 run if the file was actually rewritten.
    if result.returncode == 124:
        try:
            saved = Path(abs_path).stat().st_mtime > pre_mtime
        except OSError:
            saved = False
        if not saved:
            return {"error": (
                f"recalc timed out after {timeout}s before saving — formulas "
                f"were NOT refreshed (LibreOffice killed at the timeout). "
                f"Increase the timeout or close other soffice instances, then "
                f"rerun. The workbook was left untouched.")}

    try:
        wb = load_workbook(filename, data_only=True)
        wb_formulas = load_workbook(filename, data_only=False)

        excel_errors = [
            "#VALUE!",
            "#DIV/0!",
            "#REF!",
            "#NAME?",
            "#NULL!",
            "#NUM!",
            "#N/A",
        ]
        error_details = {err: [] for err in excel_errors}
        total_errors = 0
        formula_count = 0
        unresolved = []   # formula cells with no cached value after recalc

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            wsf = wb_formulas[sheet_name]
            # (a) error-string scan over the resolved (data_only) values
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        for err in excel_errors:
                            if err in cell.value:
                                location = f"{sheet_name}!{cell.coordinate}"
                                error_details[err].append(location)
                                total_errors += 1
                                break
            # (b) formula coverage + UNRESOLVED detection. Fix 1 Jun 2026: a
            # formula cell that LibreOffice did NOT evaluate reads back as None
            # under data_only, so a genuine error/uncomputed cell was previously
            # counted as clean. After a real store() every formula has a cached
            # value (even "" or 0), so a None here means the formula was never
            # recalculated — surface it instead of silently passing.
            for row in wsf.iter_rows():
                for cell in row:
                    fv = cell.value
                    if isinstance(fv, str) and fv.startswith("="):
                        formula_count += 1
                        if ws[cell.coordinate].value is None:
                            unresolved.append(f"{sheet_name}!{cell.coordinate}")

        wb.close()
        wb_formulas.close()

        problems = total_errors + len(unresolved)
        result = {
            "status": "success" if problems == 0 else "errors_found",
            "total_errors": total_errors,
            "unresolved_formulas": len(unresolved),
            "total_formulas": formula_count,
            "error_summary": {},
        }

        for err_type, locations in error_details.items():
            if locations:
                result["error_summary"][err_type] = {
                    "count": len(locations),
                    "locations": locations[:20],
                }
        if unresolved:
            result["unresolved_formula_locations"] = unresolved[:20]

        return result

    except Exception as e:
        return {"error": str(e)}


def main():
    if len(sys.argv) < 2:
        print("Usage: python recalc.py <excel_file> [timeout_seconds]")
        print("\nRecalculates all formulas in an Excel file using LibreOffice")
        print("\nReturns JSON with error details:")
        print("  - status: 'success' or 'errors_found'")
        print("  - total_errors: Total number of Excel errors found")
        print("  - total_formulas: Number of formulas in the file")
        print("  - error_summary: Breakdown by error type with locations")
        print("    - #VALUE!, #DIV/0!, #REF!, #NAME?, #NULL!, #NUM!, #N/A")
        sys.exit(1)

    filename = sys.argv[1]
    timeout = int(sys.argv[2]) if len(sys.argv) > 2 else 30

    result = recalc(filename, timeout)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
