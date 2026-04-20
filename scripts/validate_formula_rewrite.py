import json
from datetime import UTC, datetime
from pathlib import Path

import xlwings as xw


ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"
FIXTURES = ROOT / "fixtures"


def cell_snapshot(sheet, address):
    rng = sheet.range(address)
    return {"value": rng.value, "formula": rng.formula}


def matrix_snapshot(sheet, start_row, end_row, start_col, end_col):
    cells = {}
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            address = xw.utils.col_name(col) + str(row)
            cells[address] = cell_snapshot(sheet, address)
    return cells


def case8_snapshot(book):
    summary = book.sheets["Summary"]
    return {
        "sheets": [sheet.name for sheet in book.sheets],
        "summary": {
            "B2": cell_snapshot(summary, "B2"),
            "B3": cell_snapshot(summary, "B3"),
            "D2:F4": matrix_snapshot(summary, 2, 4, 4, 6),
        },
    }


def case9_snapshot(book):
    summary = book.sheets["Summary"]
    data = book.sheets["Data"]
    return {
        "sheets": [sheet.name for sheet in book.sheets],
        "data": {
            "A1:D6": matrix_snapshot(data, 1, 6, 1, 4),
        },
        "summary": {
            "B2": cell_snapshot(summary, "B2"),
            "B3": cell_snapshot(summary, "B3"),
            "E2:J6": matrix_snapshot(summary, 2, 6, 5, 10),
        },
    }


def open_and_capture(app, path, case_name):
    book = app.books.open(str(path))
    try:
        app.calculate()
        if case_name == "case8":
            return {"opened": True, **case8_snapshot(book)}
        if case_name == "case9":
            return {"opened": True, **case9_snapshot(book)}
        raise ValueError(f"Unknown case name: {case_name}")
    finally:
        book.close()


def main():
    validations = []
    targets = [
        ("case8", "fixture", FIXTURES / "case8_rename_fixture.xlsx"),
        ("case8", "codex", OUTPUTS / "case8_codex.xlsx"),
        ("case8", "witan", OUTPUTS / "case8_witan.xlsx"),
        ("case9", "fixture", FIXTURES / "case9_shift_fixture.xlsx"),
        ("case9", "codex", OUTPUTS / "case9_codex.xlsx"),
        ("case9", "witan", OUTPUTS / "case9_witan.xlsx"),
    ]

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for case_name, variant, path in targets:
            entry = {
                "case": case_name,
                "variant": variant,
                "path": str(path),
            }
            if not path.exists():
                entry["status"] = "missing"
                validations.append(entry)
                continue
            try:
                entry["status"] = "ok"
                entry["excel"] = open_and_capture(app, path, case_name)
            except Exception as exc:
                entry["status"] = "error"
                entry["error"] = str(exc)
            validations.append(entry)
    finally:
        app.quit()

    output_path = OUTPUTS / "formula_rewrite_excel_validation.json"
    output_path.write_text(
        json.dumps(
            {
                "validatedAt": datetime.now(UTC).isoformat(),
                "results": validations,
            },
            indent=2,
        )
        + "\n"
    )
    print(json.dumps({"outputPath": str(output_path), "results": validations}, indent=2))


if __name__ == "__main__":
    main()
