import json
from datetime import UTC, date, datetime
from pathlib import Path

import xlwings as xw


ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"


def normalize_value(value):
    if isinstance(value, datetime):
        if (
            value.hour == 0
            and value.minute == 0
            and value.second == 0
            and value.microsecond == 0
        ):
            return value.date().isoformat()
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def cell_snapshot(sheet, address):
    rng = sheet.range(address)
    return {"value": normalize_value(rng.value), "formula": normalize_value(rng.formula)}


def matrix_snapshot(sheet, start_row, end_row, start_col, end_col):
    cells = {}
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            address = xw.utils.col_name(col) + str(row)
            cells[address] = cell_snapshot(sheet, address)
    return cells


def case2_snapshot(book):
    inputs = book.sheets["Inputs"]
    model = book.sheets["Model"]
    return {
        "sheets": [sheet.name for sheet in book.sheets],
        "inputs": {"B4": cell_snapshot(inputs, "B4")},
        "model": {
            "B3": cell_snapshot(model, "B3"),
            "B4": cell_snapshot(model, "B4"),
            "B6": cell_snapshot(model, "B6"),
            "B7": cell_snapshot(model, "B7"),
        },
    }


def case4_snapshot(book):
    data = book.sheets["Data"]
    return {
        "sheets": [sheet.name for sheet in book.sheets],
        "data": {
            "A1:C3": matrix_snapshot(data, 1, 3, 1, 3),
        },
    }


def case5_snapshot(book):
    summary = book.sheets["Summary"]
    return {
        "sheets": [sheet.name for sheet in book.sheets],
        "summary": {
            "D2:D6": matrix_snapshot(summary, 2, 6, 4, 4),
        },
    }


def open_and_capture(app, path, case_name):
    book = app.books.open(str(path))
    try:
        app.calculate()
        if case_name == "case2":
            return {"opened": True, **case2_snapshot(book)}
        if case_name == "case4":
            return {"opened": True, **case4_snapshot(book)}
        if case_name == "case5":
            return {"opened": True, **case5_snapshot(book)}
        raise ValueError(f"Unknown case name: {case_name}")
    finally:
        book.close()


def main():
    validations = []
    targets = [
        ("case2", OUTPUTS / "case2_witan.xlsx"),
        ("case4", OUTPUTS / "case4_witan.xlsx"),
        ("case5", OUTPUTS / "case5_witan.xlsx"),
    ]

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for case_name, path in targets:
            entry = {"case": case_name, "path": str(path)}
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

    output_path = OUTPUTS / "reused_witan_excel_validation.json"
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
