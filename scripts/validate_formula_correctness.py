import json
from datetime import UTC, date, datetime
from pathlib import Path

import xlwings as xw


ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"
FIXTURES = ROOT / "fixtures"


def cell_snapshot(sheet, address):
    rng = sheet.range(address)
    return {"value": normalize_value(rng.value), "formula": normalize_value(rng.formula)}


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


def case10_snapshot(book):
    summary = book.sheets["Summary"]
    rows = []
    for row in range(2, 19):
        label = summary.range(f"A{row}").value
        expected = summary.range(f"B{row}").value
        actual = cell_snapshot(summary, f"C{row}")
        rows.append(
            {
                "row": row,
                "label": normalize_value(label),
                "expected": normalize_value(expected),
                "actual": actual["value"],
                "formula": actual["formula"],
                "matchesExpected": actual["value"] == normalize_value(expected),
            }
        )
    return {
        "sheets": [sheet.name for sheet in book.sheets],
        "rows": rows,
    }


def open_and_capture(app, path):
    book = app.books.open(str(path))
    try:
        app.calculate()
        return {"opened": True, **case10_snapshot(book)}
    finally:
        book.close()


def main():
    validations = []
    targets = [
        ("fixture", FIXTURES / "case10_formula_fixture.xlsx"),
        ("codex", OUTPUTS / "case10_codex.xlsx"),
        ("witan", OUTPUTS / "case10_witan.xlsx"),
    ]

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for variant, path in targets:
            entry = {
                "case": "case10",
                "variant": variant,
                "path": str(path),
            }
            if not path.exists():
                entry["status"] = "missing"
                validations.append(entry)
                continue
            try:
                entry["status"] = "ok"
                entry["excel"] = open_and_capture(app, path)
            except Exception as exc:
                entry["status"] = "error"
                entry["error"] = str(exc)
            validations.append(entry)
    finally:
        app.quit()

    output_path = OUTPUTS / "formula_correctness_excel_validation.json"
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
