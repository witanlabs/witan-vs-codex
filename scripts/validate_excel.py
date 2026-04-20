import json
import shutil
import sys
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import xlwings as xw


THREAD_NS = {"xltc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"}


def read_json(path: Path):
    return json.loads(path.read_text())


def parse_threaded_comments(path: Path):
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        thread_name = next((name for name in names if "threadedcomment" in name.lower()), None)
        person_name = next((name for name in names if "person" in name.lower()), None)
        if not thread_name:
            return {"parts": names, "threads": [], "people": []}

        thread_root = ET.fromstring(zf.read(thread_name))
        threads = []
        for node in thread_root.findall("xltc:threadedComment", THREAD_NS):
            threads.append(
                {
                    "ref": node.attrib.get("ref"),
                    "done": node.attrib.get("done"),
                    "personId": node.attrib.get("personId"),
                    "id": node.attrib.get("id"),
                    "text": node.findtext("xltc:text", default="", namespaces=THREAD_NS),
                }
            )

        people = []
        if person_name:
            person_root = ET.fromstring(zf.read(person_name))
            for person in person_root.findall("xltc:person", THREAD_NS):
                people.append(
                    {
                        "id": person.attrib.get("id"),
                        "displayName": person.attrib.get("displayName"),
                    }
                )

        return {"parts": names, "threads": threads, "people": people}


def roundtrip_excel(app, source_path: Path, suffix: str):
    target = source_path.with_name(f"{source_path.stem}_{suffix}{source_path.suffix}")
    shutil.copy2(source_path, target)
    roundtrip = app.books.open(str(target))
    try:
        roundtrip.save()
    finally:
        roundtrip.close()
    return target


def to_float(value):
    return None if value is None else float(value)


def case2(app, path: Path):
    wb = app.books.open(str(path))
    try:
        app.calculate()
        inputs = wb.sheets["Inputs"]
        model = wb.sheets["Model"]
        return {
            "opened": True,
            "bonusRate": to_float(inputs.range("B4").value),
            "profit": to_float(model.range("B3").value),
            "bonus": to_float(model.range("B4").value),
            "netIncome": to_float(model.range("B7").value),
        }
    finally:
        wb.close()


def case4(app, path: Path):
    wb = app.books.open(str(path))
    try:
        roundtrip_path = roundtrip_excel(app, path, "excel_roundtrip")
    finally:
        wb.close()

    return {
        "opened": True,
        "roundtripPath": str(roundtrip_path),
        "original": parse_threaded_comments(path),
        "roundtrip": parse_threaded_comments(roundtrip_path),
    }


def case5(app, path: Path):
    wb = app.books.open(str(path))
    try:
        app.calculate()
        sheet = wb.sheets["Summary"]
        cells = {}
        for row in range(2, 7):
            address = f"D{row}"
            cells[address] = {
                "value": sheet.range(address).value,
                "formula": sheet.range(address).formula,
            }
        roundtrip_path = roundtrip_excel(app, path, "excel_roundtrip")
        return {"opened": True, "cells": cells, "roundtripPath": str(roundtrip_path)}
    finally:
        wb.close()


def case7(app, path: Path):
    wb = app.books.open(str(path))
    try:
        app.calculate()
        sheet = wb.sheets["Summary"]
        cells = {}
        for address in ["D2", "D3", "D4", "D5", "F2", "G2", "H2"]:
            cells[address] = {
                "value": sheet.range(address).value,
                "formula": sheet.range(address).formula,
            }
        return {"opened": True, "cells": cells}
    finally:
        wb.close()


VALIDATORS = {
    "case2": case2,
    "case4": case4,
    "case5": case5,
    "case7": case7,
}


def main():
    artifact_results = read_json(Path(sys.argv[1]))
    validation = {
        "validatedAt": None,
        "artifactResults": sys.argv[1],
        "results": [],
    }

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for result in artifact_results["results"]:
            case_name = result["name"]
            entry = {
                "name": case_name,
                "artifactStatus": result["status"],
                "outputPath": result.get("outputPath"),
            }
            if result["status"] != "ok" or not result.get("outputPath"):
                entry["excelStatus"] = "skipped"
                entry["reason"] = result.get("error", "no output produced")
                validation["results"].append(entry)
                continue

            validator = VALIDATORS.get(case_name)
            if validator is None:
                entry["excelStatus"] = "skipped"
                entry["reason"] = "no xlwings validator for this case"
                validation["results"].append(entry)
                continue

            try:
                entry["excelStatus"] = "ok"
                entry["excel"] = validator(app, Path(result["outputPath"]))
            except Exception as exc:
                entry["excelStatus"] = "error"
                entry["error"] = str(exc)
            validation["results"].append(entry)
    finally:
        app.quit()

    validation["validatedAt"] = datetime.now(UTC).isoformat()
    output_path = Path(artifact_results["outputRoot"]) / "excel_validation.json"
    output_path.write_text(json.dumps(validation, indent=2) + "\n")
    print(json.dumps({"validationPath": str(output_path), "results": validation["results"]}, indent=2))


if __name__ == "__main__":
    main()
