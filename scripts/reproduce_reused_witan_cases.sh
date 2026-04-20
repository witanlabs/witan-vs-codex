#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"

rm -f \
  "$ROOT/fixtures/circular.xlsx" \
  "$ROOT/fixtures/report.xlsx" \
  "$ROOT/fixtures/report_spillref.xlsx" \
  "$ROOT/fixtures/review.xlsx" \
  "$ROOT/outputs/case2_build_fixture.json" \
  "$ROOT/outputs/case2_witan.json" \
  "$ROOT/outputs/case2_witan.xlsx" \
  "$ROOT/outputs/case3_witan.json" \
  "$ROOT/outputs/case4_build_fixture.json" \
  "$ROOT/outputs/case4_witan.json" \
  "$ROOT/outputs/case4_witan.xlsx" \
  "$ROOT/outputs/case5_build_fixture.json" \
  "$ROOT/outputs/case5_witan.json" \
  "$ROOT/outputs/case5_witan.xlsx" \
  "$ROOT/outputs/case7_build_fixture.json" \
  "$ROOT/outputs/case7_witan.json" \
  "$ROOT/outputs/reused_witan_excel_validation.json"

witan xlsx exec "$ROOT/fixtures/circular.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case2_witan_build.js" \
  --json > "$ROOT/outputs/case2_build_fixture.json"

witan xlsx exec "$ROOT/fixtures/review.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case4_build.js" \
  --json > "$ROOT/outputs/case4_build_fixture.json"

witan xlsx exec "$ROOT/fixtures/report.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case5_build.js" \
  --json > "$ROOT/outputs/case5_build_fixture.json"

witan xlsx exec "$ROOT/fixtures/report_spillref.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case7_build.js" \
  --json > "$ROOT/outputs/case7_build_fixture.json"

cp "$ROOT/fixtures/circular.xlsx" "$ROOT/outputs/case2_witan.xlsx"
witan xlsx exec "$ROOT/outputs/case2_witan.xlsx" \
  --save \
  --script "$ROOT/scripts/case2_witan.js" \
  --json > "$ROOT/outputs/case2_witan.json"

witan xlsx exec "$ROOT/fixtures/formulas.xls" \
  --script "$ROOT/scripts/case3_witan.js" \
  --json > "$ROOT/outputs/case3_witan.json"

cp "$ROOT/fixtures/review.xlsx" "$ROOT/outputs/case4_witan.xlsx"
witan xlsx exec "$ROOT/outputs/case4_witan.xlsx" \
  --save \
  --script "$ROOT/scripts/case4_witan.js" \
  --json > "$ROOT/outputs/case4_witan.json"

cp "$ROOT/fixtures/report.xlsx" "$ROOT/outputs/case5_witan.xlsx"
witan xlsx exec "$ROOT/outputs/case5_witan.xlsx" \
  --save \
  --script "$ROOT/scripts/case5_witan.js" \
  --json > "$ROOT/outputs/case5_witan.json"

witan xlsx exec "$ROOT/fixtures/report_spillref.xlsx" \
  --script "$ROOT/scripts/case7_witan.js" \
  --json > "$ROOT/outputs/case7_witan.json"

uv run --with xlwings python "$ROOT/scripts/validate_reused_witan.py"
