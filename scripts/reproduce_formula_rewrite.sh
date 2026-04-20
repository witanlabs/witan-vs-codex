#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CODEX_NODE="${CODEX_NODE:-/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node}"
CODEX_NODE_MODULES="${CODEX_NODE_MODULES:-/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules}"

ln -sfn "$CODEX_NODE_MODULES" "$ROOT/node_modules"

rm -f \
  "$ROOT/fixtures/case8_rename_fixture.xlsx" \
  "$ROOT/fixtures/case9_shift_fixture.xlsx" \
  "$ROOT/outputs/case8_codex.xlsx" \
  "$ROOT/outputs/case8_witan.xlsx" \
  "$ROOT/outputs/case8_witan.json" \
  "$ROOT/outputs/case9_codex.xlsx" \
  "$ROOT/outputs/case9_witan.xlsx" \
  "$ROOT/outputs/case9_witan.json" \
  "$ROOT/outputs/formula_rewrite_codex_results.json" \
  "$ROOT/outputs/formula_rewrite_excel_validation.json"

witan xlsx exec "$ROOT/fixtures/case8_rename_fixture.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case8_build_fixture_witan.js" \
  --json > "$ROOT/outputs/case8_build_fixture.json"

witan xlsx exec "$ROOT/fixtures/case9_shift_fixture.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case9_build_fixture_witan.js" \
  --json > "$ROOT/outputs/case9_build_fixture.json"

"$CODEX_NODE" "$ROOT/scripts/run_formula_rewrite_cases.mjs"

cp "$ROOT/fixtures/case8_rename_fixture.xlsx" "$ROOT/outputs/case8_witan.xlsx"
witan xlsx exec "$ROOT/outputs/case8_witan.xlsx" \
  --save \
  --script "$ROOT/scripts/case8_witan.js" \
  --json > "$ROOT/outputs/case8_witan.json"

cp "$ROOT/fixtures/case9_shift_fixture.xlsx" "$ROOT/outputs/case9_witan.xlsx"
witan xlsx exec "$ROOT/outputs/case9_witan.xlsx" \
  --save \
  --script "$ROOT/scripts/case9_witan.js" \
  --json > "$ROOT/outputs/case9_witan.json"

uv run --with xlwings python "$ROOT/scripts/validate_formula_rewrite.py"
