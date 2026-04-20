#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CODEX_NODE="${CODEX_NODE:-/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node}"
CODEX_NODE_MODULES="${CODEX_NODE_MODULES:-/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules}"

ln -sfn "$CODEX_NODE_MODULES" "$ROOT/node_modules"

rm -f \
  "$ROOT/fixtures/case10_formula_fixture.xlsx" \
  "$ROOT/outputs/case10_build_fixture.json" \
  "$ROOT/outputs/case10_codex.xlsx" \
  "$ROOT/outputs/case10_witan.xlsx" \
  "$ROOT/outputs/case10_witan.json" \
  "$ROOT/outputs/formula_correctness_codex_results.json" \
  "$ROOT/outputs/formula_correctness_excel_validation.json"

witan xlsx exec "$ROOT/fixtures/case10_formula_fixture.xlsx" \
  --create \
  --save \
  --script "$ROOT/scripts/case10_build_fixture_witan.js" \
  --json > "$ROOT/outputs/case10_build_fixture.json"

"$CODEX_NODE" "$ROOT/scripts/run_formula_correctness_cases.mjs"

cp "$ROOT/fixtures/case10_formula_fixture.xlsx" "$ROOT/outputs/case10_witan.xlsx"
witan xlsx exec "$ROOT/outputs/case10_witan.xlsx" \
  --save \
  --script "$ROOT/scripts/case10_witan.js" \
  --json > "$ROOT/outputs/case10_witan.json"

uv run --with xlwings python "$ROOT/scripts/validate_formula_correctness.py"
