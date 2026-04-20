#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CODEX_NODE="${CODEX_NODE:-/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node}"
CODEX_NODE_MODULES="${CODEX_NODE_MODULES:-/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules}"

ln -sfn "$CODEX_NODE_MODULES" "$ROOT/node_modules"
"$CODEX_NODE" "$ROOT/scripts/run_codex_cases.mjs" "$@"
uv run --with xlwings python "$ROOT/scripts/validate_excel.py" "$ROOT/outputs/artifact_results.json"
