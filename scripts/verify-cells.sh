#!/usr/bin/env bash
# Clone and build the Cells spreadsheet engine, wrap it as a calipers engine
# (save/run), then verify.
#
# Extra arguments are forwarded to `calipers verify` (e.g. --case roundtrip/simple).
#
# Env:
#   CELLS_REPO       clone URL (default: https://github.com/aduermael/cells)
#   CELLS_CLONE_DIR  checkout directory (default: <repo>/vendor/cells)
#   CELLS_BIN        built cells CLI (default: $CELLS_CLONE_DIR/dist/cli/cells)
#   CALIPERS_BIN     calipers binary (default: build ./cmd/calipers in this repo)
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CELLS_REPO="${CELLS_REPO:-https://github.com/aduermael/cells}"
CLONE_DIR="${CELLS_CLONE_DIR:-$ROOT/vendor/cells}"

install_cells_engine() {
  local cells_bin="$1"
  local dest="$2"
  {
    printf '%s\n' '#!/usr/bin/env bash' 'set -euo pipefail'
    printf 'CELLS_BIN=%q\n' "$cells_bin"
    cat <<'BODY'
cmd="${1:-}"
shift || true
eval_flag=()
if [[ "${1:-}" == "--recalculate" ]]; then
  eval_flag=(--eval)
  shift
fi
usage() {
  echo "usage: $0 save [--recalculate] <in.xlsx> <out.xlsx>" >&2
  echo "       $0 run [--recalculate] <in.xlsx> <script.js> <out.xlsx>" >&2
  exit 2
}
# Cells CLI is `cells -i <in> <out>` with optional --script / --eval, not save/run.
case "$cmd" in
  save)
    [[ $# -eq 2 ]] || usage
    if [[ ${#eval_flag[@]} -gt 0 ]]; then
      exec "$CELLS_BIN" -y -i "$1" "$2" "${eval_flag[@]}"
    fi
    exec "$CELLS_BIN" -y -i "$1" "$2"
    ;;
  run)
    [[ $# -eq 3 ]] || usage
    if [[ ${#eval_flag[@]} -gt 0 ]]; then
      exec "$CELLS_BIN" -y -i "$1" "$3" --script "$2" "${eval_flag[@]}"
    fi
    exec "$CELLS_BIN" -y -i "$1" "$3" --script "$2"
    ;;
  *)
    usage
    ;;
esac
BODY
  } >"$dest"
  chmod +x "$dest"
}

resolve_calipers() {
  if [[ -n "${CALIPERS_BIN:-}" ]]; then
    printf '%s\n' "$CALIPERS_BIN"
    return
  fi
  if [[ -x "$ROOT/calipers" ]]; then
    printf '%s\n' "$ROOT/calipers"
    return
  fi
  (cd "$ROOT" && go build -o calipers ./cmd/calipers)
  printf '%s\n' "$ROOT/calipers"
}

if [[ ! -d "$CLONE_DIR/.git" ]]; then
  mkdir -p "$(dirname "$CLONE_DIR")"
  git clone "$CELLS_REPO" "$CLONE_DIR"
else
  echo "cells: using existing clone at $CLONE_DIR" >&2
fi

# Documented headless/no-collab CLI build; copies dist/cli/cells.
(cd "$CLONE_DIR" && bazel run :cli-headless-no-collab)

CELLS_BIN="${CELLS_BIN:-$CLONE_DIR/dist/cli/cells}"
if [[ ! -x "$CELLS_BIN" && ! -f "$CELLS_BIN" ]]; then
  echo "cells: built CLI not found at $CELLS_BIN" >&2
  exit 1
fi
chmod +x "$CELLS_BIN" 2>/dev/null || true

ENGINE="$CLONE_DIR/dist/cli/cells-engine"
install_cells_engine "$CELLS_BIN" "$ENGINE"

CALIPERS="$(resolve_calipers)"
exec "$CALIPERS" verify --engine "$ENGINE" "$@"
