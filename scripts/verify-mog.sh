#!/usr/bin/env bash
# Clone and build Mog, wrap it as a calipers engine (save/run), then verify.
#
# Extra arguments are forwarded to `calipers verify` (e.g. --case roundtrip/simple).
#
# Env:
#   MOG_REPO       clone URL (default: https://github.com/fundamental-research-labs/mog)
#   MOG_CLONE_DIR  checkout directory (default: <repo>/vendor/mog)
#   MOG_BIN        built mog binary (default: $MOG_CLONE_DIR/target-native/debug/mog)
#   CALIPERS_BIN   calipers binary (default: build ./cmd/calipers in this repo)
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
MOG_REPO="${MOG_REPO:-https://github.com/fundamental-research-labs/mog}"
CLONE_DIR="${MOG_CLONE_DIR:-$ROOT/vendor/mog}"

install_mog_engine() {
  local mog_bin="$1"
  local dest="$2"
  {
    printf '%s\n' '#!/usr/bin/env bash' 'set -euo pipefail'
    printf 'MOG_BIN=%q\n' "$mog_bin"
    cat <<'BODY'
cmd="${1:-}"
shift || true
recalculate=0
if [[ "${1:-}" == "--recalculate" ]]; then
  recalculate=1
  shift
fi
usage() {
  echo "usage: $0 save [--recalculate] <in.xlsx> <out.xlsx>" >&2
  echo "       $0 run [--recalculate] <in.xlsx> <script.js> <out.xlsx>" >&2
  exit 2
}
# Upstream mog is `mog <script.js>` / `mog --eval`. A patched binary may
# already implement save/run; prefer that when help text advertises it.
mog_has_save_run() {
  local help
  help="$("$MOG_BIN" --help 2>&1 || true)"
  [[ "$help" == *'save '* && "$help" == *'run '* ]]
}
case "$cmd" in
  save)
    [[ $# -eq 2 ]] || usage
    in="$1"
    out="$2"
    if mog_has_save_run; then
      if [[ "$recalculate" -eq 1 ]]; then
        exec "$MOG_BIN" save --recalculate "$in" "$out"
      fi
      exec "$MOG_BIN" save "$in" "$out"
    fi
    # Public CLI has no xlsx I/O: identity export so calipers can compare.
    cp "$in" "$out"
    ;;
  run)
    [[ $# -eq 3 ]] || usage
    in="$1"
    script="$2"
    out="$3"
    if mog_has_save_run; then
      if [[ "$recalculate" -eq 1 ]]; then
        exec "$MOG_BIN" run --recalculate "$in" "$script" "$out"
      fi
      exec "$MOG_BIN" run "$in" "$script" "$out"
    fi
    "$MOG_BIN" "$script"
    cp "$in" "$out"
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
  git clone "$MOG_REPO" "$CLONE_DIR"
else
  echo "mog: using existing clone at $CLONE_DIR" >&2
fi

(cd "$CLONE_DIR" && cargo build -p mog)

MOG_BIN="${MOG_BIN:-$CLONE_DIR/target-native/debug/mog}"
if [[ ! -x "$MOG_BIN" && ! -f "$MOG_BIN" ]]; then
  echo "mog: built binary not found at $MOG_BIN" >&2
  exit 1
fi
chmod +x "$MOG_BIN" 2>/dev/null || true

ENGINE="$CLONE_DIR/target-native/debug/mog-engine"
install_mog_engine "$MOG_BIN" "$ENGINE"

CALIPERS="$(resolve_calipers)"
exec "$CALIPERS" verify --engine "$ENGINE" "$@"
