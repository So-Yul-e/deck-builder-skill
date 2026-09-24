#!/usr/bin/env bash
# Narrow PPTX -> PDF render wrapper for sandboxed agents (Codex on macOS).
# LibreOffice does not run inside the Codex sandbox, so the render must run outside it.
# Approving this wrapper instead of a free-form command limits that run to a .pptx input
# and a .pdf output inside one trusted project. It never launches PowerPoint.
set -euo pipefail

usage() {
  echo "Usage: bash render-safe-macos.sh input.pptx [output.pdf]" >&2
}

[ "$#" -ge 1 ] && [ "$#" -le 2 ] || {
  usage
  exit 2
}

SRC="$1"
OUT="${2:-${SRC%.pptx}.pdf}"

case "$SRC" in
  *.pptx) ;;
  *) echo "Input must be a .pptx file: $SRC" >&2; exit 2 ;;
esac
case "$OUT" in
  *.pdf) ;;
  *) echo "Output must be a .pdf file: $OUT" >&2; exit 2 ;;
esac

[ -f "$SRC" ] || { echo "Missing input file: $SRC" >&2; exit 1; }
[ ! -L "$SRC" ] || { echo "Symbolic link input is not allowed: $SRC" >&2; exit 1; }
[ ! -e "$OUT" ] || [ ! -L "$OUT" ] || {
  echo "Symbolic link output is not allowed: $OUT" >&2
  exit 1
}

absolute_existing_file() {
  file="$1"
  dir="$(cd "$(dirname "$file")" && pwd -P)" || return 1
  printf '%s/%s\n' "$dir" "$(basename "$file")"
}

absolute_output_file() {
  file="$1"
  dir="$(cd "$(dirname "$file")" && pwd -P)" || return 1
  printf '%s/%s\n' "$dir" "$(basename "$file")"
}

SRC_ABS="$(absolute_existing_file "$SRC")" || {
  echo "Cannot resolve input path: $SRC" >&2
  exit 1
}
OUT_ABS="$(absolute_output_file "$OUT")" || {
  echo "Output directory does not exist: $(dirname "$OUT")" >&2
  exit 1
}

if REPO_ROOT="$(git -C "$(dirname "$SRC_ABS")" rev-parse --show-toplevel 2>/dev/null)"; then
  PROJECT_ROOT="$(cd "$REPO_ROOT" && pwd -P)"
else
  TRUSTED_BASE="$HOME/workspace"
  [ -d "$TRUSTED_BASE" ] || {
    echo "Not inside a Git project and ~/workspace does not exist: $SRC_ABS" >&2
    exit 1
  }
  TRUSTED_BASE="$(cd "$TRUSTED_BASE" && pwd -P)"
  case "$SRC_ABS" in
    "$TRUSTED_BASE"/*/*) ;;
    *)
      echo "Input outside a Git project must be under ~/workspace/<project>/: $SRC_ABS" >&2
      exit 1
      ;;
  esac
  RELATIVE_SRC="${SRC_ABS#"$TRUSTED_BASE"/}"
  PROJECT_NAME="${RELATIVE_SRC%%/*}"
  PROJECT_ROOT="$TRUSTED_BASE/$PROJECT_NAME"
fi

case "$SRC_ABS" in
  "$PROJECT_ROOT"/*) ;;
  *) echo "Input is outside the project boundary: $SRC_ABS" >&2; exit 1 ;;
esac
case "$OUT_ABS" in
  "$PROJECT_ROOT"/*) ;;
  *) echo "Output is outside the project boundary: $OUT_ABS" >&2; exit 1 ;;
esac

command -v soffice >/dev/null 2>&1 || {
  echo "LibreOffice is not installed." >&2
  exit 2
}
command -v unzip >/dev/null 2>&1 || {
  echo "unzip is not installed." >&2
  exit 2
}

ppt_uses_pretendard() {
  unzip -p "$SRC_ABS" 'ppt/slides/slide*.xml' 2>/dev/null | grep -qi 'typeface="Pretendard"'
}

verify_pdf_fonts() {
  pdf="$1"
  [ -f "$pdf" ] || return 1
  if ! ppt_uses_pretendard || ! command -v pdffonts >/dev/null 2>&1; then
    return 0
  fi
  font_report="$(pdffonts "$pdf" 2>/dev/null || true)"
  echo "$font_report" | grep -qi 'Pretendard' || {
    echo "PDF font check failed: Pretendard is not embedded." >&2
    return 1
  }
  if echo "$font_report" | grep -Eqi 'ArialNarrow|STHeiti'; then
    echo "PDF font check failed: a substitute font with different widths was found." >&2
    return 1
  fi
}

TMP_BASE="/private/tmp"
[ -d "$TMP_BASE" ] || TMP_BASE="/tmp"
JOB_DIR="$(mktemp -d "$TMP_BASE/deck-render-safe.XXXXXX")"
trap 'case "$JOB_DIR" in /private/tmp/deck-render-safe.*|/tmp/deck-render-safe.*) rm -rf -- "$JOB_DIR" ;; esac' EXIT

MAX_ATTEMPTS=2
attempt=1
while [ "$attempt" -le "$MAX_ATTEMPTS" ]; do
  ATTEMPT_DIR="$JOB_DIR/attempt-$attempt"
  PROFILE="$ATTEMPT_DIR/profile"
  OUT_DIR="$ATTEMPT_DIR/out"
  LOG="$ATTEMPT_DIR/soffice.log"
  mkdir -p "$PROFILE" "$OUT_DIR"

  set +e
  soffice "-env:UserInstallation=file://$PROFILE" --headless --nologo \
    --nodefault --nolockcheck --nofirststartwizard \
    --convert-to pdf --outdir "$OUT_DIR" "$SRC_ABS" >"$LOG" 2>&1
  RENDER_RC=$?
  set -e

  if [ "$RENDER_RC" -eq 0 ]; then
    GENERATED="$OUT_DIR/$(basename "${SRC_ABS%.pptx}.pdf")"
    if [ -f "$GENERATED" ]; then
      if verify_pdf_fonts "$GENERATED"; then
        mv -f "$GENERATED" "$OUT_ABS"
        echo "LibreOffice runs: $attempt" >&2
        echo "$OUT_ABS"
        exit 0
      fi
      echo "A font check failure does not recover on retry; stopping." >&2
      break
    fi
  fi

  echo "LibreOffice render failed $attempt/$MAX_ATTEMPTS (exit=$RENDER_RC)" >&2
  tail -20 "$LOG" >&2 || true

  if [ "$attempt" -ge "$MAX_ATTEMPTS" ]; then
    break
  fi
  if [ "$RENDER_RC" -ge 128 ]; then
    echo "Treating this as a process crash; not retrying." >&2
    break
  fi
  if [ "$RENDER_RC" -ne 0 ] && \
    ! grep -Eqi 'lock|locked|busy|profile|another instance|user installation|temporary' "$LOG"; then
    echo "Not a recoverable transient error; not retrying." >&2
    break
  fi

  echo "Transient error; retrying once." >&2
  sleep 8
  attempt=$((attempt + 1))
done

echo "LibreOffice render stopped. Do not rerun it automatically: $SRC_ABS" >&2
exit 1
