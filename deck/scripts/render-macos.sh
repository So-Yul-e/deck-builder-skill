#!/usr/bin/env bash
set -eu

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DECK_PY="$(bash "$SCRIPT_DIR/deps-macos.sh" --python)"

if [ "$#" -lt 1 ]; then
  echo "Usage: bash render-macos.sh input.pptx [output.pdf]" >&2
  exit 2
fi

SRC="$1"
OUT="${2:-${SRC%.pptx}.pdf}"
if [ ! -f "$SRC" ]; then
  echo "Missing file: $SRC" >&2
  exit 1
fi

SRC_ABS="$(cd "$(dirname "$SRC")" && pwd)/$(basename "$SRC")"
mkdir -p "$(dirname "$OUT")"
OUT_ABS="$(cd "$(dirname "$OUT")" && pwd)/$(basename "$OUT")"
TMP_BASE="${TMPDIR:-/tmp}"

cleanup() {
  if [ -n "${JOB_DIR:-}" ] && [ -d "$JOB_DIR" ]; then
    rm -rf "$JOB_DIR"
  fi
  JOB_DIR=""
}
trap cleanup EXIT

verify_pdf() {
  "$DECK_PY" "$SKILL_DIR/scripts/verify_pdf_fonts.py" "$SRC_ABS" "$OUT_ABS"
}

render_with_libreoffice() {
  cleanup
  JOB_DIR="$(mktemp -d "$TMP_BASE/deck-render.XXXXXX")"
  PROFILE="$JOB_DIR/profile"
  OUTDIR="$JOB_DIR/out"
  mkdir -p "$PROFILE" "$OUTDIR"
  if soffice "-env:UserInstallation=file://$PROFILE" --headless --convert-to pdf --outdir "$OUTDIR" "$SRC_ABS" >/dev/null 2>&1; then
    LO_PDF="$OUTDIR/$(basename "${SRC_ABS%.pptx}.pdf")"
    if [ -f "$LO_PDF" ]; then
      mv -f "$LO_PDF" "$OUT_ABS"
      if verify_pdf; then
        printf '%s\n' "$OUT_ABS"
        return 0
      fi
    fi
  fi
  return 1
}

render_with_powerpoint() {
  if [ ! -d "/Applications/Microsoft PowerPoint.app" ] && [ ! -d "$HOME/Applications/Microsoft PowerPoint.app" ]; then
    return 1
  fi

  cleanup
  JOB_DIR="$(mktemp -d "$TMP_BASE/deck-render.XXXXXX")"
  PPT_PDF="$JOB_DIR/out.pdf"
  # One attempt only. Retrying reopened PowerPoint windows and recovery prompts.
  for _ in 1; do
    rm -f "$PPT_PDF"
    osascript - "$SRC_ABS" "$PPT_PDF" >/dev/null 2>&1 <<'APPLESCRIPT' || true
on run argv
set sourcePath to item 1 of argv
set outputPath to item 2 of argv
set outputFile to POSIX file outputPath
tell application "Microsoft PowerPoint"
  open POSIX file sourcePath
  delay 8
  set p to active presentation
  save p in outputFile as save as PDF
  delay 6
  close p saving no
end tell
end run
APPLESCRIPT
    if [ -f "$PPT_PDF" ]; then
      mv -f "$PPT_PDF" "$OUT_ABS"
      if verify_pdf; then
        printf '%s\n' "$OUT_ABS"
        return 0
      fi
    fi
    osascript -e 'tell application "Microsoft PowerPoint" to quit saving no' >/dev/null 2>&1 || true
    sleep 5
  done
  return 1
}

if command -v soffice >/dev/null 2>&1 && render_with_libreoffice; then
  exit 0
fi

# PowerPoint can open windows, macOS automation consent prompts and recovery dialogs,
# so it is never an automatic fallback. Use it only when the user asked for it.
if [ "${DECK_ALLOW_POWERPOINT:-0}" != "1" ]; then
  echo "Render failed: LibreOffice could not export the deck. PowerPoint fallback is disabled." >&2
  echo "Only if the user explicitly approves: DECK_ALLOW_POWERPOINT=1 bash $0 $SRC" >&2
  exit 1
fi

if command -v osascript >/dev/null 2>&1 && render_with_powerpoint; then
  exit 0
fi

echo "Render failed: LibreOffice or Microsoft PowerPoint was not able to export the deck." >&2
exit 1
