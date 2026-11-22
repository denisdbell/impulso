#!/bin/zsh
# ---------------------------------------------------------------------------
# Hourly DNI extraction — fetch new photos from Drive, OCR only the
# NOT-yet-scanned folders (incremental cache), rewrite «Datos DNI (OCR)».
# Shows a macOS popup (notification banner) when it starts and when it ends.
#
# Scheduled by the launchd USER AGENT (hourly at :00, no Full Disk Access
# needed — launchd agents run in the login session, so macOS just shows a
# one-click permission prompt on first run):
#   ~/Library/LaunchAgents/com.impulso.dni-extract.plist
#   launchctl kickstart gui/$UID/com.impulso.dni-extract   # run it now
#   launchctl bootout   gui/$UID/com.impulso.dni-extract   # uninstall
#
# Log: bcra/out/cron_dni.log (plus launchd stray output in out/launchd_dni.log)
# ---------------------------------------------------------------------------

BCRA="$(cd "$(dirname "$0")" && pwd)"
PY=/opt/homebrew/bin/python3.14          # the python with google-auth + Vision
LOG="$BCRA/out/cron_dni.log"
mkdir -p "$BCRA/out"

notify() { /usr/bin/osascript -e "display notification \"$2\" with title \"$1\"" >/dev/null 2>&1 || true; }

notify "Impulso · DNI" "Extracción DNI → hoja: iniciando (solo nuevos)…"
echo "==== $(date '+%Y-%m-%d %H:%M:%S') start ====" >> "$LOG"

OUT="$("$PY" "$BCRA/extract_dni_to_sheet.py" --fetch --incremental 2>&1)"
STATUS=$?
echo "$OUT" >> "$LOG"
echo "==== $(date '+%Y-%m-%d %H:%M:%S') end (exit $STATUS) ====" >> "$LOG"

if [ $STATUS -eq 0 ]; then
  NEW="$(echo "$OUT" | sed -n 's/^OCR: \([0-9][0-9]*\) newly scanned.*/\1/p' | tail -1)"
  WROTE="$(echo "$OUT" | sed -n 's/^Wrote \([0-9][0-9]*\) row(s).*/\1/p' | tail -1)"
  notify "Impulso · DNI" "Listo: ${NEW:-0} nuevo(s) escaneado(s) · ${WROTE:-?} fila(s) en la hoja."
else
  notify "Impulso · DNI" "ERROR (exit $STATUS) — ver bcra/out/cron_dni.log"
fi
exit $STATUS
