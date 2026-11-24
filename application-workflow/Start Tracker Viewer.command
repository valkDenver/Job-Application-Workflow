#!/bin/bash
# Mac: double-click this file (if it will not open, run once in Terminal:  chmod +x "Start Tracker Viewer.command")
cd "$(dirname "$0")"
TRK="$1"
if [ -z "$TRK" ] && [ -f tracker_path.txt ]; then TRK="$(cat tracker_path.txt)"; fi
if [ -z "$TRK" ]; then
  read -r -p "Drag your Master Job Tracker .xlsx file here, then press Enter: " TRK
fi
# Clean what Terminal gives for a dragged file: trailing space, "quotes" or 'quotes', or backslash-escaped spaces (Master\ Job\ Tracker...)
TRK="$(printf '%s' "$TRK" | tr -d '\r' | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')"
case "$TRK" in
  \"*\") TRK="${TRK#\"}"; TRK="${TRK%\"}" ;;
  \'*\') TRK="${TRK#\'}"; TRK="${TRK%\'}" ;;
  *) TRK="$(printf '%s' "$TRK" | sed 's/\\\(.\)/\1/g')" ;;
esac
if [ ! -f "$TRK" ]; then echo "Could not find: $TRK"; read -r -p "Press Enter to close"; exit 1; fi
printf '%s\n' "$TRK" > tracker_path.txt
python3 Tools/tracker_server.py "$TRK" || { echo; echo "The helper did not start. Install Python 3 from python.org, then run once:  pip3 install openpyxl"; read -r -p "Press Enter to close"; }
