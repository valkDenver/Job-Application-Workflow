#!/bin/bash
# Mac: double-click this file (if it will not open, run once in Terminal:  chmod +x "Dice Auto Apply.command")
# Checks the setup, then asks for the line your AI assistant gave you (or press Enter for a default round).
# The line is saved to a file and checked by dice_apply.py bridge; it is never run by this shell.
cd "$(dirname "$0")"
PY="$HOME/.job-workflow/venv/bin/python"
[ -x "$PY" ] || PY="$(command -v python3)"
if [ -z "$PY" ]; then
  echo "Python 3 is not installed. Install it from python.org, then run this again."; read -r -p "Press Enter to close"; exit 1
fi
"$PY" Tools/dice_apply.py doctor
rc=$?
if [ $rc -eq 1 ]; then
  echo
  read -r -p "Some things need fixing. Fix them now? [y/N] " FIX
  if [ "$FIX" = "y" ] || [ "$FIX" = "Y" ]; then "$PY" Tools/dice_apply.py doctor --fix; fi
  PY="$HOME/.job-workflow/venv/bin/python"; [ -x "$PY" ] || PY="$(command -v python3)"
fi
echo
echo "Paste the line from your assistant, then press Enter."
echo "(Press Enter on its own for a round with today's default settings.)"
read -r -p "> " LINE
printf '%s\n' "$LINE" > .pasted_line.txt
"$PY" Tools/dice_apply.py bridge --line-file .pasted_line.txt
rm -f .pasted_line.txt
echo
read -r -p "Press Enter to close"
