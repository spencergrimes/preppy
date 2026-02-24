#!/bin/zsh
set -e

cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
pip install -q -r requirements.txt

URL="http://127.0.0.1:5000"

echo "Starting Preppy at $URL"
python app.py &
APP_PID=$!

cleanup() {
  if ps -p "$APP_PID" >/dev/null 2>&1; then
    kill "$APP_PID" >/dev/null 2>&1 || true
  fi
}

trap cleanup EXIT INT TERM

sleep 1
open "$URL"

echo "Preppy is running. Close this window to stop it."
wait "$APP_PID"
