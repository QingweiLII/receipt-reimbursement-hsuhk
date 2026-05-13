#!/bin/bash
set -e

cd -- "$(dirname "$0")"

if ! command -v cloudflared >/dev/null 2>&1; then
  echo "cloudflared is not installed."
  echo "Install it first with: brew install cloudflared"
  exit 1
fi

export OPEN_BROWSER=0
export HOST=127.0.0.1
export PORT="${PORT:-8000}"

APP_PID=""
if curl -sS "http://127.0.0.1:${PORT}/health" >/dev/null 2>&1; then
  echo "Local app is already running on port ${PORT}."
else
  ./start-local.command &
  APP_PID=$!
fi

cleanup() {
  if [[ -n "$APP_PID" ]]; then
    kill "$APP_PID" >/dev/null 2>&1 || true
  fi
}
trap cleanup EXIT INT TERM

echo "Waiting for local app..."
for _ in {1..30}; do
  if curl -sS "http://127.0.0.1:${PORT}/health" >/dev/null 2>&1; then
    break
  fi
  sleep 1
done

echo ""
echo "Cloudflare will print a temporary https://...trycloudflare.com link below."
echo "Use that domain plus /hsuhk-receipt-report-page"
echo "Stop: press Ctrl+C in this terminal window"
echo ""

cloudflared tunnel --url "http://127.0.0.1:${PORT}"
