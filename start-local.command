#!/bin/zsh
set -e

cd -- "$(dirname "$0")"

PYTHON_BIN="/Users/qingweili/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
if [[ ! -x "$PYTHON_BIN" ]]; then
  PYTHON_BIN="$(command -v python3 || true)"
fi
if [[ -z "$PYTHON_BIN" ]]; then
  echo "python3 was not found. Install Python 3 first."
  exit 1
fi

export HOST="${LOCAL_HOST:-127.0.0.1}"
export PORT="${PORT:-8000}"
export STORAGE_PROVIDER="local"
export APP_TOKEN="${LOCAL_APP_TOKEN:-}"

# Local machines can afford longer work than Render's free web service.
export LLM_TIMEOUT_SECONDS="${LLM_TIMEOUT_SECONDS:-120}"
export RECEIPT_PROCESS_TIMEOUT_SECONDS="${RECEIPT_PROCESS_TIMEOUT_SECONDS:-600}"
export OCR_TIMEOUT_SECONDS="${OCR_TIMEOUT_SECONDS:-60}"
export OCR_MAX_LONG_EDGE="${OCR_MAX_LONG_EDGE:-1400}"
export OCR_SMALL_MAX_LONG_EDGE="${OCR_SMALL_MAX_LONG_EDGE:-900}"
export OCR_MAX_CANDIDATES="${OCR_MAX_CANDIDATES:-6}"
export OCR_VARIANTS="${OCR_VARIANTS:-gray,rgb,rotations}"
export OCR_ORIENTATION_MAX_LONG_EDGE="${OCR_ORIENTATION_MAX_LONG_EDGE:-1800}"
export MAX_PARALLEL_RECEIPTS="${MAX_PARALLEL_RECEIPTS:-1}"
export UPLOAD_JOB_WORKERS="${UPLOAD_JOB_WORKERS:-1}"
export MAX_FILES_PER_UPLOAD="${MAX_FILES_PER_UPLOAD:-20}"

if command -v lsof >/dev/null 2>&1 && lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  echo "Port $PORT is already in use."
  echo "Close the other local server, or run: PORT=8001 ./start-local.command"
  exit 1
fi

URL="http://${HOST}:${PORT}/hsuhk-receipt-report-page"
echo ""
echo "Receipt Reimbursement local version"
echo "Open: $URL"
echo "Stop: press Ctrl+C in this terminal window"
echo ""

if [[ "${OPEN_BROWSER:-1}" == "1" ]] && command -v open >/dev/null 2>&1; then
  (sleep 2; open "$URL" >/dev/null 2>&1 || true) &
fi

exec "$PYTHON_BIN" app.py
