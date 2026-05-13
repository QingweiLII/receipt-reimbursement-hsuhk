# Receipt Reimbursement

A small receipt upload app for reimbursement tracking.

It accepts images and PDFs, sends them to a configurable LLM API, extracts reimbursement fields, and generates a reimbursement workbook.

Each browser gets its own local `client_id`. Uploads and downloads are stored under `data/clients/<client_id>/`, so different colleagues using the same link get separate Excel files. The upload form accepts multiple files and processes them in parallel.
Uploads run as background jobs. The browser gets a job id immediately and polls for completion, so slow OCR/LLM/Drive work does not hold one long HTTP request open.

## Run locally

```bash
cd receipt-reimbursement
python3 app.py
```

Open `http://127.0.0.1:8000`.

The named page route is also available at:

```text
http://127.0.0.1:8000/hsuhk-receipt-report-page
```

## Configure an LLM API

The app reads `receipt-reimbursement/.env` automatically. For MiniMax, paste your key here:

```env
LLM_PROVIDER=minimax
LLM_BASE_URL=https://api.minimax.io/anthropic
LLM_API_KEY=your-minimax-api-key
LLM_MODEL=MiniMax-M2.7
LLM_TEMPERATURE=0.1
LLM_MAX_TOKENS=2500
```

For MiniMax accounts/keys from the China developer platform, use:

```env
LLM_BASE_URL=https://api.minimaxi.com/anthropic
```

MiniMax's compatible text endpoint does not accept image/PDF bytes directly in this app, so this provider first runs local OCR/PDF text extraction and then sends the extracted text to MiniMax.

You can also set environment variables before starting:

```bash
export LLM_PROVIDER=openai
export LLM_API_KEY=...
export LLM_MODEL=gpt-4o-mini
python3 app.py
```

OpenAI-compatible APIs use the same request shape:

```bash
export LLM_PROVIDER=openai_compatible
export LLM_BASE_URL=https://your-provider.example/v1
export LLM_API_KEY=...
export LLM_MODEL=your-vision-model
python3 app.py
```

Anthropic and Gemini adapters are also included:

```bash
export LLM_PROVIDER=anthropic
export LLM_API_KEY=...
export LLM_MODEL=claude-3-5-sonnet-latest
python3 app.py
```

```bash
export LLM_PROVIDER=gemini
export LLM_API_KEY=...
export LLM_MODEL=gemini-1.5-pro
python3 app.py
```

## Public link mode

Run on a public host with:

```bash
export HOST=0.0.0.0
export PORT=8000
python3 app.py
```

Then share:

```text
https://your-domain.example/hsuhk-receipt-report-page
```

If you later want a simple shared password, set `APP_TOKEN` and open the page with `?token=...`.

## Per-user Google Drive storage

For cloud deployment where each colleague signs in with their own Google account, use OAuth-backed Google Drive storage:

```env
STORAGE_PROVIDER=user_google_drive
GOOGLE_OAUTH_CLIENT_ID=your-oauth-client-id.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=your-oauth-client-secret
GOOGLE_DRIVE_FOLDER_NAME=HSUHK Receipt Reports
```

In Google Cloud, create an OAuth Client ID for a Web application, enable the Google Drive API, and add this authorized redirect URI:

```text
https://your-render-url.onrender.com/auth/google/callback
```

Users click `Connect Google Drive` in the page. The app can then list image/PDF receipt files from their Drive and save generated Excel files into a folder named `HSUHK Receipt Reports` in their own Drive.

The default OAuth scopes are:

```text
openid email profile
https://www.googleapis.com/auth/drive.readonly
https://www.googleapis.com/auth/drive.file
```

## Cloud deployment

The repository includes a Dockerfile and `render.yaml` for Render-style deployment.

For the first Render deploy, provide only:

```env
LLM_API_KEY=your-minimax-api-key
```

Render will then give you a stable URL, usually like `https://your-service-name.onrender.com`. In Google Cloud, create a Web OAuth client and add this redirect URI:

```text
https://your-service-name.onrender.com/auth/google/callback
```

Then add these Render environment variables and redeploy:

```env
GOOGLE_OAUTH_CLIENT_ID=your-oauth-client-id.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=your-oauth-client-secret
```

`render.yaml` already sets the non-secret defaults for MiniMax, per-user Google Drive storage, OCR, and the app route.
On small Render instances, keep `MAX_PARALLEL_RECEIPTS` low. The included Render config uses `UPLOAD_JOB_WORKERS=1` and `MAX_PARALLEL_RECEIPTS=2` so one upload can continue in the background without exhausting the instance.

`PUBLIC_BASE_URL` is optional. Set it only if your host does not report the public HTTPS URL correctly, or after you attach a custom domain.

## Type and HKD conversion

The workbook includes a `Type` column before `Activities`. Values are normalized to:

```text
Flight, Meal, Accommondation, Transportation, Others
```

The app fills `Amount in HKD` from the receipt date using Frankfurter historical exchange rates by default. Frankfurter does not require an API key.

```bash
export EXCHANGE_RATE_PROVIDER=frankfurter
export FX_BASE_URL=https://api.frankfurter.dev/v2
python3 app.py
```

If the exact receipt date has no published rate, the app looks back up to 7 days for the nearest prior available rate. Change this with:

```bash
export FX_LOOKBACK_DAYS=10
```

Static fallback mode is also available:

```bash
export EXCHANGE_RATE_PROVIDER=static
export FX_RATES_JSON='{"USD":7.8,"CNY":1.08,"DKK":1.14,"SGD":5.75}'
```

## Sample fixture import

The included workspace sample receipts can be loaded without any API key:

```bash
python3 app.py import-samples --reset
```

This uses `LLM_PROVIDER=fixture` logic for the five sample filenames only. Real uploads require an actual provider.
