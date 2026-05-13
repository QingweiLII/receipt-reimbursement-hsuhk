# Local Run

This mode runs the receipt app on this Mac instead of Render. The link works only while the Mac is on and the Terminal window is running.

## Same-Mac Link

Double-click:

```text
start-local.command
```

Then open:

```text
http://127.0.0.1:8000/hsuhk-receipt-report-page
```

To stop it, press `Ctrl+C` in the Terminal window.

## Temporary Public Link

If you need a shareable link while this Mac is running, install Cloudflare Tunnel once:

```bash
brew install cloudflared
```

Then double-click:

```text
start-temporary-link.command
```

Copy the `https://...trycloudflare.com` URL printed by Cloudflare and add:

```text
/hsuhk-receipt-report-page
```

That temporary link stops working when the Terminal window closes or this Mac shuts down.

## Local Defaults

The local launcher forces:

- `STORAGE_PROVIDER=local`
- one upload worker at a time
- longer OCR and LLM timeouts than Render
- more OCR image variants for difficult receipt photos

It still reads `.env` for the MiniMax key/model, but it does not depend on Render.
