# ydisk-bulk-upload

Bulk-upload large file sets to **Yandex Disk** via the official REST API, publish
public share links and export a registry to **Excel** — built for the
"200 000 HTML documents → Disk → links table" use case.

## Features

- **REST API only** — no browser, works on any server with an OAuth token.
- **Resumable** — every file is recorded in a local SQLite ledger; kill and
  restart the run any time, already-uploaded files are skipped.
- **Idempotent / no duplicates** — the key is `relative path + size`; a second
  run never re-uploads or re-publishes.
- **Gentle retries** — exponential backoff on 429/5xx with `Retry-After`
  honouring, so the run survives API throttling.
- **Nested folders preserved** — `src/sub/a.html` → `/disk-root/sub/a.html`.
- **Excel registry** — `file | disk_path | public_url | status`, one row per
  document, failures included for manual follow-up.
- **`--mock-api`** — full pipeline emulation without network, to verify the
  logic on any number of files before touching the real account.

## Verified

- 23 000 files processed end-to-end in mock mode, 0 failures
  (~215 files/s against the mocked API; real-API speed is set by Yandex
  rate limits, roughly hours for 200k files — the tool is built to run
  unattended and resume).
- Re-run on a finished set: all files skipped, no duplicates.
- `examples/sample_report.xlsx` — real output of a 3 000-file mock run.

## Usage

```bash
pip install requests openpyxl

# dry run on your file set (no token, no network)
python ydisk_bulk_upload.py --src ./docs --mock-api --xlsx report.xlsx

# real run (token belongs to the client — OAuth for Yandex Disk)
export YADISK_TOKEN=...
python ydisk_bulk_upload.py --src ./docs --disk-root /incoming_docs \
    --xlsx links_registry.xlsx
```

Interrupts are safe: run the same command again and it continues where it
stopped. Delete `_ydisk_ledger.sqlite` to start fresh.

## Notes for a 200k-file job

- Upload speed is bounded by Yandex API quotas, not by the tool — expect the
  run to take hours; it is designed to be left unattended.
- Check that the target Disk has enough storage quota (200k HTML ≈ several GB).
- Public-link publishing is one API call per file; the ledger keeps every
  result so a partial run still yields a usable registry.

## License

MIT
