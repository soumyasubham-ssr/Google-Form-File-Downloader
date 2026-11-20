# Developer Guide

## Architecture

```text
google_form_file_downloader.py ──► gffd.app.main()
                                     │
         ┌───────────────────────────┼─────────────────────────────┐
         ▼                           ▼                             ▼
   gffd/app.py                 gffd/dialogs.py               gffd/widgets.py
   Main window, event loop,    About + Setup Guide           Theme colours, DPI
   stats, log, cancel                                        scaling, Btn, PBar
         │
         │  read_rows() → plan_jobs() → DownloadEngine.run()  (background thread)
         ▼
   gffd/core.py   ── pure Python, no Tk, fully unit-tested
```

### `gffd/core.py`

| Piece | Responsibility |
| --- | --- |
| `read_rows(path)` | Reads columns A/B of `.xlsx`/`.xlsm` (openpyxl, uses hyperlink targets) or `.csv`. Skips the header and trailing blank rows. |
| `plan_jobs(rows)` | Splits multi-link cells, removes duplicate links (compared after Drive-URL normalisation), and emits `Link Missing`/`Skipped` records up front. |
| `download_url_for(link)` | Turns Drive sharing links on Google hosts into `drive.usercontent.google.com/download?…&confirm=t`. Other hosts pass through unchanged. |
| `DownloadEngine` | Thread pool (default 20 workers). One `requests.Session` per thread with urllib3 `Retry` (3 attempts, backoff, 429/5xx). Streams into a hidden `.part` file, then `os.replace`s it into place. Reserves names under a lock so concurrent jobs never collide, and releases them on failure. Supports cooperative cancel. |
| `on_event(kind, payload)` | Callback used by the engine: `started`, `bytes` (throttled to 4 Hz per worker) and `result`. The UI pushes these into a `queue.Queue` and drains it on the Tk thread every 100 ms. Tk is never touched from worker threads. |
| `write_log(path, records)` | UTF-8-BOM CSV, so Excel opens it with the right encoding. |

### UI notes

- `widgets.px()` converts 96-DPI design sizes to physical pixels, so canvas-drawn controls scale with fonts on high-DPI screens.
- Classic `tk.Scrollbar` ignores colours on Windows, so the app uses a `ttk` scrollbar styled on the `clam` theme.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements-dev.txt
```

## Test

```powershell
python -m pytest
```

The tests start a throw-away local HTTP server, so they need no internet connection. They cover URL parsing, filename sanitising, spreadsheet reading (including hyperlinks), duplicate and missing detection, renaming without overwriting, `.part` cleanup, error reporting and cancel.

## Build

```powershell
pyinstaller --noconfirm "Google Form File Downloader.spec"
```

The spec reads the version from `gffd/__init__.py`, embeds Windows file-version metadata and the icon, bundles `assets/`, and excludes unused heavy modules. The output is `dist\Google Form File Downloader.exe` (≈ 22 MB).

To regenerate the icon (`assets/icon.ico` plus the PNGs in `docs/images/`):

```powershell
python tools/make_icon.py
```

## Release checklist

1. Bump `__version__` in `gffd/__init__.py` and add a section to `CHANGELOG.md`.
2. `python -m pytest`, then build and smoke-test the exe.
3. Merge to `main`.
4. Create the release. Either:
   - **Locally:** `gh release create vX.Y.Z "dist/Google Form File Downloader.exe#Google.Form.File.Downloader.vX.Y.Z.exe" SHA256SUMS.txt --notes-file notes.md`
   - **CI:** run the **Release** workflow (*Actions → Release → Run workflow*) with the tag. It tests, builds on `windows-latest`, and publishes the exe plus `SHA256SUMS.txt`.

## Continuous integration

| Workflow | Trigger | What it does |
| --- | --- | --- |
| `ci.yml` | push / pull request | Runs pytest on Windows and Ubuntu (Python 3.9 and 3.13), and builds the exe on Windows as a downloadable artifact |
| `release.yml` | manual | Tests, builds and publishes a GitHub release for the given tag |
