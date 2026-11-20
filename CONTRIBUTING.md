# Contributing

Thanks for helping improve Google Form File Downloader.

## Before you start

- Keep changes focused and explain the user-facing impact.
- Do not commit spreadsheets containing personal data, downloaded files, credentials, or private URLs.
- Preserve the existing download workflow and file-handling behavior unless a change explicitly targets it.

## Local setup

```powershell
python -m pip install -r requirements-dev.txt
python google_form_file_downloader.py
```

See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for the architecture overview.

## Validation

Run the test suite before opening a pull request (no internet connection needed):

```powershell
python -m pytest
```

Keep download, parsing and filename logic in `gffd/core.py` (no Tkinter imports) and add tests for it in `tests/`.

If you change packaging, also test the PyInstaller build:

```powershell
pyinstaller --noconfirm "Google Form File Downloader.spec"
```

## Pull requests

Include:

- A concise description of the problem and solution.
- Screenshots for visual changes.
- Verification steps and their results.
- Any known limitations or follow-up work.
