# Contributing

Thanks for helping improve Google Form File Downloader.

## Before you start

- Keep changes focused and explain the user-facing impact.
- Do not commit spreadsheets containing personal data, downloaded files, credentials, or private URLs.
- Preserve the existing download workflow and file-handling behavior unless a change explicitly targets it.

## Local setup

```powershell
python -m pip install -r requirements.txt
python google_form_file_downloader.py
```

## Validation

Run the syntax check before opening a pull request:

```powershell
python -m py_compile google_form_file_downloader.py
```

If you change packaging, also test the PyInstaller build:

```powershell
python -m pip install pyinstaller
pyinstaller "Google Form File Downloader.spec"
```

## Pull requests

Include:

- A concise description of the problem and solution.
- Screenshots for visual changes.
- Verification steps and their results.
- Any known limitations or follow-up work.
