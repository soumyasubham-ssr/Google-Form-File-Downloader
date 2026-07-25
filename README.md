# Google Form File Downloader

> A polished Windows desktop utility for downloading files from Excel-based Google Form responses.

Google Form File Downloader reads a spreadsheet containing filenames and file URLs, downloads the files concurrently, safely renames them, and writes a detailed CSV activity log.

## Features

- Download files from Google Drive sharing links and direct URLs
- Process up to 20 downloads concurrently
- Detect duplicate URLs and prevent filename collisions
- Apply safe filenames with automatic deduplication
- Track progress, completion, failures, missing links, and transfer speed
- Save a detailed CSV log in the selected output folder
- Use the bundled Excel template to get started quickly
- Modern dark desktop interface with Windows high-DPI support
- Build a standalone Windows `.exe` with PyInstaller

## Requirements

- Windows 10 or later
- Python 3.9 or later
- Internet access for downloading files
- An `.xlsx` or `.xls` workbook with at least two columns

## Run from source

```powershell
python -m pip install -r requirements.txt
python google_form_file_downloader.py
```

On Windows, you can also run `run_gui_downloader.bat`.

## Prepare the spreadsheet

The first row is treated as a header. The application reads:

| Column | Content |
| --- | --- |
| A | Desired filename. The extension is optional. Leave blank to use the server-provided name. |
| B | Google Drive sharing link or direct download URL. |

Use `assets/template.xlsx` or the **Template** button in the application for a ready-made workbook.

## Use the application

1. Select the Excel workbook.
2. Select the output folder.
3. Click **Start Download**.
4. Review the live activity log and the CSV log saved in the output folder.

For Google Drive files, confirm that the files are accessible with the sharing permissions required by your organization.

## Build the Windows executable

Install PyInstaller and run the included spec file:

```powershell
python -m pip install pyinstaller
pyinstaller "Google Form File Downloader.spec"
```

The executable is generated in `dist/`. The spec file bundles the Excel template and application icon.

## Project layout

```text
.
├── assets/
│   ├── icon.ico
│   └── template.xlsx
├── google_form_file_downloader.py
├── Google Form File Downloader.spec
├── requirements.txt
└── run_gui_downloader.bat
```

## Development checks

Run the Python syntax check before submitting changes:

```powershell
python -m py_compile google_form_file_downloader.py
```

The application is intentionally distributed as a focused desktop utility and does not require a database or external service configuration.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development and pull request guidance.

## Security

Please see [SECURITY.md](SECURITY.md) for reporting security concerns.

## License

Copyright © 2026 Soumya Subham Rout. All rights reserved. See [LICENSE](LICENSE).
