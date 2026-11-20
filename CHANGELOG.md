# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and releases use [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [2.1.0] - 2026-11-20

### Added

- **Cancel button.** Stop a run at any time. Running transfers are aborted and partial files deleted.
- **Multiple links per cell.** Comma-separated uploads (how Google Forms stores multi-file answers) are each downloaded.
- **CSV input** alongside `.xlsx`/`.xlsm`. Cells formatted as hyperlinks now use the link target.
- **Skipped** stat card for duplicate links, and a status badge (Ready / Running / Completed / Done with errors / Cancelled).
- Live download speed based on bytes as they arrive, plus total data downloaded.
- Colour-coded activity log with spreadsheet row numbers. New `Row` column in the CSV log.
- **Open Folder** button and an offer to open the output folder when a run finishes.
- Confirmation before closing the window while downloads are running.
- Clickable email, GitHub and documentation links in the About dialog.
- Automatic retries with backoff for network errors and HTTP 429/5xx.
- New application icon that stays legible at every size, with Windows version metadata in the exe.
- Unit test suite (pytest), CI on Windows and Linux, manual release workflow.
- User Guide, Developer Guide and real screenshots in the documentation.

### Changed

- Code split from one 800-line file into the `gffd` package (`core`, `app`, `dialogs`, `widgets`). The download engine has no UI dependency and is fully tested.
- Replaced **pandas** with direct **openpyxl** reading. Startup is faster and the exe shrank from about 38 MB to about 22 MB.
- One HTTP session per worker thread (connection reuse) instead of a new session per file.
- Setup Guide updated: *Viewer* access is enough (previously asked for *Editor*).
- Window size and canvas-drawn controls now scale with Windows display scaling.
- Dark themed scrollbars on Windows.

### Fixed

- Private Drive files were saved as broken `.html` pages and counted as successes. They're now reported as failures with a clear sharing hint.
- Any URL containing `/d/` or `?id=` (even on non-Google sites) was rewritten to a Drive download URL. Rewriting now applies only to Google hosts.
- The same Drive file linked in different formats (`open?id=` vs `/file/d/…/view`) was downloaded twice.
- Duplicate links were counted as **Failed**.
- File names reserved after a failed download stayed blocked for the rest of the run.
- Numeric file names such as `101` could become `101.0`.
- Reserved Windows names (`CON`, `NUL`, …) and over-long names are now handled safely.
- On 150 %+ display scaling, button labels overflowed and the window could extend under the taskbar.
- `.xls` was offered in the file picker but could not be read (no `xlrd`). The picker now lists only supported formats.
- The DPI-awareness fallback treated a successful `SetProcessDpiAwareness` call as a failure.
- `run_gui_downloader.bat` now works from any current directory and reports dependency installation errors.

## [2.0.0] - 2026-07-25

### Added

- Modern dark enterprise desktop interface.
- High-DPI awareness for sharper Windows rendering across display scaling levels.
- Live progress metrics, speed reporting, and activity logging.
- Setup guide, About dialog, bundled workbook template, and standalone build configuration.

### Improved

- Safer filename handling and duplicate-name protection.
- Google Drive link handling and concurrent download workflow.
- Responsive startup sizing for different screen resolutions.

### Fixed

- Removed an invalid global Tk font override that prevented the application from launching on some Python/Tk versions.

## [1.0.0]

Initial public version.

[2.1.0]: https://github.com/soumyasubham-ssr/Google-Form-File-Downloader/releases/tag/v2.1.0
[2.0.0]: https://github.com/soumyasubham-ssr/Google-Form-File-Downloader/releases/tag/v2.0.0
