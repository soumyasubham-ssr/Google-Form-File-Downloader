"""Main Tkinter window."""

from __future__ import annotations

import os
import queue
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

from . import APP_TITLE, __version__
from .core import (CANCELLED, DEFAULT_WORKERS, MISSING, SKIPPED, SUCCESS, DownloadEngine,
                   format_bytes, plan_jobs, read_rows, remove_unreachable_local_proxy,
                   write_log)
from .dialogs import show_about, show_guide
from .widgets import (ACCENT, AMBER, BG, BORDER, ENTRY_BG, FONT, FONT_BOLD, FONT_MONO,
                      GREEN, LOG_BG, RED, SURFACE, TEXT, TEXT_DIM,
                      TEXT_SEC, Btn, PBar, apply_icon, card, enable_windows_dpi_awareness,
                      label, px, resource_path, round_rect, scrollbar,
                      style_scrollbars)

POLL_MS = 100
STATS = (("Total links", TEXT), ("Completed", GREEN), ("Failed", RED),
         ("Skipped", TEXT_SEC), ("Link Missing", AMBER), ("Remaining", TEXT_SEC))
SPREADSHEET_TYPES = [("Spreadsheets", "*.xlsx *.xlsm *.csv"), ("Excel workbook", "*.xlsx *.xlsm"),
                     ("CSV file", "*.csv")]


def open_path(path: Path):
    """Open a file or folder with the system's default handler."""
    if os.name == "nt":
        os.startfile(str(path))  # noqa: S606
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])


class FileDownloaderApp:
    def __init__(self, root: tk.Tk, proxy_removed: bool = False):
        self.root = root
        self.proxy_removed = proxy_removed
        root.title(f"{APP_TITLE} {__version__}")
        root.configure(bg=BG)
        apply_icon(root)
        style_scrollbars(root)

        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready to download")
        self.speed_text = tk.StringVar(value="Speed: --")
        self.progress_text = tk.StringVar(value="0 of 0 processed")
        self.stat_vars = {name: tk.StringVar(value="0") for name, _ in STATS}

        self.events: queue.Queue = queue.Queue()
        self.engine: DownloadEngine | None = None
        self.running = False
        self._reset_counters()

        self._build_ui()
        self._size_window()
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        root.after(POLL_MS, self._process_events)

    # ── State ──────────────────────────────────────────────────────

    def _reset_counters(self):
        self.started_at = None
        self.total = self.processed = self.active = 0
        self.counts = {SUCCESS: 0, "Failed": 0, SKIPPED: 0, MISSING: 0, CANCELLED: 0}
        self.downloaded_bytes = 0
        self.log_rows: list[dict] = []
        self.output_dir: Path | None = None

    # ── Layout ─────────────────────────────────────────────────────

    def _size_window(self):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        # Leave room for the title bar and taskbar (both scale with DPI).
        avail_w, avail_h = sw - px(self.root, 40), sh - px(self.root, 120)
        min_w, min_h = min(px(self.root, 760), avail_w), min(px(self.root, 540), avail_h)
        ww = max(min_w, min(px(self.root, 1000), avail_w))
        wh = max(min_h, min(px(self.root, 740), avail_h))
        y = max(0, (sh - px(self.root, 80) - wh) // 2)
        self.root.geometry(f"{ww}x{wh}+{(sw - ww) // 2}+{y}")
        self.root.minsize(min_w, min_h)

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG, padx=24, pady=18)
        outer.pack(fill="both", expand=True)
        self._build_header(outer)
        self._build_inputs(outer)
        self._build_stats(outer)
        self._build_progress(outer)
        self._build_log(outer)

    def _build_header(self, parent):
        hdr = tk.Frame(parent, bg=BG)
        hdr.pack(fill="x", pady=(0, 18))

        brand = tk.Frame(hdr, bg=BG)
        brand.pack(side="left", fill="x", expand=True)
        size = px(brand, 34)
        mark = tk.Canvas(brand, width=size, height=size, bg=BG, highlightthickness=0)
        mark.pack(side="left", padx=(0, 11))
        round_rect(mark, 1, 1, size - 1, size - 1, px(brand, 10), fill=ACCENT, outline="")
        mark.create_text(size // 2, size // 2, text="↓", fill="#ffffff", font=(FONT_BOLD, 15))
        titles = tk.Frame(brand, bg=BG)
        titles.pack(side="left", anchor="w")
        label(titles, APP_TITLE, 16, True, anchor="w").pack(anchor="w")
        label(titles, "Reliable bulk file transfer for your form responses",
              fg=TEXT_SEC, anchor="w").pack(anchor="w", pady=(2, 0))

        self.state_label = label(hdr, "●  READY", 8, True, GREEN)
        self.state_label.pack(side="left", padx=(0, 18))

        Btn.secondary(hdr, "Template", self._save_template, w=84).pack(side="right", padx=(5, 0))
        Btn.secondary(hdr, "Guide", lambda: show_guide(self.root), w=62).pack(side="right", padx=(5, 0))
        Btn.secondary(hdr, "About", lambda: show_about(self.root), w=62).pack(side="right")

    def _entry(self, parent, variable):
        return tk.Entry(parent, textvariable=variable, font=(FONT, 9), bg=ENTRY_BG, fg=TEXT,
                        insertbackground=TEXT, relief="flat", highlightthickness=1,
                        highlightbackground=BORDER, highlightcolor=ACCENT,
                        disabledbackground=ENTRY_BG, disabledforeground=TEXT_DIM)

    def _build_inputs(self, parent):
        fc = card(parent, padx=18, pady=16)
        fc.pack(fill="x", pady=(0, 12))
        fc.columnconfigure(1, weight=1)

        label(fc, "SOURCE & DESTINATION", 8, True, ACCENT).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        rows = (("Spreadsheet", self.excel_path, self.browse_excel),
                ("Output Folder", self.output_path, self.browse_folder))
        self.entries, self.browse_buttons = [], []
        for index, (text, variable, command) in enumerate(rows, start=1):
            label(fc, text, 9, True).grid(row=index, column=0, sticky="w", padx=(0, 14), pady=5)
            entry = self._entry(fc, variable)
            entry.grid(row=index, column=1, sticky="ew", pady=5, ipady=6)
            button = Btn.secondary(fc, "Browse", command)
            button.grid(row=index, column=2, padx=(8, 0), pady=5)
            self.entries.append(entry)
            self.browse_buttons.append(button)

        label(fc, "Column A = filename, Column B = link(s), Row 1 = header  •  .xlsx or .csv",
              8, fg=TEXT_DIM).grid(row=3, column=1, columnspan=2, sticky="w", pady=(3, 0))

    def _build_stats(self, parent):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 12))
        for index, (name, color) in enumerate(STATS):
            row.columnconfigure(index, weight=1, uniform="stats")
            c = card(row, padx=12, pady=10)
            c.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 3, 0))
            tk.Label(c, textvariable=self.stat_vars[name], font=(FONT_BOLD, 17),
                     fg=color, bg=SURFACE).pack(anchor="w")
            label(c, name.upper(), 7, True, TEXT_DIM).pack(anchor="w")

    def _build_progress(self, parent):
        pc = card(parent, padx=18, pady=14)
        pc.pack(fill="x", pady=(0, 12))
        pc.columnconfigure(0, weight=1)

        top = tk.Frame(pc, bg=SURFACE)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        tk.Label(top, textvariable=self.status_text, font=(FONT_BOLD, 10),
                 fg=TEXT_SEC, bg=SURFACE).grid(row=0, column=0, sticky="w")
        tk.Label(top, textvariable=self.speed_text, font=(FONT, 9),
                 fg=TEXT_DIM, bg=SURFACE).grid(row=0, column=1, sticky="e")

        self.progress = PBar(pc, h=8)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(5, 3))

        bottom = tk.Frame(pc, bg=SURFACE)
        bottom.grid(row=2, column=0, sticky="ew")
        tk.Label(bottom, textvariable=self.progress_text, font=(FONT, 8),
                 fg=TEXT_DIM, bg=SURFACE).pack(side="left")
        self.start_btn = Btn(bottom, text="Start Download  →", w=160, h=36, r=8,
                             command=self.start_download)
        self.start_btn.pack(side="right", pady=(3, 0))
        self.open_btn = Btn.secondary(bottom, "Open Folder", self._open_output, w=96, h=36)

    def _build_log(self, parent):
        lc = card(parent, padx=14, pady=12)
        lc.pack(fill="both", expand=True)
        label(lc, "ACTIVITY LOG", 8, True, TEXT_SEC).pack(anchor="w", pady=(0, 3))
        frame = tk.Frame(lc, bg=SURFACE)
        frame.pack(fill="both", expand=True)
        self.log = tk.Text(frame, wrap="word", relief="flat", bg=LOG_BG, fg=TEXT_SEC,
                           font=(FONT_MONO, 9), state="disabled", insertbackground=TEXT,
                           selectbackground=ACCENT, highlightthickness=0, padx=10, pady=8)
        for tag, color in (("ok", GREEN), ("err", RED), ("warn", AMBER), ("dim", TEXT_DIM)):
            self.log.tag_configure(tag, foreground=color)
        sb = scrollbar(frame, self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ── Helpers ────────────────────────────────────────────────────

    def _append_log(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now():%H:%M:%S}] ", "dim")
        self.log.insert("end", f"{text}\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _set_state(self, text, color):
        self.state_label.configure(text=f"●  {text}", fg=color)

    def _set_running(self, running: bool):
        self.running = running
        for entry in self.entries:
            entry.configure(state="disabled" if running else "normal")
        for button in self.browse_buttons:
            button.set_enabled(not running)
        if running:
            self.open_btn.pack_forget()
            self.start_btn.configure_button("Cancel", self.cancel_download,
                                            (RED, "#ff8a96", "#d9505e"))
            self._set_state("RUNNING", ACCENT)
        else:
            self.start_btn.configure_button("Start Download  →", self.start_download,
                                            (ACCENT, "#a69cff", "#6659d6"))
            self.start_btn.set_enabled(True)

    # ── Commands ───────────────────────────────────────────────────

    def _save_template(self):
        src = resource_path("assets/template.xlsx")
        if not src.exists():
            messagebox.showerror("Template not found", "The template file could not be located.")
            return
        dest = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel workbook", "*.xlsx")],
                                            initialfile="template.xlsx",
                                            title="Save Excel Template")
        if dest:
            try:
                shutil.copyfile(src, dest)
                messagebox.showinfo("Template saved", f"Template saved to:\n{dest}")
            except OSError as err:
                messagebox.showerror("Save failed", str(err))

    def browse_excel(self):
        selection = filedialog.askopenfilename(title="Select spreadsheet",
                                               filetypes=SPREADSHEET_TYPES)
        if selection:
            self.excel_path.set(selection)
            if not self.output_path.get():
                self.output_path.set(str(Path(selection).parent))

    def browse_folder(self):
        selection = filedialog.askdirectory(title="Select output folder",
                                            initialdir=self.output_path.get() or None)
        if selection:
            self.output_path.set(selection)

    def _open_output(self):
        if self.output_dir and self.output_dir.is_dir():
            open_path(self.output_dir)

    # ── Download orchestration ─────────────────────────────────────

    def start_download(self):
        if self.running:
            return
        sheet = Path(self.excel_path.get().strip().strip('"'))
        output_dir = Path(self.output_path.get().strip().strip('"'))
        if not sheet.is_file():
            messagebox.showerror("Missing information", "Select a valid spreadsheet file.")
            return
        if not output_dir.is_dir():
            messagebox.showerror("Missing information", "Select a valid output folder.")
            return
        try:
            jobs, records = plan_jobs(read_rows(sheet))
        except Exception as error:  # noqa: BLE001 – shown to the user
            messagebox.showerror("Cannot read spreadsheet", str(error))
            return
        if not jobs and not records:
            messagebox.showerror("No download rows", "The spreadsheet contains no rows after the header.")
            return

        self._reset_counters()
        self.output_dir = output_dir
        self.total = len(jobs) + len(records)
        self.started_at = time.monotonic()
        self.progress.set(0)
        self._clear_log()
        self._set_running(True)

        if self.proxy_removed:
            self._append_log("Ignored an invalid localhost:9 proxy setting so downloads can connect directly.", "warn")
        workers = min(DEFAULT_WORKERS, max(1, len(jobs)))
        self._append_log(f"Queue started: {self.total} link(s), up to {workers} simultaneous downloads.")
        for record in records:
            self._handle_result(record)
        self._update_stats()

        self.engine = DownloadEngine(output_dir, on_event=lambda *e: self.events.put(e),
                                     trust_env=not self.proxy_removed)
        threading.Thread(target=self._run_engine, args=(self.engine, jobs),
                         daemon=True, name="gffd-queue").start()

    def _run_engine(self, engine: DownloadEngine, jobs):
        try:
            engine.run(jobs)
        finally:
            self.events.put(("finished", None))

    def cancel_download(self):
        if self.engine and self.running:
            self.engine.cancel()
            self.start_btn.set_enabled(False)
            self.status_text.set("Cancelling – finishing open transfers…")
            self._append_log("Cancel requested. Stopping downloads…", "warn")

    # ── Event processing ───────────────────────────────────────────

    def _process_events(self):
        changed = False
        try:
            while True:
                kind, payload = self.events.get_nowait()
                changed = True
                if kind == "started":
                    self.active += 1
                elif kind == "bytes":
                    self.downloaded_bytes += payload
                elif kind == "result":
                    if payload["Status"] != CANCELLED or payload["Error message"] == "Cancelled by user":
                        self.active = max(0, self.active - 1)
                    self._handle_result(payload)
                elif kind == "finished":
                    self._finish_run()
        except queue.Empty:
            pass
        if changed and self.running:
            self._update_stats()
        self.root.after(POLL_MS, self._process_events)

    def _handle_result(self, record):
        self.log_rows.append(record)
        self.processed += 1
        status = record["Status"]
        self.counts[status] = self.counts.get(status, 0) + 1
        row = f"Row {record['Row']}: "
        if status == SUCCESS:
            self._append_log(f"{row}Downloaded {record['Saved filename']} "
                             f"({format_bytes(record['Bytes downloaded'])})", "ok")
        elif status in (MISSING, SKIPPED):
            self._append_log(f"{row}{status} – {record['Error message']}", "warn")
        elif status == CANCELLED:
            pass  # summarised once at the end
        else:
            self._append_log(f"{row}{status} – {record['Error message']}", "err")

    def _update_stats(self):
        remaining = max(0, self.total - self.processed)
        values = (self.total, self.counts[SUCCESS], self.counts["Failed"], self.counts[SKIPPED],
                  self.counts[MISSING], remaining)
        for variable, value in zip(self.stat_vars.values(), values):
            variable.set(str(value))
        percent = self.processed / self.total * 100 if self.total else 0
        self.progress.set(percent)
        self.progress_text.set(f"{self.processed} of {self.total} processed ({percent:.0f}%)")
        elapsed = max(time.monotonic() - self.started_at, 0.001) if self.started_at else 0.001
        self.speed_text.set(f"{format_bytes(self.downloaded_bytes)}  •  "
                            f"{self.downloaded_bytes / elapsed / 1_048_576:.2f} MB/s  •  "
                            f"{self.active} active")
        if self.engine and self.engine.cancel_event.is_set():
            return
        self.status_text.set("Downloading files…" if remaining else "Finalizing download log…")

    def _finish_run(self):
        if not self.running:
            return
        self._update_stats()
        elapsed = time.monotonic() - self.started_at
        cancelled = self.counts[CANCELLED]
        log_path = self.output_dir / f"download_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        try:
            log_location = str(write_log(log_path, self.log_rows))
        except OSError as error:
            log_location = f"Could not save log: {error}"

        self._set_running(False)
        self.active = 0
        self.open_btn.pack(side="right", padx=(0, 8), pady=(3, 0))
        failed = self.counts["Failed"] + self.counts[MISSING]
        if cancelled:
            self._set_state("CANCELLED", AMBER)
            self.status_text.set("Download run cancelled")
        elif failed:
            self._set_state("DONE WITH ERRORS", AMBER)
            self.status_text.set("Download run completed with errors")
        else:
            self._set_state("COMPLETED", GREEN)
            self.status_text.set("Download run completed")
        self.speed_text.set(f"{format_bytes(self.downloaded_bytes)} in {elapsed:.1f} s")

        summary = (f"Success: {self.counts[SUCCESS]}; failed: {self.counts['Failed']}; "
                   f"skipped: {self.counts[SKIPPED]}; link missing: {self.counts[MISSING]}"
                   + (f"; cancelled: {cancelled}" if cancelled else ""))
        self._append_log(f"Finished. {summary}.", "ok" if not failed and not cancelled else "warn")
        self._append_log(f"Log: {log_location}")

        message = (f"Total links processed: {self.processed}\n"
                   f"Successfully downloaded: {self.counts[SUCCESS]}\n"
                   f"Failed downloads: {self.counts['Failed']}\n"
                   f"Duplicates skipped: {self.counts[SKIPPED]}\n"
                   f"Link missing: {self.counts[MISSING]}\n"
                   + (f"Cancelled: {cancelled}\n" if cancelled else "")
                   + f"Total time: {elapsed:.1f} seconds\n\n"
                   f"Log file:\n{log_location}\n\nOpen the output folder now?")
        if messagebox.askyesno("Download complete", message):
            self._open_output()

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno("Downloads in progress",
                                       "Downloads are still running. Cancel them and exit?"):
                return
            if self.engine:
                self.engine.cancel()
        self.root.destroy()


def main():
    proxy_removed = remove_unreachable_local_proxy()
    enable_windows_dpi_awareness()
    root = tk.Tk()
    FileDownloaderApp(root, proxy_removed=proxy_removed)
    root.mainloop()
