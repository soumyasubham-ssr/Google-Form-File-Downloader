"""UI-independent download engine.

Everything in this module is pure Python (no Tkinter) so it can be unit tested
and reused from the command line.
"""

from __future__ import annotations

import csv
import mimetypes
import os
import re
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable
from urllib.parse import unquote, urlparse

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from . import __version__

DEFAULT_WORKERS = 20
CHUNK_SIZE = 256 * 1024
TIMEOUT = (15, 180)  # (connect, read) seconds
USER_AGENT = f"Mozilla/5.0 (compatible; GoogleFormFileDownloader/{__version__})"
MAX_FILENAME_LENGTH = 180

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
URL_PATTERN = re.compile(r"https?://[^\s,;]+", re.I)
DRIVE_ID_PATTERNS = (re.compile(r"/d/([\w-]{10,})"), re.compile(r"[?&]id=([\w-]{10,})"))
DRIVE_HOSTS = ("drive.google.com", "docs.google.com", "drive.usercontent.google.com")
WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                    *(f"LPT{i}" for i in range(1, 10))}

LOG_FIELDS = ["Timestamp", "Row", "Source URL", "Status", "Saved filename",
              "Error message", "Bytes downloaded"]

# Result statuses
SUCCESS, FAILED, MISSING, SKIPPED, CANCELLED = (
    "Success", "Failed", "Link Missing", "Skipped", "Cancelled")


# ── Environment ────────────────────────────────────────────────────

def remove_unreachable_local_proxy() -> bool:
    """Drop the known-unusable ``localhost:9`` proxy some systems inject.

    Returns True when a proxy variable was removed.
    """
    removed = False
    for key in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        value = os.environ.get(key, "")
        if not value:
            continue
        parsed = urlparse(value if "://" in value else f"//{value}")
        try:
            port = parsed.port
        except ValueError:
            continue
        if parsed.hostname in {"127.0.0.1", "localhost", "::1"} and port == 9:
            os.environ.pop(key, None)
            removed = True
    return removed


# ── Text helpers ───────────────────────────────────────────────────

def cell_text(value) -> str:
    """Normalise a spreadsheet cell to a trimmed string ('' for empty)."""
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        if value.is_integer():
            value = int(value)
    return str(value).strip()


def parse_links(cell: str) -> list[str]:
    """Return every http(s) URL in *cell*.

    Google Forms stores multiple uploads for one question in a single cell,
    separated by commas, so a cell can yield several links.
    """
    return [m.group(0).rstrip(").]") for m in URL_PATTERN.finditer(cell or "")]


def extract_file_id(link: str) -> str | None:
    """Return the Drive file ID when *link* is a Google Drive/Docs URL."""
    host = (urlparse(link).hostname or "").lower()
    if host not in DRIVE_HOSTS:
        return None
    for pattern in DRIVE_ID_PATTERNS:
        match = pattern.search(link)
        if match:
            return match.group(1)
    return None


def download_url_for(link: str) -> tuple[str, str | None]:
    """Map a sharing link to a direct download URL. Returns (url, drive_id)."""
    drive_id = extract_file_id(link)
    if drive_id:
        return (f"https://drive.usercontent.google.com/download"
                f"?id={drive_id}&export=download&confirm=t"), drive_id
    return link, None


def safe_filename(value, fallback: str = "download") -> str:
    """Make *value* safe to use as a Windows file name."""
    name = Path(str(value).strip()).name
    name = INVALID_FILENAME_CHARS.sub("_", name).strip().rstrip(". ")
    if Path(name).stem.upper() in WINDOWS_RESERVED:
        name = f"_{name}"
    if len(name) > MAX_FILENAME_LENGTH:
        suffix = Path(name).suffix[:16]
        name = name[:MAX_FILENAME_LENGTH - len(suffix)] + suffix
    return name or fallback


def filename_from_headers(headers, final_url: str) -> str:
    """Get a sensible filename from Content-Disposition, then the URL path."""
    disposition = headers.get("content-disposition", "")
    match = (re.search(r"filename\*\s*=\s*UTF-8''([^;]+)", disposition, re.I)
             or re.search(r'filename\s*=\s*"?([^";]+)', disposition, re.I))
    if match:
        return safe_filename(unquote(match.group(1)))
    path_name = Path(unquote(urlparse(final_url).path)).name
    return safe_filename(path_name) if path_name else "download"


def format_bytes(count: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if count < 1024 or unit == "GB":
            return f"{count:.0f} {unit}" if unit == "B" else f"{count:.1f} {unit}"
        count /= 1024
    return f"{count:.1f} GB"


# ── Spreadsheet input ──────────────────────────────────────────────

def read_rows(path: Path) -> list[tuple[str, str]]:
    """Read (filename, link-cell) pairs from an .xlsx/.xlsm or .csv file.

    The first row is a header and is skipped. Only the first two columns are
    used. When the link cell is a hyperlink whose text is not a URL, the
    hyperlink target is used instead.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    rows: list[tuple[str, str]] = []

    if suffix == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as handle:
            reader = csv.reader(handle)
            next(reader, None)
            for record in reader:
                record = (record + ["", ""])[:2]
                rows.append((cell_text(record[0]), cell_text(record[1])))
        return _trim_trailing_blanks(rows)

    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError("Unsupported file type. Use an .xlsx, .xlsm or .csv file.")

    from openpyxl import load_workbook  # imported lazily: keeps startup fast

    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook.active
        if sheet.max_column < 2:
            raise ValueError("The sheet must have at least two columns: filename and download URL.")
        for name_cell, link_cell in sheet.iter_rows(min_row=2, max_col=2):
            link = cell_text(link_cell.value)
            if link_cell.hyperlink is not None and not parse_links(link):
                link = cell_text(link_cell.hyperlink.target) or link
            rows.append((cell_text(name_cell.value), link))
    finally:
        workbook.close()
    return _trim_trailing_blanks(rows)


def _trim_trailing_blanks(rows):
    while rows and not any(rows[-1]):
        rows.pop()
    return rows


# ── Jobs and results ───────────────────────────────────────────────

@dataclass(frozen=True)
class Job:
    row: int            # spreadsheet row number (1-based, header = row 1)
    url: str
    name: str           # requested filename ('' = use server name)


def make_record(row, url, status, filename="", error="", byte_count=0) -> dict:
    return {"Timestamp": datetime.now().isoformat(timespec="seconds"), "Row": row,
            "Source URL": url, "Status": status, "Saved filename": filename,
            "Error message": error, "Bytes downloaded": byte_count}


def plan_jobs(rows: Iterable[tuple[str, str]]) -> tuple[list[Job], list[dict]]:
    """Turn spreadsheet rows into download jobs.

    Returns ``(jobs, records)`` where *records* already holds the results for
    rows that need no download (missing link, duplicate URL).
    """
    jobs: list[Job] = []
    records: list[dict] = []
    seen: set[str] = set()
    for row_number, (name, link_cell) in enumerate(rows, start=2):
        if not name and not link_cell:
            continue
        links = parse_links(link_cell)
        if not links:
            reason = "URL cell is blank" if not link_cell else f"No valid URL in cell: {link_cell[:80]}"
            records.append(make_record(row_number, link_cell, MISSING, error=reason))
            continue
        for link in links:
            key = download_url_for(link)[0]
            if key in seen:
                records.append(make_record(row_number, link, SKIPPED, error="Duplicate URL skipped"))
                continue
            seen.add(key)
            jobs.append(Job(row_number, link, name))
    return jobs, records


def write_log(path: Path, records: Iterable[dict]) -> Path:
    with Path(path).open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=LOG_FIELDS)
        writer.writeheader()
        writer.writerows(records)
    return Path(path)


# ── Download engine ────────────────────────────────────────────────

class DownloadCancelled(Exception):
    pass


@dataclass
class DownloadEngine:
    """Download jobs concurrently into *output_dir*.

    ``on_event(kind, payload)`` is called from worker threads with:
      * ``("started", job)`` when a worker picks up a job
      * ``("bytes", n)`` as data arrives (for live speed)
      * ``("result", record)`` when a job finishes
    """

    output_dir: Path
    on_event: Callable[[str, object], None] = lambda kind, payload: None
    workers: int = DEFAULT_WORKERS
    trust_env: bool = True
    cancel_event: threading.Event = field(default_factory=threading.Event)

    def __post_init__(self):
        self.output_dir = Path(self.output_dir)
        self._name_lock = threading.Lock()
        self._reserved: set[str] = set()
        self._local = threading.local()

    # Public API
    def cancel(self):
        self.cancel_event.set()

    def run(self, jobs: list[Job]) -> list[dict]:
        if not jobs:
            return []
        with ThreadPoolExecutor(max_workers=max(1, min(self.workers, len(jobs))),
                                thread_name_prefix="gffd-download") as pool:
            return list(pool.map(self._run_job, jobs))

    # Internals
    def _session(self) -> requests.Session:
        session = getattr(self._local, "session", None)
        if session is None:
            session = requests.Session()
            session.trust_env = self.trust_env
            session.headers["User-Agent"] = USER_AGENT
            retry = Retry(total=3, connect=3, read=2, backoff_factor=1.0,
                          status_forcelist=(429, 500, 502, 503, 504),
                          allowed_methods=frozenset({"GET"}), raise_on_status=False)
            adapter = HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4)
            session.mount("https://", adapter)
            session.mount("http://", adapter)
            self._local.session = session
        return session

    def _reserve(self, filename: str) -> Path:
        filename = safe_filename(filename)
        stem, suffix = Path(filename).stem, Path(filename).suffix
        with self._name_lock:
            candidate, number = filename, 2
            while candidate.lower() in self._reserved or (self.output_dir / candidate).exists():
                candidate = f"{stem}_{number}{suffix}"
                number += 1
            self._reserved.add(candidate.lower())
        return self.output_dir / candidate

    def _release(self, target: Path):
        with self._name_lock:
            self._reserved.discard(target.name.lower())

    def _run_job(self, job: Job) -> dict:
        if self.cancel_event.is_set():
            record = make_record(job.row, job.url, CANCELLED, error="Cancelled before start")
            self.on_event("result", record)
            return record
        self.on_event("started", job)
        record = self._download(job)
        self.on_event("result", record)
        return record

    def _download(self, job: Job) -> dict:
        url, drive_id = download_url_for(job.url)
        temporary = target = None
        try:
            with self._session().get(url, stream=True, timeout=TIMEOUT,
                                     allow_redirects=True) as response:
                response.raise_for_status()
                content_type = response.headers.get("content-type", "").split(";")[0].strip().lower()
                if drive_id and content_type == "text/html":
                    raise RuntimeError("Google Drive returned a web page instead of the file. "
                                       "Make sure the file is shared as 'Anyone with the link'.")

                server_name = filename_from_headers(response.headers, response.url)
                filename = job.name or server_name
                if not Path(filename).suffix:
                    filename += (Path(server_name).suffix
                                 or mimetypes.guess_extension(content_type) or "")
                target = self._reserve(filename)
                temporary = self.output_dir / f".{uuid.uuid4().hex}.part"

                byte_count, pending, last_report = 0, 0, time.monotonic()
                with temporary.open("wb") as output:
                    for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                        if self.cancel_event.is_set():
                            raise DownloadCancelled()
                        if chunk:
                            output.write(chunk)
                            byte_count += len(chunk)
                            pending += len(chunk)
                            now = time.monotonic()
                            if now - last_report >= 0.25:
                                self.on_event("bytes", pending)
                                pending, last_report = 0, now
                if pending:
                    self.on_event("bytes", pending)
                if byte_count == 0:
                    raise RuntimeError("The server returned an empty file.")
                os.replace(temporary, target)
                return make_record(job.row, job.url, SUCCESS, target.name, byte_count=byte_count)
        except DownloadCancelled:
            self._cleanup(temporary, target)
            return make_record(job.row, job.url, CANCELLED, error="Cancelled by user")
        except Exception as error:  # noqa: BLE001 – every failure is reported in the log
            self._cleanup(temporary, target)
            message = str(error) or error.__class__.__name__
            if drive_id and "Google Drive" not in message:
                message = f"Google Drive download failed: {message}"
            return make_record(job.row, job.url, FAILED, error=message)

    def _cleanup(self, temporary: Path | None, target: Path | None):
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
        if target is not None:
            self._release(target)
