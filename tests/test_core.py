import csv
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest
from openpyxl import Workbook

from gffd.core import (FAILED, MISSING, SKIPPED, SUCCESS, DownloadEngine, Job, cell_text,
                       download_url_for, extract_file_id, filename_from_headers, parse_links,
                       plan_jobs, read_rows, safe_filename, write_log)

DRIVE_ID = "1AbCdEfGhIjKlMnOpQrStUv"


@pytest.mark.parametrize("link", [
    f"https://drive.google.com/open?id={DRIVE_ID}",
    f"https://drive.google.com/file/d/{DRIVE_ID}/view?usp=sharing",
    f"https://docs.google.com/uc?export=download&id={DRIVE_ID}",
])
def test_extract_file_id_from_drive_links(link):
    assert extract_file_id(link) == DRIVE_ID


def test_extract_file_id_ignores_other_hosts():
    assert extract_file_id(f"https://example.com/d/{DRIVE_ID}/file") is None
    assert download_url_for("https://example.com/a.pdf") == ("https://example.com/a.pdf", None)


def test_parse_links_handles_multiple_uploads_in_one_cell():
    cell = f"https://drive.google.com/open?id={DRIVE_ID}, https://drive.google.com/open?id=ZZZZZZZZZZZZ"
    assert len(parse_links(cell)) == 2
    assert parse_links("not a link") == []
    assert parse_links("") == []


@pytest.mark.parametrize("raw, expected", [
    ("report.pdf", "report.pdf"),
    ('a<b>c:"d"|e?.pdf', "a_b_c__d__e_.pdf"),
    ("../../etc/passwd", "passwd"),
    ("name. ", "name"),
    ("CON.txt", "_CON.txt"),
    ("", "download"),
])
def test_safe_filename(raw, expected):
    assert safe_filename(raw) == expected


def test_safe_filename_keeps_extension_when_truncating():
    name = safe_filename("x" * 300 + ".pdf")
    assert len(name) <= 180 and name.endswith(".pdf")


def test_cell_text_normalises_numbers_and_blanks():
    assert cell_text(None) == ""
    assert cell_text(float("nan")) == ""
    assert cell_text(101.0) == "101"
    assert cell_text("  hi ") == "hi"


def test_filename_from_headers():
    assert filename_from_headers({"content-disposition": "attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.pdf"},
                                 "https://x/y") == "résumé.pdf"
    assert filename_from_headers({"content-disposition": 'attachment; filename="a b.png"'},
                                 "https://x/y") == "a b.png"
    assert filename_from_headers({}, "https://x/files/photo%201.jpg") == "photo 1.jpg"
    assert filename_from_headers({}, "https://x/") == "download"


def test_plan_jobs_reports_missing_and_duplicates():
    link = f"https://drive.google.com/open?id={DRIVE_ID}"
    same_file = f"https://drive.google.com/file/d/{DRIVE_ID}/view"
    jobs, records = plan_jobs([("a", link), ("b", ""), ("c", same_file), ("", ""), ("d", "n/a")])
    assert [j.name for j in jobs] == ["a"]
    assert [(r["Row"], r["Status"]) for r in records] == [(3, MISSING), (4, SKIPPED), (6, MISSING)]


def test_read_rows_xlsx_and_csv(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["File name", "Link"])
    ws.append([1, "https://example.com/a.pdf"])
    ws.append(["Linked", "Open file"])
    ws["B3"].hyperlink = "https://example.com/b.pdf"
    ws.append([None, None])
    xlsx = tmp_path / "in.xlsx"
    wb.save(xlsx)
    assert read_rows(xlsx) == [("1", "https://example.com/a.pdf"), ("Linked", "https://example.com/b.pdf")]

    csv_path = tmp_path / "in.csv"
    csv_path.write_text("name,link\nx,https://example.com/x\n", encoding="utf-8")
    assert read_rows(csv_path) == [("x", "https://example.com/x")]

    with pytest.raises(ValueError):
        read_rows(tmp_path / "in.txt")


@pytest.fixture()
def http_server(tmp_path):
    root = tmp_path / "www"
    root.mkdir()
    (root / "doc.pdf").write_bytes(b"%PDF-1.4 test" * 100)
    (root / "image.png").write_bytes(b"\x89PNG" + b"0" * 500)
    (root / "empty.bin").write_bytes(b"")
    handler = partial(SimpleHTTPRequestHandler, directory=str(root))
    handler.log_message = lambda *a, **k: None
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{server.server_address[1]}"
    server.shutdown()


def test_engine_downloads_renames_and_logs(tmp_path, http_server):
    out = tmp_path / "out"
    out.mkdir()
    (out / "Student_1.pdf").write_bytes(b"existing")
    jobs = [Job(2, f"{http_server}/doc.pdf", "Student_1"),
            Job(3, f"{http_server}/image.png", ""),
            Job(4, f"{http_server}/missing.pdf", "x"),
            Job(5, f"{http_server}/empty.bin", "y")]
    events = []
    results = DownloadEngine(out, on_event=lambda *e: events.append(e), trust_env=False).run(jobs)

    by_row = {r["Row"]: r for r in results}
    assert by_row[2]["Status"] == SUCCESS and by_row[2]["Saved filename"] == "Student_1_2.pdf"
    assert by_row[3]["Status"] == SUCCESS and by_row[3]["Saved filename"] == "image.png"
    assert by_row[4]["Status"] == FAILED and "404" in by_row[4]["Error message"]
    assert by_row[5]["Status"] == FAILED and "empty" in by_row[5]["Error message"]
    assert (out / "Student_1.pdf").read_bytes() == b"existing"   # never overwritten
    assert not list(out.glob("*.part")) and not list(out.glob(".*.part"))
    assert any(kind == "bytes" for kind, _ in events)

    log = write_log(out / "log.csv", results)
    with log.open(encoding="utf-8-sig") as handle:
        assert len(list(csv.DictReader(handle))) == 4


def test_engine_cancel_before_start(tmp_path, http_server):
    out = tmp_path / "out"
    out.mkdir()
    engine = DownloadEngine(out, trust_env=False)
    engine.cancel()
    results = engine.run([Job(2, f"{http_server}/doc.pdf", "a")])
    assert results[0]["Status"] == "Cancelled"
    assert not list(Path(out).iterdir())
