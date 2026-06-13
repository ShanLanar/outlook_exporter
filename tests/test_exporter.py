"""Tests für die GUI-unabhängige Logik des Outlook Advanced Exporters.

Ausführen:  pytest    (benötigt weder Windows noch Outlook)
"""
from datetime import datetime

import pytest

import outlook_advanced_exporter as ox


# ── safe_filename ─────────────────────────────────────────────────────────────
@pytest.mark.parametrize("raw, expected", [
    ("Rechnung 2024.pdf", "Rechnung 2024.pdf"),
    ("a/b\\c:d*e?.txt", "abcde.txt"),
    ("Bericht (final).docx", "Bericht final.docx"),
    ("", "unbenannt"),
    ("///", "unbenannt"),
])
def test_safe_filename(raw, expected):
    assert ox.safe_filename(raw) == expected


def test_safe_filename_handles_non_str():
    assert ox.safe_filename(12345) == "12345"


# ── filter_matches ────────────────────────────────────────────────────────────
@pytest.mark.parametrize("needle, haystack, expected", [
    ("", "egal", True),                       # leerer Filter passt immer
    ("@firma.de", "max@firma.de", True),       # Teilstring
    ("RECHNUNG", "Monatsrechnung", True),      # case-insensitiv
    ("xyz", "abc", False),
    ("123", 12345, True),                       # nicht-String haystack
])
def test_filter_matches(needle, haystack, expected):
    assert ox.filter_matches(needle, haystack) is expected


# ── Datums-Filter ─────────────────────────────────────────────────────────────
def test_parse_date_formats():
    assert ox.parse_date("15.03.2024") == datetime(2024, 3, 15)
    assert ox.parse_date("2024-03-15") == datetime(2024, 3, 15)


def test_parse_date_end_of_day():
    assert ox.parse_date("15.03.2024", end_of_day=True) == datetime(2024, 3, 15, 23, 59, 59)


def test_parse_date_invalid_and_empty():
    assert ox.parse_date("") is None
    assert ox.parse_date("kein datum") is None
    assert ox.parse_date("32.13.2024") is None


def test_date_in_range():
    start, end = datetime(2024, 1, 1), datetime(2024, 12, 31)
    assert ox.date_in_range(datetime(2024, 6, 1), start, end) is True
    assert ox.date_in_range(datetime(2023, 6, 1), start, end) is False
    assert ox.date_in_range(datetime(2025, 6, 1), start, end) is False
    assert ox.date_in_range(datetime(2030, 1, 1), None, None) is True


# ── Endungs-Filter ────────────────────────────────────────────────────────────
def test_parse_extensions():
    assert ox.parse_extensions("pdf, .XLSX ,, docx") == {"pdf", "xlsx", "docx"}
    assert ox.parse_extensions("") == set()


@pytest.mark.parametrize("name, allowed, expected", [
    ("rechnung.pdf", {"pdf"}, True),
    ("rechnung.PDF", {"pdf"}, True),
    ("bild.png", {"pdf", "xlsx"}, False),
    ("ohneendung", {"pdf"}, False),
    ("egal.xyz", set(), True),          # kein Filter ⇒ alles erlaubt
])
def test_extension_allowed(name, allowed, expected):
    assert ox.extension_allowed(name, allowed) is expected


# ── DASL-Restrict-Filter ──────────────────────────────────────────────────────
def test_build_restrict_filter_empty():
    assert ox.build_restrict_filter("", "", None, None) == ""


def test_build_restrict_filter_combines():
    flt = ox.build_restrict_filter("@firma.de", "Rechnung", datetime(2024, 1, 1), None)
    assert flt.startswith("@SQL=")
    assert "urn:schemas:httpmail:fromemail" in flt
    assert "%@firma.de%" in flt
    assert "urn:schemas:httpmail:subject" in flt
    assert "datereceived" in flt
    assert " AND " in flt


def test_build_restrict_filter_escapes_quotes():
    flt = ox.build_restrict_filter("o'brien", "", None, None)
    assert "o''brien" in flt


# ── has_category ──────────────────────────────────────────────────────────────
@pytest.mark.parametrize("categories, name, expected", [
    ("Exportiert", "Exportiert", True),
    ("Wichtig, Exportiert, Rot", "exportiert", True),   # case-insensitiv
    ("Wichtig, Rot", "Exportiert", False),
    ("", "Exportiert", False),
    (None, "Exportiert", False),
    ("Exportiert", "", False),
])
def test_has_category(categories, name, expected):
    assert ox.has_category(categories, name) is expected


# ── category_filter_matches ───────────────────────────────────────────────────
@pytest.mark.parametrize("categories, wanted, expected", [
    ("Rechnung", "", True),                       # kein Filter ⇒ alles
    ("Rechnung, Wichtig", "Rechnung", True),
    ("Rechnung", "rechnung", True),                # case-insensitiv
    ("Wichtig", "Rechnung", False),
    ("Steuer", "Rechnung, Steuer", True),          # ODER-Liste
    ("", "Rechnung", False),
    (None, "Rechnung", False),
])
def test_category_filter_matches(categories, wanted, expected):
    assert ox.category_filter_matches(categories, wanted) is expected


# ── file_hash ─────────────────────────────────────────────────────────────────
def test_file_hash_identical_content(tmp_path):
    a = tmp_path / "a.bin"
    b = tmp_path / "b.bin"
    a.write_bytes(b"hallo welt" * 100)
    b.write_bytes(b"hallo welt" * 100)
    assert ox.file_hash(a) == ox.file_hash(b) != ""


def test_file_hash_differs(tmp_path):
    a = tmp_path / "a.bin"
    b = tmp_path / "b.bin"
    a.write_bytes(b"inhalt A")
    b.write_bytes(b"inhalt B")
    assert ox.file_hash(a) != ox.file_hash(b)


def test_file_hash_missing_file(tmp_path):
    assert ox.file_hash(tmp_path / "gibtsnicht.bin") == ""


# ── write_export_csv ──────────────────────────────────────────────────────────
def test_write_export_csv(tmp_path):
    records = [
        ox.ExportRecord("2024-01-01 10:00", "a@b.de", "Test", "Anhang",
                        "datei.pdf", "Gespeichert"),
        ox.ExportRecord("2024-01-02 11:00", "c@d.de", "Zweite", "Body",
                        "mail.txt", "Übersprungen", "ältere Version"),
    ]
    csv_path = ox.write_export_csv(records, tmp_path)
    assert csv_path.exists()
    text = csv_path.read_text(encoding="utf-8-sig")
    assert "Empfangen;Absender;Betreff" in text
    assert "datei.pdf;Gespeichert" in text
    assert "ältere Version" in text


# ── find_folder_by_path ───────────────────────────────────────────────────────
class FakeFolder:
    """Minimaler Ersatz für ein Outlook-Ordner-COM-Objekt."""
    def __init__(self, name, children=None):
        self.Name = name
        self.Folders = children or []


def _tree():
    rechnungen = FakeFolder("Rechnungen")
    projekt = FakeFolder("Projekt", [rechnungen])
    inbox = FakeFolder("Posteingang", [projekt])
    return FakeFolder("Postfach", [inbox]), rechnungen


def test_find_folder_by_path_nested():
    root, rechnungen = _tree()
    found = ox.find_folder_by_path(root, "Postfach\\Posteingang\\Projekt\\Rechnungen")
    assert found is rechnungen


def test_find_folder_by_path_root():
    root, _ = _tree()
    assert ox.find_folder_by_path(root, "Postfach") is root


def test_find_folder_by_path_missing():
    root, _ = _tree()
    assert ox.find_folder_by_path(root, "Postfach\\GibtsNicht") is None


# ── RunLock ───────────────────────────────────────────────────────────────────
def test_runlock_acquire_release(tmp_path):
    lock = ox.RunLock(tmp_path)
    assert lock.acquire() is True
    assert lock.path.exists()
    lock.release()
    assert not lock.path.exists()


def test_runlock_blocks_when_running(tmp_path):
    first = ox.RunLock(tmp_path)
    with first:                      # hält die Sperre (eigene, laufende PID)
        second = ox.RunLock(tmp_path)
        assert second.acquire() is False
    # nach Freigabe wieder möglich
    third = ox.RunLock(tmp_path)
    assert third.acquire() is True
    third.release()


def test_runlock_overwrites_stale(tmp_path):
    # Lock-Datei mit nicht existierender PID = verwaist -> überschreibbar
    (tmp_path / "outlook_exporter.lock").write_text("999999\n2020-01-01\n", encoding="utf-8")
    lock = ox.RunLock(tmp_path)
    assert lock.acquire() is True
    lock.release()


# ── config_from_settings ──────────────────────────────────────────────────────
def test_config_from_settings_maps_folder():
    cfg = ox.config_from_settings({
        "folder": "Postfach\\Posteingang",
        "output_dir": "/tmp/out",
        "export_type": "msg",
        "skip_marked": True,
        "unbekannt": "ignoriert",   # unbekannte Keys werden verworfen
    })
    assert cfg.folder_path == "Postfach\\Posteingang"
    assert cfg.output_dir == "/tmp/out"
    assert cfg.export_type == "msg"
    assert cfg.skip_marked is True


def test_config_from_settings_empty():
    cfg = ox.config_from_settings({})
    assert cfg.folder_path == ""
    assert cfg.export_type == "attachments"


# ── ExportConfig ──────────────────────────────────────────────────────────────
def test_export_config_fields():
    cfg = ox.ExportConfig(
        folder_path="Postfach\\Posteingang",
        include_subfolders=True,
        filter_sender="@firma.de",
        filter_subject="Rechnung",
        filter_body="",
        export_type="attachments",
        output_dir="/tmp/out",
    )
    assert cfg.export_type == "attachments"
    assert cfg.include_subfolders is True
