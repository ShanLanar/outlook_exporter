"""Tests für die GUI-unabhängige Logik des Outlook Advanced Exporters.

Ausführen:  pytest    (benötigt weder Windows noch Outlook)
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import outlook_advanced_exporter as ox  # noqa: E402


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
