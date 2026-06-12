#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Advanced Exporter
Exportiert E-Mails und/oder Anhänge aus Outlook 2010+
- Ordner-Dropdown mit vollständigen Pfaden
- Unterordner einschließen (optional)
- Multi-Filter: Absender, Betreff, Text
- Intelligente Konfliktauflösung (neuere Anhänge gewinnen)

Architektur (angelehnt an die Geschwister-Tools pdftool / bilddownloader):
- Klassenbasierte GUI (``App(tk.Tk)``) statt globaler Variablen
- Export läuft in einem ``ExportWorker``-Thread mit Abbruch-Möglichkeit
- Thread-sicheres Logging über eine ``queue.Queue`` (Tkinter ist NICHT
  thread-safe – Worker schreibt nie direkt ins Widget)
- Einstellungen werden als JSON gemerkt
"""

import csv
import hashlib
import json
import logging
import queue
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    import win32com.client
except ImportError:  # nur unter Windows mit Outlook verfügbar
    win32com = None

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ──────────────────────────────────────────────────────────────────────────────
# Logging – alle Records laufen über die Queue und werden im Main-Thread
# ins Text-Widget geschrieben (siehe QueueLogHandler / App._drain_queue).
# ──────────────────────────────────────────────────────────────────────────────
log = logging.getLogger("outlook_exporter")
log.setLevel(logging.INFO)


class QueueLogHandler(logging.Handler):
    """Leitet Log-Records thread-sicher in die GUI-Queue um."""

    def __init__(self, msg_queue: "queue.Queue"):
        super().__init__()
        self.q = msg_queue

    def emit(self, record: logging.LogRecord) -> None:
        self.q.put(("log", (self.format(record), record.levelname)))


def app_base_dir() -> Path:
    """Basisverzeichnis: Ordner der ``.exe`` (PyInstaller) bzw. des Skripts."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


SETTINGS_PATH = app_base_dir() / "outlook_exporter_settings.json"


# ──────────────────────────────────────────────────────────────────────────────
# Datenstruktur für einen Exportlauf
# ──────────────────────────────────────────────────────────────────────────────
@dataclass
class ExportConfig:
    folder_path: str
    include_subfolders: bool
    filter_sender: str
    filter_subject: str
    filter_body: str
    export_type: str  # "attachments" | "body"
    output_dir: str
    filter_date_from: str = ""     # "TT.MM.JJJJ" oder "JJJJ-MM-TT"
    filter_date_to: str = ""
    filter_extensions: str = ""    # z.B. "pdf, xlsx" (leer = alle)
    skip_inline_images: bool = True


@dataclass
class ExportRecord:
    """Eine Zeile des Exportprotokolls (siehe CSV-Export)."""
    received: str
    sender: str
    subject: str
    kind: str       # "Anhang" | "Body"
    filename: str
    status: str     # "Gespeichert" | "Überschrieben" | "Übersprungen"
    note: str = ""


# ──────────────────────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────────────────────────────────────
def safe_filename(name: str) -> str:
    """Erzeugt einen sicheren Dateinamen (alphanumerisch + ``  ._-``)."""
    cleaned = "".join(c for c in str(name) if c.isalnum() or c in " ._-").strip()
    return cleaned or "unbenannt"


def filter_matches(needle: str, haystack) -> bool:
    """True, wenn der Filter leer ist oder als Teilstring (case-insensitiv) passt."""
    return not needle or needle.lower() in str(haystack).lower()


def parse_date(text: str, end_of_day: bool = False):
    """Parst ``TT.MM.JJJJ`` oder ``JJJJ-MM-TT`` zu ``datetime``. Leerer/ungültiger
    Text ergibt ``None``. Bei ``end_of_day`` zählt der ganze Tag bis 23:59:59."""
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.replace(hour=23, minute=59, second=59) if end_of_day else dt
        except ValueError:
            continue
    return None


def date_in_range(dt: datetime, start, end) -> bool:
    """True, wenn ``dt`` innerhalb [start, end] liegt (Grenzen optional)."""
    if start and dt < start:
        return False
    if end and dt > end:
        return False
    return True


def parse_extensions(raw: str) -> set:
    """``"pdf, .xlsx"`` → ``{"pdf", "xlsx"}`` (klein, ohne Punkt)."""
    return {e.strip().lower().lstrip(".") for e in (raw or "").split(",") if e.strip()}


def extension_allowed(filename: str, allowed: set) -> bool:
    """True, wenn keine Endungen vorgegeben sind oder die Datei passt."""
    if not allowed:
        return True
    if "." not in filename:
        return False
    return filename.rsplit(".", 1)[-1].lower() in allowed


def build_restrict_filter(sender: str, subject: str, dt_from, dt_to) -> str:
    """Baut einen Outlook-DASL-``@SQL``-Filter (oder "" wenn nichts zu filtern).

    Damit lassen sich Absender/Betreff/Datum serverseitig vorfiltern, statt jede
    Mail einzeln zu prüfen. Die clientseitigen Checks bleiben als Sicherheitsnetz.
    """
    def esc(s: str) -> str:
        return s.replace("'", "''")

    clauses = []
    if sender:
        clauses.append(f"\"urn:schemas:httpmail:fromemail\" LIKE '%{esc(sender)}%'")
    if subject:
        clauses.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{esc(subject)}%'")
    if dt_from:
        clauses.append(f"\"urn:schemas:httpmail:datereceived\" >= '{dt_from:%m/%d/%Y %H:%M}'")
    if dt_to:
        clauses.append(f"\"urn:schemas:httpmail:datereceived\" <= '{dt_to:%m/%d/%Y %H:%M}'")
    return "@SQL=" + " AND ".join(clauses) if clauses else ""


# MAPI-Property: Content-ID → kennzeichnet inline eingebettete Bilder (Signaturen)
_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"


def is_inline_attachment(attachment) -> bool:
    """True, wenn der Anhang eine Content-ID hat (inline im HTML, z.B. Signaturlogo)."""
    try:
        return bool(attachment.PropertyAccessor.GetProperty(_PR_ATTACH_CONTENT_ID))
    except Exception:  # noqa: BLE001 – Property fehlt ⇒ kein Inline-Bild
        return False


def file_hash(path: Path, chunk: int = 65536) -> str:
    """SHA-256-Hash einer Datei (erste 2 MB genügen für Duplikaterkennung)."""
    h = hashlib.sha256()
    read = 0
    max_bytes = 2 * 1024 * 1024
    try:
        with open(path, "rb") as f:
            while read < max_bytes:
                buf = f.read(min(chunk, max_bytes - read))
                if not buf:
                    break
                h.update(buf)
                read += len(buf)
    except OSError:
        return ""
    return h.hexdigest()


def write_export_csv(records: list, out_dir: Path) -> Path:
    """Schreibt das Exportprotokoll als CSV (``;`` getrennt, Excel-tauglich)."""
    csv_path = out_dir / f"_export_protokoll_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Empfangen", "Absender", "Betreff", "Typ",
                         "Dateiname", "Status", "Hinweis"])
        for r in records:
            writer.writerow([r.received, r.sender, r.subject, r.kind,
                             r.filename, r.status, r.note])
    return csv_path


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(data: dict) -> None:
    try:
        SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError as e:
        log.warning("Einstellungen konnten nicht gespeichert werden: %s", e)


# ──────────────────────────────────────────────────────────────────────────────
# Outlook-Zugriff (GUI-unabhängig)
# ──────────────────────────────────────────────────────────────────────────────
def get_all_outlook_folders() -> list:
    """
    Sammelt alle Ordner inkl. Unterordner als Liste vollständiger Pfade
    (z.B. ``"Posteingang\\Projekt\\Rechnungen"``). Wirft bei Fehler eine
    Exception – die GUI entscheidet, wie sie darauf reagiert.
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folders = []

    def traverse(folder, parent_path=""):
        current_path = f"{parent_path}\\{folder.Name}" if parent_path else folder.Name
        folders.append(current_path)
        for subfolder in folder.Folders:
            traverse(subfolder, current_path)

    for store in outlook.Folders:  # Postfach, Archiv usw.
        traverse(store)
    return folders


def find_folder_by_path(folder, path):
    """Durchsucht rekursiv nach einem Pfad (z.B. ``"Posteingang\\Projekt"``)."""
    if folder.Name == path.split("\\")[0]:
        if "\\" not in path:
            return folder
        remaining = "\\".join(path.split("\\")[1:])
        for sub in folder.Folders:
            found = find_folder_by_path(sub, remaining)
            if found:
                return found
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Export-Worker (läuft im eigenen Thread)
# ──────────────────────────────────────────────────────────────────────────────
class ExportWorker(threading.Thread):
    """Führt den Export aus und meldet Fortschritt über die Queue."""

    def __init__(self, config: ExportConfig, msg_queue: "queue.Queue",
                 stop_event: threading.Event):
        super().__init__(daemon=True)
        self.config = config
        self.q = msg_queue
        self.stop_event = stop_event
        self.records: list = []
        # Filter einmalig vorparsen
        self._dt_from = parse_date(config.filter_date_from)
        self._dt_to = parse_date(config.filter_date_to, end_of_day=True)
        self._exts = parse_extensions(config.filter_extensions)

    def run(self) -> None:
        try:
            count = self._export()
            if self.stop_event.is_set():
                self.q.put(("cancelled", count))
            else:
                self.q.put(("done", count))
        except Exception as e:  # noqa: BLE001 – an die GUI weiterreichen
            self.q.put(("error", str(e)))

    # -- intern ----------------------------------------------------------------
    def _export(self) -> int:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        log.info("✅ Mit Outlook verbunden.")

        target_folder = None
        for store in outlook.Folders:
            target_folder = find_folder_by_path(store, self.config.folder_path)
            if target_folder:
                break
        if not target_folder:
            raise RuntimeError(f"Ordner '{self.config.folder_path}' nicht gefunden.")

        log.info("📁 Verarbeite Ordner: %s (inkl. Unterordner: %s)",
                 self.config.folder_path, self.config.include_subfolders)

        out_dir = Path(self.config.output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        count = self._process_folder(target_folder, self.config.include_subfolders, out_dir)

        if self.records:
            csv_path = write_export_csv(self.records, out_dir)
            log.info("🧾 Protokoll geschrieben: %s", csv_path.name)
        return count

    def _process_folder(self, folder, include_subfolders: bool, out_dir: Path) -> int:
        count = 0
        cfg = self.config
        try:
            items = folder.Items
            # Absender/Betreff/Datum serverseitig vorfiltern (Performance)
            restrict = build_restrict_filter(cfg.filter_sender, cfg.filter_subject,
                                             self._dt_from, self._dt_to)
            if restrict:
                try:
                    items = items.Restrict(restrict)
                except Exception as e:  # noqa: BLE001 – Fallback auf volle Suche
                    log.warning("Restrict-Filter nicht angewendet (%s) – volle Suche.", e)
            items.Sort("[ReceivedTime]", True)

            for mail in items:
                if self.stop_event.is_set():
                    return count
                if mail.Class != 43:  # nur E-Mails
                    continue

                # Clientseitige Filter (Sicherheitsnetz, auch ohne Restrict korrekt)
                if not filter_matches(cfg.filter_sender, mail.SenderEmailAddress):
                    continue
                if not filter_matches(cfg.filter_subject, mail.Subject):
                    continue
                if not filter_matches(cfg.filter_body, mail.Body):
                    continue
                if self._dt_from or self._dt_to:
                    rt = datetime.fromtimestamp(mail.ReceivedTime.timestamp())
                    if not date_in_range(rt, self._dt_from, self._dt_to):
                        continue

                log.info("📧 %s - %s",
                         mail.ReceivedTime.strftime("%Y-%m-%d %H:%M"), mail.Subject)

                if cfg.export_type == "attachments":
                    count += self._save_attachments(mail, out_dir)
                elif cfg.export_type == "body":
                    count += self._save_body(mail, out_dir)

            if include_subfolders:
                for sub in folder.Folders:
                    if self.stop_event.is_set():
                        break
                    count += self._process_folder(sub, True, out_dir)
            return count
        except Exception as e:  # noqa: BLE001
            log.error("Fehler in %s: %s", folder.Name, e)
            return count

    def _record(self, mail, kind, filename, status, note=""):
        self.records.append(ExportRecord(
            received=mail.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
            sender=str(mail.SenderEmailAddress),
            subject=str(mail.Subject),
            kind=kind, filename=filename, status=status, note=note,
        ))

    def _save_attachments(self, mail, out_dir: Path) -> int:
        saved = 0
        cfg = self.config
        for attachment in mail.Attachments:
            # Signatur-/Inline-Bilder und unerwünschte Dateitypen still überspringen
            if cfg.skip_inline_images and is_inline_attachment(attachment):
                continue
            if not extension_allowed(str(attachment.FileName), self._exts):
                continue

            safe_name = safe_filename(attachment.FileName)
            target_path = out_dir / safe_name

            if not target_path.exists():
                attachment.SaveAsFile(str(target_path))
                log.info("   ✅ Gespeichert: %s", safe_name)
                self._record(mail, "Anhang", safe_name, "Gespeichert")
                saved += 1
                continue

            # Konflikt: erst per Inhalts-Hash auf echtes Duplikat prüfen,
            # dann (bei abweichendem Inhalt) gewinnt die neuere Version.
            tmp_path = out_dir / (safe_name + ".tmp_export")
            attachment.SaveAsFile(str(tmp_path))
            if file_hash(tmp_path) == file_hash(target_path):
                tmp_path.unlink(missing_ok=True)
                log.info("   ⏭️  Übersprungen (identisch): %s", safe_name)
                self._record(mail, "Anhang", safe_name, "Übersprungen", "identischer Inhalt")
                continue

            existing_mtime = target_path.stat().st_mtime
            try:
                att_mtime = attachment.LastModifiedTime.timestamp()
            except AttributeError:
                att_mtime = mail.ReceivedTime.timestamp()

            if att_mtime <= existing_mtime:
                tmp_path.unlink(missing_ok=True)
                log.info("   ⏭️  Übersprungen (ältere Version): %s", safe_name)
                self._record(mail, "Anhang", safe_name, "Übersprungen", "ältere Version")
                continue

            tmp_path.replace(target_path)
            log.info("   🔄 Überschrieben (neuere Version): %s", safe_name)
            self._record(mail, "Anhang", safe_name, "Überschrieben", "neuere Version")
            saved += 1
        return saved

    def _save_body(self, mail, out_dir: Path) -> int:
        timestamp = mail.ReceivedTime.strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}_{safe_filename(mail.Subject)}.txt"
        filepath = out_dir / filename
        filepath.write_text(
            f"Von: {mail.SenderEmailAddress}\n"
            f"Betreff: {mail.Subject}\n"
            f"Empfangen: {mail.ReceivedTime}\n"
            + "\n" + "=" * 60 + "\n\n"
            + str(mail.Body),
            encoding="utf-8",
        )
        log.info("   📄 Body gespeichert: %s", filename)
        self._record(mail, "Body", filename, "Gespeichert")
        return 1


# ──────────────────────────────────────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook-Exporter Pro (mit Ordnerbaum & Konfliktauflösung)")
        self.geometry("850x750")
        self.minsize(700, 600)

        self.queue: "queue.Queue" = queue.Queue()
        self.worker: ExportWorker | None = None
        self.stop_event = threading.Event()

        # Logging an die Queue hängen (thread-sicher)
        handler = QueueLogHandler(self.queue)
        handler.setFormatter(logging.Formatter("%(message)s"))
        log.handlers.clear()
        log.addHandler(handler)

        ttk.Style().theme_use("clam")
        self.settings = load_settings()
        self._build_ui()
        self._load_folders()
        self._apply_settings()
        self.after(100, self._drain_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # -- UI-Aufbau -------------------------------------------------------------
    def _build_ui(self) -> None:
        mainframe = ttk.Frame(self, padding="10")
        mainframe.pack(fill=tk.BOTH, expand=True)
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(10, weight=1)

        # 1. Ordnerauswahl
        ttk.Label(mainframe, text="Outlook-Ordner:").grid(column=0, row=0, sticky=tk.W, pady=2)
        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(mainframe, textvariable=self.folder_var,
                                          width=80, state="readonly")
        self.folder_combo.grid(column=0, row=1, sticky=(tk.W, tk.E), pady=2)
        ttk.Button(mainframe, text="Ordner neu laden",
                   command=self._load_folders).grid(column=1, row=1, padx=5)

        # 2. Unterordner einschließen
        self.include_sub_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(mainframe, text="Unterordner einschließen",
                        variable=self.include_sub_var).grid(column=0, row=2, sticky=tk.W, pady=5)

        # 3. Filter
        filter_frame = ttk.LabelFrame(mainframe, text="Filter (leer = alle)", padding="5")
        filter_frame.grid(column=0, row=3, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(filter_frame, text="Absender (enthält):").grid(column=0, row=0, sticky=tk.W)
        self.sender_entry = ttk.Entry(filter_frame, width=40)
        self.sender_entry.grid(column=1, row=0, sticky=tk.W, padx=5)
        ttk.Label(filter_frame, text="Betreff (enthält):").grid(column=0, row=1, sticky=tk.W)
        self.subject_entry = ttk.Entry(filter_frame, width=40)
        self.subject_entry.grid(column=1, row=1, sticky=tk.W, padx=5)
        ttk.Label(filter_frame, text="Text (enthält):").grid(column=0, row=2, sticky=tk.W)
        self.body_entry = ttk.Entry(filter_frame, width=40)
        self.body_entry.grid(column=1, row=2, sticky=tk.W, padx=5)

        ttk.Label(filter_frame, text="Datum von (TT.MM.JJJJ):").grid(column=0, row=3, sticky=tk.W)
        self.date_from_entry = ttk.Entry(filter_frame, width=40)
        self.date_from_entry.grid(column=1, row=3, sticky=tk.W, padx=5)
        ttk.Label(filter_frame, text="Datum bis (TT.MM.JJJJ):").grid(column=0, row=4, sticky=tk.W)
        self.date_to_entry = ttk.Entry(filter_frame, width=40)
        self.date_to_entry.grid(column=1, row=4, sticky=tk.W, padx=5)
        ttk.Label(filter_frame, text="Nur Dateitypen (z.B. pdf, xlsx):").grid(
            column=0, row=5, sticky=tk.W)
        self.ext_entry = ttk.Entry(filter_frame, width=40)
        self.ext_entry.grid(column=1, row=5, sticky=tk.W, padx=5)

        self.skip_inline_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(filter_frame, text="Signatur-/Inline-Bilder überspringen",
                        variable=self.skip_inline_var).grid(column=0, row=6, columnspan=2,
                                                             sticky=tk.W, pady=(4, 0))

        # 4. Export-Typ
        ttk.Label(mainframe, text="Export-Typ:").grid(column=0, row=4, sticky=tk.W, pady=5)
        self.export_var = tk.StringVar(value="attachments")
        ttk.Radiobutton(mainframe, text="📎 Nur Anhänge", variable=self.export_var,
                        value="attachments").grid(column=0, row=5, sticky=tk.W)
        ttk.Radiobutton(mainframe, text="📄 Nur E-Mail-Body (als .txt)", variable=self.export_var,
                        value="body").grid(column=0, row=6, sticky=tk.W)

        # 5. Ausgabeverzeichnis
        ttk.Label(mainframe, text="Ausgabeverzeichnis:").grid(column=0, row=7, sticky=tk.W, pady=5)
        output_frame = ttk.Frame(mainframe)
        output_frame.grid(column=0, row=8, sticky=(tk.W, tk.E))
        self.output_entry = ttk.Entry(output_frame, width=70)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Durchsuchen...",
                   command=self._select_output).pack(side=tk.RIGHT, padx=5)

        # 6. Log
        ttk.Label(mainframe, text="Log:").grid(column=0, row=9, sticky=tk.W, pady=5)
        self.log_widget = scrolledtext.ScrolledText(mainframe, height=18, state=tk.NORMAL)
        self.log_widget.grid(column=0, row=10, columnspan=2,
                             sticky=(tk.N, tk.S, tk.E, tk.W), pady=5)
        # Farbliche Kategorien für die Log-Ausgabe
        self.log_widget.tag_config("error", foreground="#c0392b")
        self.log_widget.tag_config("warning", foreground="#d35400")
        self.log_widget.tag_config("success", foreground="#1e8449")
        self.log_widget.tag_config("muted", foreground="#7f8c8d")
        self.progress_label = ttk.Label(mainframe, text="Bereit.")
        self.progress_label.grid(column=0, row=11, sticky=tk.W, pady=5)

        # 7. Buttons
        btn_frame = ttk.Frame(mainframe)
        btn_frame.grid(column=0, row=12, columnspan=2, pady=10)
        self.btn_start = ttk.Button(btn_frame, text="🚀 Export starten", command=self._on_start)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(btn_frame, text="⏹ Abbrechen",
                                   command=self._on_stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)

    # -- Aktionen --------------------------------------------------------------
    def _load_folders(self) -> None:
        try:
            names = get_all_outlook_folders()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Fehler", f"Outlook-Ordner konnten nicht gelesen werden:\n{e}")
            return
        self.folder_combo["values"] = names
        if names:
            current = self.folder_var.get()
            self.folder_combo.current(names.index(current) if current in names else 0)
        log.info("🔄 Ordnerliste geladen (%d Ordner).", len(names))

    def _select_output(self) -> None:
        d = filedialog.askdirectory()
        if d:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, d)

    def _on_start(self) -> None:
        if not self.folder_var.get():
            messagebox.showwarning("Fehler", "Bitte wählen Sie einen Ordner aus.")
            return
        if not self.output_entry.get():
            messagebox.showwarning("Fehler", "Bitte wählen Sie ein Ausgabeverzeichnis.")
            return

        # Datumseingaben validieren (leer ist erlaubt)
        for label, value in (("Datum von", self.date_from_entry.get()),
                             ("Datum bis", self.date_to_entry.get())):
            if value.strip() and parse_date(value) is None:
                messagebox.showwarning(
                    "Fehler", f"'{label}' ist kein gültiges Datum (TT.MM.JJJJ).")
                return

        config = ExportConfig(
            folder_path=self.folder_var.get(),
            include_subfolders=self.include_sub_var.get(),
            filter_sender=self.sender_entry.get(),
            filter_subject=self.subject_entry.get(),
            filter_body=self.body_entry.get(),
            export_type=self.export_var.get(),
            output_dir=self.output_entry.get(),
            filter_date_from=self.date_from_entry.get(),
            filter_date_to=self.date_to_entry.get(),
            filter_extensions=self.ext_entry.get(),
            skip_inline_images=self.skip_inline_var.get(),
        )
        self._save_settings()

        self.log_widget.delete(1.0, tk.END)
        self.progress_label.config(text="Export läuft...")
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)

        self.stop_event.clear()
        self.worker = ExportWorker(config, self.queue, self.stop_event)
        self.worker.start()

    def _on_stop(self) -> None:
        self.stop_event.set()
        self.progress_label.config(text="Abbruch angefordert...")
        self.btn_stop.config(state=tk.DISABLED)

    # -- Queue-Verarbeitung (Main-Thread) -------------------------------------
    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    msg, level = payload
                    self._append_log(msg, level)
                elif kind == "done":
                    self._finish(f"✨ Fertig! {payload} Elemente exportiert.",
                                 "Erfolg", f"Export beendet.\n{payload} Elemente gespeichert.")
                elif kind == "cancelled":
                    self._finish(f"⏹ Abgebrochen. {payload} Elemente exportiert.",
                                 None, None)
                elif kind == "error":
                    self._finish(f"❌ Kritischer Fehler: {payload}",
                                 "Fehler", payload, error=True)
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    _TAG_BY_LEVEL = {"ERROR": "error", "CRITICAL": "error", "WARNING": "warning"}
    _SUCCESS_MARKERS = ("✅", "✨", "🔄", "🧾", "📄")

    def _append_log(self, text: str, level: str = "INFO") -> None:
        tag = self._TAG_BY_LEVEL.get(level, "")
        if not tag:
            stripped = text.lstrip()
            if stripped.startswith(self._SUCCESS_MARKERS):
                tag = "success"
            elif stripped.startswith("⏭"):
                tag = "muted"
        self.log_widget.insert(tk.END, text + "\n", tag or ())
        self.log_widget.see(tk.END)

    def _finish(self, status: str, title, message, error: bool = False) -> None:
        self._append_log("\n" + status)
        self.progress_label.config(text=status)
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.worker = None
        if title and message:
            (messagebox.showerror if error else messagebox.showinfo)(title, message)

    # -- Einstellungen ---------------------------------------------------------
    def _apply_settings(self) -> None:
        s = self.settings
        if s.get("folder") in self.folder_combo["values"]:
            self.folder_var.set(s["folder"])
        self.include_sub_var.set(s.get("include_subfolders", True))
        self.sender_entry.insert(0, s.get("filter_sender", ""))
        self.subject_entry.insert(0, s.get("filter_subject", ""))
        self.body_entry.insert(0, s.get("filter_body", ""))
        self.export_var.set(s.get("export_type", "attachments"))
        self.output_entry.insert(0, s.get("output_dir", ""))
        self.date_from_entry.insert(0, s.get("filter_date_from", ""))
        self.date_to_entry.insert(0, s.get("filter_date_to", ""))
        self.ext_entry.insert(0, s.get("filter_extensions", ""))
        self.skip_inline_var.set(s.get("skip_inline_images", True))

    def _save_settings(self) -> None:
        save_settings({
            "folder": self.folder_var.get(),
            "include_subfolders": self.include_sub_var.get(),
            "filter_sender": self.sender_entry.get(),
            "filter_subject": self.subject_entry.get(),
            "filter_body": self.body_entry.get(),
            "export_type": self.export_var.get(),
            "output_dir": self.output_entry.get(),
            "filter_date_from": self.date_from_entry.get(),
            "filter_date_to": self.date_to_entry.get(),
            "filter_extensions": self.ext_entry.get(),
            "skip_inline_images": self.skip_inline_var.get(),
        })

    def _on_close(self) -> None:
        self.stop_event.set()
        self._save_settings()
        self.destroy()


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
