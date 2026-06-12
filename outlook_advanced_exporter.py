#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Advanced Exporter
Exportiert E-Mails und/oder Anhänge aus Outlook 2010+
- Ordner-Dropdown mit vollständigen Pfaden
- Unterordner einschließen (optional)
- Multi-Filter: Absender, Betreff, Text
- Intelligente Konfliktauflösung (neuere Anhänge gewinnen)
"""

import os
import win32com.client
from tkinter import *
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
from datetime import datetime


# ---------- Outlook Ordner rekursiv sammeln (für Dropdown) ----------
def get_all_outlook_folders():
    """
    Sammelt alle Ordner inkl. Unterordner und gibt eine Liste von (Anzeigename, Ordner-Objekt) zurück.
    Der Anzeigename ist der vollständige Pfad (z.B. "Posteingang\\Projekt\\Rechnungen").
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folders = []

        def traverse(folder, parent_path=""):
            current_path = f"{parent_path}\\{folder.Name}" if parent_path else folder.Name
            folders.append((current_path, folder))
            for subfolder in folder.Folders:
                traverse(subfolder, current_path)

        for store in outlook.Folders:  # Postfach, Archiv usw.
            traverse(store)
        return folders
    except Exception as e:
        messagebox.showerror("Fehler", f"Outlook-Ordner konnten nicht gelesen werden:\n{e}")
        return []


# ---------- Rekursive Suche in Unterordnern ----------
def process_folder(folder, include_subfolders, filter_sender, filter_subject, filter_body, export_type, output_dir, log_widget):
    """
    Durchläuft einen Ordner (und optional Unterordner) und exportiert nach Vorgabe.
    Gibt die Anzahl der verarbeiteten E-Mails zurück.
    """
    count = 0
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        for mail in items:
            if mail.Class != 43:  # nur E-Mails
                continue

            # Filter anwenden
            if filter_sender and filter_sender.lower() not in str(mail.SenderEmailAddress).lower():
                continue
            if filter_subject and filter_subject.lower() not in str(mail.Subject).lower():
                continue
            if filter_body and filter_body.lower() not in str(mail.Body).lower():
                continue

            log_widget.insert(END, f"\n📧 {mail.ReceivedTime.strftime('%Y-%m-%d %H:%M')} - {mail.Subject}\n")
            log_widget.see(END)

            if export_type == "attachments":
                for attachment in mail.Attachments:
                    # Sicherer Dateiname
                    safe_name = "".join(c for c in attachment.FileName if c.isalnum() or c in " ._-")
                    target_path = os.path.join(output_dir, safe_name)

                    # Konfliktauflösung: Wenn Datei existiert, prüfe, welcher Anhang neuer ist
                    if os.path.exists(target_path):
                        # Hole letzte Änderungszeit der existierenden Datei
                        existing_mtime = os.path.getmtime(target_path)
                        # Versuche, das Änderungsdatum des Anhangs aus Outlook zu bekommen
                        try:
                            # Manche Attachment-Objekte haben LastModifiedTime
                            att_mtime = attachment.LastModifiedTime.timestamp()
                        except AttributeError:
                            # Fallback: Empfangsdatum der E-Mail
                            att_mtime = mail.ReceivedTime.timestamp()

                        if att_mtime <= existing_mtime:
                            log_widget.insert(END, f"   ⏭️  Übersprungen (ältere Version): {safe_name}\n")
                            continue
                        else:
                            log_widget.insert(END, f"   🔄 Überschrieben (neuere Version): {safe_name}\n")
                    # Speichern
                    attachment.SaveAsFile(target_path)
                    log_widget.insert(END, f"   ✅ Gespeichert: {safe_name}\n")
                    log_widget.see(END)
                    count += 1

            elif export_type == "body":
                # Body in Textdatei speichern
                timestamp = mail.ReceivedTime.strftime('%Y%m%d_%H%M%S')
                safe_subject = "".join(c for c in mail.Subject if c.isalnum() or c in " ._-")
                filename = f"{timestamp}_{safe_subject}.txt"
                filepath = os.path.join(output_dir, filename)
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(f"Von: {mail.SenderEmailAddress}\n")
                    f.write(f"Betreff: {mail.Subject}\n")
                    f.write(f"Empfangen: {mail.ReceivedTime}\n")
                    f.write("\n" + "="*60 + "\n\n")
                    f.write(mail.Body)
                log_widget.insert(END, f"   📄 Body gespeichert: {filename}\n")
                log_widget.see(END)
                count += 1

        # Unterordner verarbeiten, falls gewünscht
        if include_subfolders:
            for sub in folder.Folders:
                count += process_folder(sub, True, filter_sender, filter_subject, filter_body,
                                        export_type, output_dir, log_widget)
        return count
    except Exception as e:
        log_widget.insert(END, f"Fehler in {folder.Name}: {e}\n")
        return count


# ---------- Haupt-Export-Funktion (wird im Thread gestartet) ----------
def start_export(folder_path, include_subfolders, filter_sender, filter_subject, filter_body,
                 export_type, output_dir, log_widget, progress_label, root):
    """
    Verbindet mit Outlook, findet den ausgewählten Ordner (anhand des Pfades) und startet die Verarbeitung.
    """
    def log(msg):
        log_widget.insert(END, msg + "\n")
        log_widget.see(END)
        root.update_idletasks()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        log("✅ Mit Outlook verbunden.")

        # Den Ordner anhand des gespeicherten Pfades wiederfinden
        target_folder = None
        for store in outlook.Folders:
            target_folder = find_folder_by_path(store, folder_path)
            if target_folder:
                break

        if not target_folder:
            log(f"❌ Ordner '{folder_path}' nicht gefunden.")
            return

        log(f"📁 Verarbeite Ordner: {folder_path} (inkl. Unterordner: {include_subfolders})")
        count = process_folder(target_folder, include_subfolders,
                               filter_sender, filter_subject, filter_body,
                               export_type, output_dir, log_widget)

        log(f"\n✨ Fertig! {count} Elemente exportiert.")
        progress_label.config(text="Export abgeschlossen.")
        messagebox.showinfo("Erfolg", f"Export beendet.\n{count} Elemente wurden gespeichert.")
    except Exception as e:
        log(f"❌ Kritischer Fehler: {e}")
        messagebox.showerror("Fehler", str(e))


def find_folder_by_path(folder, path):
    """
    Durchsucht rekursiv einen Outlook-Ordner nach einem bestimmten Pfad (z.B. "Posteingang\\Projekt").
    """
    if folder.Name == path.split("\\")[0]:
        if "\\" not in path:
            return folder
        remaining = "\\".join(path.split("\\")[1:])
        for sub in folder.Folders:
            found = find_folder_by_path(sub, remaining)
            if found:
                return found
    return None


# ---------- GUI ----------
def main():
    root = Tk()
    root.title("Outlook-Exporter Pro (mit Ordnerbaum & Konfliktauflösung)")
    root.geometry("850x750")
    root.resizable(True, True)

    style = ttk.Style()
    style.theme_use('clam')

    mainframe = ttk.Frame(root, padding="10")
    mainframe.pack(fill=BOTH, expand=True)

    # 1. Ordnerauswahl (Dropdown)
    ttk.Label(mainframe, text="Outlook-Ordner:").grid(column=0, row=0, sticky=W, pady=2)
    folder_var = StringVar()
    folder_combo = ttk.Combobox(mainframe, textvariable=folder_var, width=80, state="readonly")
    folder_combo.grid(column=0, row=1, sticky=(W, E), pady=2)
    # Ordnerliste laden
    all_folders = get_all_outlook_folders()
    folder_names = [name for name, _ in all_folders]
    folder_combo['values'] = folder_names
    if folder_names:
        folder_combo.current(0)
    # Button zum Neu-Laden der Ordner (falls sich Struktur ändert)
    def reload_folders():
        global all_folders, folder_names
        all_folders = get_all_outlook_folders()
        folder_names = [name for name, _ in all_folders]
        folder_combo['values'] = folder_names
        if folder_names:
            folder_combo.current(0)
        log_widget.insert(END, "🔄 Ordnerliste neu geladen.\n")
    btn_reload = ttk.Button(mainframe, text="Ordner neu laden", command=reload_folders)
    btn_reload.grid(column=1, row=1, padx=5)

    # 2. Option: Unterordner einschließen
    include_sub_var = BooleanVar(value=True)
    ttk.Checkbutton(mainframe, text="Unterordner einschließen", variable=include_sub_var).grid(column=0, row=2, sticky=W, pady=5)

    # 3. Filter
    filter_frame = ttk.LabelFrame(mainframe, text="Filter (leer = alle)", padding="5")
    filter_frame.grid(column=0, row=3, sticky=(W, E), pady=5)

    ttk.Label(filter_frame, text="Absender (enthält):").grid(column=0, row=0, sticky=W)
    sender_entry = ttk.Entry(filter_frame, width=40)
    sender_entry.grid(column=1, row=0, sticky=W, padx=5)

    ttk.Label(filter_frame, text="Betreff (enthält):").grid(column=0, row=1, sticky=W)
    subject_entry = ttk.Entry(filter_frame, width=40)
    subject_entry.grid(column=1, row=1, sticky=W, padx=5)

    ttk.Label(filter_frame, text="Text (enthält):").grid(column=0, row=2, sticky=W)
    body_entry = ttk.Entry(filter_frame, width=40)
    body_entry.grid(column=1, row=2, sticky=W, padx=5)

    # 4. Export-Typ
    ttk.Label(mainframe, text="Export-Typ:").grid(column=0, row=4, sticky=W, pady=5)
    export_var = StringVar(value="attachments")
    ttk.Radiobutton(mainframe, text="📎 Nur Anhänge", variable=export_var, value="attachments").grid(column=0, row=5, sticky=W)
    ttk.Radiobutton(mainframe, text="📄 Nur E-Mail-Body (als .txt)", variable=export_var, value="body").grid(column=0, row=6, sticky=W)

    # 5. Ausgabeverzeichnis
    ttk.Label(mainframe, text="Ausgabeverzeichnis:").grid(column=0, row=7, sticky=W, pady=5)
    output_frame = ttk.Frame(mainframe)
    output_frame.grid(column=0, row=8, sticky=(W, E))
    output_entry = ttk.Entry(output_frame, width=70)
    output_entry.pack(side=LEFT, fill=X, expand=True)
    def select_output():
        d = filedialog.askdirectory()
        if d:
            output_entry.delete(0, END)
            output_entry.insert(0, d)
    ttk.Button(output_frame, text="Durchsuchen...", command=select_output).pack(side=RIGHT, padx=5)

    # 6. Fortschritt / Log
    ttk.Label(mainframe, text="Log:").grid(column=0, row=9, sticky=W, pady=5)
    log_widget = scrolledtext.ScrolledText(mainframe, height=18, state=NORMAL)
    log_widget.grid(column=0, row=10, sticky=(N, S, E, W), pady=5)
    progress_label = ttk.Label(mainframe, text="Bereit.")
    progress_label.grid(column=0, row=11, sticky=W, pady=5)

    # 7. Start-Button
    def on_start():
        # Eingaben prüfen
        if not folder_var.get():
            messagebox.showwarning("Fehler", "Bitte wählen Sie einen Ordner aus.")
            return
        if not output_entry.get():
            messagebox.showwarning("Fehler", "Bitte wählen Sie ein Ausgabeverzeichnis.")
            return
        # Log leeren
        log_widget.delete(1.0, END)
        progress_label.config(text="Export läuft...")
        # Thread starten
        threading.Thread(target=start_export, args=(
            folder_var.get(),
            include_sub_var.get(),
            sender_entry.get(),
            subject_entry.get(),
            body_entry.get(),
            export_var.get(),
            output_entry.get(),
            log_widget,
            progress_label,
            root
        ), daemon=True).start()

    btn_start = ttk.Button(mainframe, text="🚀 Export starten", command=on_start)
    btn_start.grid(column=0, row=12, pady=10)

    # Gewichte für resizing
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(10, weight=1)

    root.mainloop()


if __name__ == "__main__":
    main()
