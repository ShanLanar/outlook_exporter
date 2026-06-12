# SETUP – Outlook Advanced Exporter

## Schritt-für-Schritt Installation

### 1. Python installieren (falls noch nicht vorhanden)

Laden Sie Python von **https://www.python.org/downloads/** herunter (Version 3.8 oder neuer).

**Wichtig:** Bei der Installation „Add Python to PATH" ☑ aktivieren!

### 2. Outlook starten

Öffnen Sie **Outlook 2010** oder neuer und melden Sie sich an. Das Programm benötigt eine laufende Outlook-Instanz.

### 3. Abhängigkeiten installieren

Öffnen Sie die **Eingabeaufforderung (CMD)** und navigieren Sie in das Verzeichnis mit den Programm-Dateien:

```bash
cd C:\Pfad\zu\outlook_exporter
```

Installieren Sie dann die benötigten Pakete:

```bash
pip install -r requirements.txt
```

oder direkt:

```bash
pip install pywin32
```

Falls die Installation fehlschlägt, versuchen Sie:

```bash
pip install --upgrade pip
pip install pywin32 --user
```

### 4. Programm starten

**Option A – Schnell (per Batch):**  
Doppelklick auf `START.bat`

**Option B – Manuell (CMD):**
```bash
python outlook_advanced_exporter.py
```

### 5. Outlook-Sicherheit (falls nötig)

Falls Windows eine Sicherheitswarnung zeigt, geben Sie der Anwendung Zugriff.

Falls Outlook selbst eine Warnung zeigt („Ein Programm versucht..."):
1. Öffnen Sie Outlook
2. Gehen Sie zu **Datei** → **Optionen** → **Sicherheitscenter** → **Sicherheitscenter-Einstellungen**
3. Klicken Sie auf **Programmatischer Zugriff**
4. Stellen Sie sicher, dass **Programmatischer Zugriff auf Outlook-Objekt-Modell zulassen** aktiviert ist
5. OK

## Fehlerbehebung

### „FileNotFoundError: Python nicht gefunden"
- Python ist nicht im PATH. Installieren Sie neu und ☑ „Add to PATH"
- Oder: Öffnen Sie CMD im Programm-Verzeichnis und geben Sie `python outlook_advanced_exporter.py` ein

### „ModuleNotFoundError: No module named 'win32com'"
- Führen Sie aus: `pip install pywin32`
- Falls das fehlschlägt: `pip install --upgrade pip` zuerst

### „Das System kann die angegebene Datei nicht finden" (pywin32)
- pywin32 hat zusätzliche Schritte nach der Installation nötig. Führen Sie aus:
  ```bash
  python -m pip install --upgrade pywin32
  python -m pywin32_postinstall -install
  ```

### Outlook zeigt "...Fehler in Ordner..."
- Stellen Sie sicher, dass der Ordner in Outlook sichtbar ist
- Versuchen Sie „Ordner neu laden" im Programm
- Schließen und öffnen Sie Outlook erneut

### Export zeigt "0 Elemente"
- Filter sind zu restriktiv → Filter leer lassen und versuchen
- Unterordner-Option ☑ aktiviert?
- Der Ordner hat tatsächlich E-Mails?

## Deinstallation

Um das Programm zu entfernen:

1. Löschen Sie das Verzeichnis mit `outlook_advanced_exporter.py`
2. Optional: `pip uninstall pywin32` (falls nicht von anderen Programmen genutzt)

---

**Bei Problemen:** Überprüfen Sie, dass:
- ✅ Outlook läuft und Sie angemeldet sind
- ✅ Python Version 3.6+ installiert ist
- ✅ `pip install -r requirements.txt` erfolgreich war
- ✅ Sie Administrator-Rechte haben (evt. CMD als Admin starten)
