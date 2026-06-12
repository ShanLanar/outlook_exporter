# Outlook Advanced Exporter

**Exportiert E-Mails und/oder Anhänge aus Outlook 2010+ mit erweiterten Filtermöglichkeiten.**

## Features

✅ **Ordner-Dropdown** – Alle Outlook-Ordner (inkl. Pfade) in einer Auswahlliste  
✅ **Unterordner-Support** – Rekursive Suche in Unterordnern (optional)  
✅ **Multi-Filter** – Nach Absender, Betreff, Text **und Empfangsdatum (von/bis)** gleichzeitig filtern  
✅ **Anhang-Filter** – Nur bestimmte Dateitypen (z.B. `pdf, xlsx`) und optional ohne Signatur-/Inline-Bilder  
✅ **Schnell bei großen Ordnern** – Absender/Betreff/Datum werden serverseitig (Outlook-DASL) vorgefiltert  
✅ **Zwei Export-Modi:**
  - **Anhänge**: Speichert alle Dateianhänge
  - **Body**: Exportiert E-Mail-Inhalte als .txt-Dateien  
✅ **Intelligente Konfliktauflösung** – Identische Anhänge (per Inhalts-Hash) werden übersprungen, bei echten Unterschieden gewinnt die neuere Version  
✅ **Abbrechen-Button** – Laufende Exporte lassen sich jederzeit stoppen  
✅ **CSV-Exportprotokoll** – Jeder Lauf schreibt `_export_protokoll_*.csv` mit allen verarbeiteten Elementen  
✅ **Einstellungen gemerkt** – Ordner, Filter und Zielverzeichnis werden für den nächsten Start gespeichert  
✅ **Detailliertes Logging** – Thread-sicher, Fortschritt und Status in Echtzeit  

## Systemvoraussetzungen

- **Windows** mit **Outlook 2010+** (muss laufen während des Exports)
- **Python 3.10+**
- `pywin32` Paket

## Standalone-.exe bauen (kein Python auf dem Zielrechner nötig)

Für die Weitergabe an Kolleg:innen ohne Python-Installation lässt sich eine
einzelne `.exe` erzeugen:

```bat
build.bat
```

Das Skript installiert `pywin32` + `pyinstaller` und baut über
`OutlookExporter.spec` die Datei **`dist\OutlookExporter.exe`** (ohne
Konsolenfenster). Diese kann einfach kopiert und per Doppelklick gestartet
werden – Outlook muss laufen.

## Entwicklung & Tests

Die GUI-unabhängige Logik (Dateinamen, Hashing, CSV-Protokoll, Ordnersuche)
ist mit `pytest` abgedeckt und läuft ohne Windows/Outlook:

```bash
pip install pytest
pytest
```

## Installation

### 1. Abhängigkeiten installieren

```bash
pip install -r requirements.txt
```

oder manuell:

```bash
pip install pywin32
```

### 2. Programm starten

```bash
python outlook_advanced_exporter.py
```

## Bedienung

### Schritt 1: Ordner auswählen
- Dropdown zeigt alle verfügbaren Outlook-Ordner mit vollständigem Pfad
- Beispiel: `Personal-Postfach - Andreas\Inbox\Rechnungen`
- Button „Ordner neu laden" – Aktualisiert die Liste, falls sich Struktur ändert

### Schritt 2: Optionen setzen
- ☑ **Unterordner einschließen** – Durchsucht auch alle Subordner des gewählten Ordners
- **Filter (optional)**:
  - `Absender (enthält)`: z.B. `@firma.de` oder `max.mustermann`
  - `Betreff (enthält)`: z.B. `Rechnung` oder `Bericht`
  - `Text (enthält)`: z.B. `dringend`
  - *Leere Filter werden ignoriert*

### Schritt 3: Export-Typ wählen
- 📎 **Nur Anhänge** – Speichert alle Dateien aus gefilterten E-Mails
- 📄 **Nur E-Mail-Body** – Speichert Text-Inhalte als `.txt` (eine pro E-Mail)

### Schritt 4: Ausgabeordner auswählen
- Button „Durchsuchen..." – Zielverzeichnis für exportierte Dateien

### Schritt 5: Export starten
- Button „🚀 Export starten" – Durchsucht Ordner(struktur) und speichert Dateien
- Log zeigt live:
  - 📧 Gefundene E-Mails
  - ✅ Gespeicherte Dateien
  - ⏭️  Übersprungene ältere Versionen (Konfliktauflösung)

## Beispiele

### Anhänge aus "Rechnungen" exportieren
```
Ordner:        Personal-Postfach - Andreas\Rechnungen
Unterordner:   ☑ (eingeschlossen)
Absender:      [leer]
Betreff:       [leer]
Typ:           📎 Nur Anhänge
```
→ Speichert **alle PDF/Excel aus Rechnungen + Unterordnern**

### E-Mails eines Absenders als Text exportieren
```
Ordner:        Personal-Postfach - Andreas\Inbox
Unterordner:   ☑
Absender:      @abe-brands.de
Betreff:       [leer]
Typ:           📄 Nur E-Mail-Body
```
→ Speichert **Text aller E-Mails von @abe-brands.de als .txt**

### Nur dringende Rechnungen von November
```
Ordner:        Personal-Postfach - Andreas\Rechnungen
Unterordner:   ☑
Absender:      [leer]
Betreff:       Rechnung
Text:          dringend
Typ:           📎 Nur Anhänge
```
→ Speichert **Anhänge aus E-Mails mit "Rechnung" im Betreff UND "dringend" im Text**

## Konfliktauflösung (bei Anhängen)

Wenn zwei E-Mails **denselben Dateinamen** haben:
- Das Änderungsdatum beider wird verglichen
- Die **neuere Version gewinnt** und überschreibt die ältere
- Log zeigt: `🔄 Überschrieben (neuere Version): dateiname.pdf`
- Ältere Versionen werden übersprungen mit: `⏭️  Übersprungen (ältere Version): dateiname.pdf`

## Häufig gestellte Fragen

**F: Ordner-Dropdown zeigt nichts**  
A: Stellen Sie sicher, dass Outlook geöffnet und Sie angemeldet sind. Button „Ordner neu laden" drücken.

**F: Export zeigt "0 Elemente"**  
A: 
- Filter sind zu restriktiv (zu spezifische Werte)
- Ordner ist leer oder Filter passen auf keine Mails
- Unterordner-Option ist aus – aber die Mails sind in Unterordnern

**F: Sicherheitswarnung von Windows**  
A: Das ist normal beim ersten Zugriff. Erlauben Sie Outlook-Zugriff. Eventuell müssen Sie in Outlook Trustcenter → Sicherheit → „Programmatischer Zugriff" erlauben.

**F: Kann ich auch Kalender/Kontakte/Aufgaben exportieren?**  
A: Nein, dieses Tool unterstützt derzeit nur E-Mails. Erweiterung ist möglich.

**F: Export auf einem anderen Computer?**  
A: Das Tool benötigt Outlook und pywin32 auf dem Ziel-PC. Portable Installation ist nicht möglich (wegen COM-Schnittstelle).

## Fehlerbehandlung

Falls das Programm abstürzt oder Fehler zeigt:

1. **Stellen Sie sicher, Outlook läuft**
2. **Versuchen Sie „Ordner neu laden"**
3. **Ausgabeverzeichnis hat Schreibrechte?**
4. **Große Mengen E-Mails?** → Kann länger dauern. Bitte warten!

## Kontakt & Support

Bei Fragen, Fehlern oder Erweiterungswünschen bitte Bescheid geben.

---

**Version:** 1.0  
**Lizenz:** Intern  
**Letztes Update:** Juni 2026
