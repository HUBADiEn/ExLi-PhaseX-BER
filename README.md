# ExLi PhaseX - Berechnung

PMS/PIM Vergleichstool für die Phase X Berechnung.

Vergleicht zwei CSV-Dateien (PMS und PIM) anhand der EAN, führt 15 Prüfungen durch und exportiert die Ergebnisse als Excel-Datei.

## Quick Start

1. Alle `.ps1`-Dateien in denselben Ordner legen
2. `Start_v1.6.ps1` ausführen (Rechtsklick → "Mit PowerShell ausführen")
3. PMS- und PIM-Datei auswählen
4. Ergebnisse werden als Excel-Datei exportiert

---

## Modulstruktur

| Datei | Version | Beschreibung |
|-------|---------|--------------|
| `Start_v1.6.ps1` | 1.6 | **Einstiegspunkt** - Versionsprüfung und Start |
| `main_v1.2.ps1` | 1.2 | Hauptlogik, Datenverarbeitung, Export |
| `config_v1.0.ps1` | 1.0 | Konstanten, Lookup-Tabellen, Header-Definitionen |
| `functions-checks_v1.3.ps1` | 1.3 | Check-Funktionen 1-14 |
| `functions-excel_v1.0.ps1` | 1.0 | Excel-Formatierung, Summary-Rows |
| `functions-dialogs_v1.0.ps1` | 1.0 | Datei-Dialoge, CSV-Export |
| `functions-helpers_v1.0.ps1` | 1.0 | Hilfsfunktionen (ZeitDiff-Berechnung) |

### Abhängigkeiten

```
Start_v1.6.ps1
    ├── config_v1.0.ps1
    ├── functions-excel_v1.0.ps1
    ├── functions-dialogs_v1.0.ps1
    ├── functions-helpers_v1.0.ps1
    ├── functions-checks_v1.3.ps1
    └── main_v1.2.ps1
```

---

## Checks Übersicht

| Check | Name | Vergleicht | Besonderheiten |
|-------|------|------------|----------------|
| 0 | Vorhanden in beiden Quellen | EAN in PMS und PIM | - |
| 1 | Status | PMS_SLLPAS ↔ PIM_Status | - |
| 2 | Kategorie | PMS_SLLCAT ↔ PIM_Kategorie | "UKN" + leer = ok |
| 3 | Genre | PMS_SLLGNR ↔ PIM_Genre | Array-Vergleich |
| 4 | Preiscode | PMS_SLLPCD ↔ PIM_Preiscode | - |
| 5 | Standard VP | PMS_FLGSTP ↔ PIM_Standard VP | - |
| 6 | Fixer VP | PMS_FLGFXP ↔ PIM_Fixer VP | - |
| 7 | Release-Datum | PMS_RELDAT ↔ PIM_Release Date | Datumsformat-Konvertierung |
| 8 | Errorcode | PMS_SLLERR ↔ PIM_Errorcode | - |
| 9 | VP | PMS_SLLVPL ↔ PIM_VP | Buch-Kat Warnung, PMS 0 = PIM leer |
| 10 | PrioEP | PMS_SLLEPL ↔ PIM_PrioEP | Toleranz: absolut 0.02 ODER relativ 0.01%, PMS 0 = PIM leer |
| 11 | RgEP | PMS_SLOEPF ↔ PIM_RgEP | PMS 0 = PIM leer |
| 12 | Tiefpreis | PMS_REDVPL ↔ PIM_Tiefpreis | Warnung bei anderem priorisierten Lf, PMS 0 = PIM leer |
| 13 | L-Prio Fehlercode | PMS_SAAPNT ↔ PIM_Fehlercode | Title-Warnung bei 999914 |
| 14 | L-Prio | PMS_SAAPNT ↔ PIM_L-Prio-Punkte | Korrelation mit PrioEP-Diff (Faktor 250) |

---

## Check-Details

### Check 9: VP (Verkaufspreis)
- **OK:** Werte identisch
- **OK:** PMS = 0 und PIM = leer
- **Warnung:** Buch-Kategorie (B, B-EN, B-FR) mit Standard VP und unterschiedlichen Werten
- **Nicht OK:** Sonstige Unterschiede

### Check 10: PrioEP (Priorisierter Einkaufspreis)
- **OK:** Werte identisch
- **OK:** PMS = 0 und PIM = leer
- **OK:** Differenz ≤ 0.02 (absolute Toleranz)
- **OK:** Differenz ≤ 0.0001 × PMS-Wert (relative Toleranz, 0.01%)
- **Nicht OK:** Differenz ausserhalb beider Toleranzen

### Check 12: Tiefpreis
- **OK:** Werte identisch
- **OK:** PMS = 0 und PIM = leer
- **OK:** PMS VP = PMS Tiefpreis (wird als "kein Tiefpreis" behandelt)
- **Warnung:** Anderer Lieferant priorisiert (SAASEL = 0, REDVPL < SLLVPL, Check 9 = ok)
- **Nicht OK:** Sonstige Unterschiede

### Check 14: L-Prio (Lieferanten-Priorität)
Prüfungsreihenfolge:
1. Status = passive → OK
2. SAAPNT ≥ 900000 und identisch mit PIM_Fehlercode → OK (Fehlercode)
3. PIM_L-Prio leer und Check 13 hat Title-Warnung → Warnung
4. Werte identisch → OK
5. Check 9 hat Buch-Kat Warnung → Warnung
6. Check 10 = nicht ok UND |L-Prio Diff| < 250 × |PrioEP Diff| → Warnung (kommt von PrioEP)
7. Sonst → Nicht OK

---

## Versionierung

### Namenskonvention
Jede Datei hat die Version im Dateinamen: `modulname_v{version}.ps1`

### Versionsprüfung
`Start.ps1` prüft beim Start, ob alle Module in der erwarteten Version vorhanden sind. Bei Versionskonflikten wird eine Fehlermeldung angezeigt.

### Script-Version
Die angezeigte Script-Version (z.B. "Berechnung_V1.6") wird von `Start.ps1` gesetzt und automatisch an alle Module weitergegeben.

---

## Änderungshistorie

### Start_v1.6
- ScriptVersion wird zentral gesetzt ("Berechnung_V1.6")
- Erwartet main_v1.2

### main_v1.2
- PrioEP Diff wird immer berechnet wenn Werte unterschiedlich (auch bei Toleranz-OK)
- Wichtig für Check 14 Korrelationsprüfung

### functions-checks_v1.3
- Check 14: Korrelation mit PrioEP-Diff hinzugefügt
- Check 9, 10, 11, 12: PMS "0" entspricht leerem PIM-Feld
- Check 10: Relative Toleranz (0.01%) zusätzlich zu absoluter Toleranz (0.02)

---

## Technische Anforderungen

- **PowerShell:** Version 5.1 oder höher
- **Excel-Modul:** ImportExcel (wird bei Bedarf automatisch installiert)
- **Encoding:** UTF-8 mit BOM

---

## Ordnerstruktur

```
PhaseX_Berechnung/
├── Start_v1.6.ps1              ← Dieses Script starten
├── main_v1.2.ps1
├── config_v1.0.ps1
├── functions-checks_v1.3.ps1
├── functions-excel_v1.0.ps1
├── functions-dialogs_v1.0.ps1
├── functions-helpers_v1.0.ps1
└── README.md
```

---

## Support

Bei Fragen oder Änderungswünschen: Issue erstellen oder direkt im Chat mit Claude besprechen.

Repository: https://github.com/HUBADiEn/ExLi-PhaseX-BER
