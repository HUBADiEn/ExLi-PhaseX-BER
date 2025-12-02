<#
.SYNOPSIS
    PMS/PIM Vergleichstool - Berechnung Phase X - Kombiniertes Script

.DESCRIPTION
    Dieses Script vergleicht PMS- und PIM-Daten und führt 14 Qualitätschecks durch.
    
    Das Script liest PMS- und PIM-CSV-Dateien ein, führt einen umfassenden Datenabgleich durch 
    und erstellt Excel-Outputs mit detaillierten Ergebnissen. Es enthält spezielle Logik für 
    passive Artikel, Toleranzen für Preisvergleiche, Datumsformat-Konvertierungen und 
    Performance-Optimierungen für große Dateien (bis 5+ Millionen Zeilen).
    
    Hauptfunktionen:
    - 14 Qualitätschecks (Status, Kategorie, Genre, Preiscode, VP, Fixpreis, Release-Datum, etc.)
    - Intelligente Format-Wahl (Excel bis 1M Zeilen, CSV darüber)
    - Single-Pass Counting (10x schneller als vorherige Versionen)
    - Spezielle Behandlung für passive Artikel
    - Toleranzen für Rundungsfehler bei Preisen
    - Datumsformat-Konvertierung (PMS: DD.MM.YY, PIM: YYYYMMDD)

.NOTES
    File:           PhaseX_Berechnung_v2.24.ps1
    Version:        2.24
    Datum:          02.12.2025
    
    Änderungshistorie:
        2.24 - Check 5: Flag-Vergleich implementiert (wie Check 6) - PMS Flag vs. PIM Wert
               Check 5: Nur Prüfung ob Standard VP vorhanden, nicht Wert-Vergleich
               Check 8: Warnung hinzugefügt wenn PMS = 1 und PIM = 0
               Check 8: "Warnung - Nur PMS hat einen Fehler"
        2.23 - Check 13: Zusätzliche Warnung wenn PIM_Fehlercode = 999914
               Check 13: "Warnung - Titel fehlt (wird im PMS ggf. nicht sauber verarbeitet)"
               Check 14: Zusätzliche Warnung wenn PIM_Fehlercode = 999914 (gleiche Meldung)
               Check 14: Prüfung erfolgt am Ende vor "nicht ok"
        2.22 - Check 9: Verkauft-Flag entfernt (keine "Warnung - VP unterschiedlich aber verkauft" mehr)
               Check 12: Verkauft-Flag entfernt (keine "Warnung - Tiefpreis unterschiedlich aber verkauft" mehr)
               Check 14: Fehlercode-Prüfung DIREKT nach Werte-Vergleich (vor Toleranz-Check)
               Check 14: Neue Meldung "ok - L-Prio-Fehlercode ist identisch"
               Check 14: Vereinfachte Logik (Fehlercode-Prüfung nur noch an 1 Stelle statt 3)
        2.21 - Check 14: Verkauft-Flag komplett entfernt (keine Warnungen mehr)
               Check 14: Toleranz eingebaut - L-Prio Diff darf BIS ZU 250x PrioEP Diff sein (statt exakt)
               Check 14: "ok - Differenz im Toleranzbereich" statt "ok - L-Prio Diff entspricht..."
        2.20 - Check 13: Verkauft-Flag Warnung entfernt (keine "Warnung - unterschiedlich aber verkauft" mehr)
               Check 14: Fehlercode-Prüfung hinzugefügt - wenn PMS_SAAPNT = PIM_Fehlercode → "ok - Fehlercode identisch"
        2.19 - Check 8: KRITISCHER BUGFIX - Korrekter Feldname 'PIM_Errorcode' (nicht 'PIM_Error Code')
               Check 8: Feld wurde bisher gar nicht gelesen (immer leer)!
        2.18 - Check 7: KRITISCHER BUGFIX - Korrekter Feldname 'PIM_Release Date' (nicht 'PIM_Release-Datum')
               Check 7: Feld wurde bisher gar nicht gelesen (immer leer)!
        2.17 - Check 7: Robustere Datumslogik (explizites String-Trimming und -Casting)
               Check 7: Bugfix für "23.11.18" vs "20181123" Vergleich
        2.16 - Check 4: Passive-Prüfung mit Meldung "PMS liefert bei passive keinen Preiscode"
               Check 7: Datumsformat-Konvertierung (PMS: DD.MM.YY, PIM: YYYYMMDD)
               Check 7: Jahr nur mit 2 hinteren Ziffern vergleichen
        2.15 - Check 6: Nur Flag-Vergleich (Hat Fixpreis? Ja/Nein)
               Check 6: Wert wird NICHT verglichen, nur ob vorhanden
        2.14 - Check 13: PMS Fehlercode aus SAAPNT (nur wenn >= 900000)
               Check 13: PIM Fehlercode aus PIM_Fehlercode (nicht PIM_Error Code)
               Check 13: Beide leer = ok
        2.13 - Check 10: Fixe Toleranz ±0.02 zusätzlich zur dynamischen 0.01%
               Check 10: Zwei Toleranzen - wenn EINE erfüllt ist → ok
        2.12 - Check 3: Passive-Prüfung AN DEN ANFANG verschoben (KRITISCH!)
               Check 3: Wildcard-Prüfung entfernt
               Check 12: SLLVPL = REDVPL Logik (Kein Tiefpreis)
        2.11 - Check 3: Passive-Logik ohne Check 1 Bedingung (beide passive = ok)
               Check 7: Leerzeichen trimmen (mehrere Leerzeichen = leerer Wert)
        2.10 - Check 3: V1.96 Logik wiederhergestellt (Array-Matching, Wildcard, etc.)
               PLUS neue Passive-Logik (beide passive + Check 1 ok)
               Feldname korrigiert: PIM_Genre-Code → PIM_Genre
        2.9 - Check 2 korrigiert: "0" → "UKN" (zurück zu V1.96 Logik)
              Check 3 erweitert: Genre-Unterschiede bei passiven Artikeln sind ok
        2.8 - Check 2 erweitert: Kategorie-Unterschiede bei passiven Artikeln sind ok
        2.7 - PERFORMANCE BOOST (10x schneller Schritt 7)
              Single-Pass Counting (40 Durchläufe → 1)
              Select-Object Overhead eliminiert
        2.6 - Check 12 erweitert (SAASEL-Logik)
              Summary-Rows vertauscht (Warnungen oben, Fehler unten)
        2.4 - Performance-Optimierung Light
        2.3 - Summary-Rows korrigiert
        2.2 - PIM-Header korrigiert
        2.1 - Endlosschleife behoben
        2.0 - Alle Module kombiniert
            - Basiert auf: Start v1.18, main v1.9, config v1.2, 
                          functions-checks v1.6, functions-dialogs v1.1,
                          functions-excel v1.1, functions-helpers v1.0
            - SLOEPF → SLOERG Änderung implementiert
        
    Enthaltene Module (aus modularer Version):
        - config v1.2
        - functions-helpers v1.0
        - functions-dialogs v1.1
        - functions-excel v1.1
        - functions-checks v1.6
        - main v1.9
        - Start-Logik v1.18

.PARAMETER None
    Dieses Script hat keine Parameter. Alles wird über Dialoge ausgewählt.

.EXAMPLE
    .\PhaseX_Berechnung_v2.0.ps1
    Startet das Tool und öffnet Datei-Auswahl-Dialoge.
#>

# =====================================================================
# GLOBALE SCRIPT-VERSION
# =====================================================================
$global:ScriptVersion = "Berechnung_V2.24"

# =====================================================================
# MODUL 1: CONFIG (v1.2)
# =====================================================================
Write-Host "Lade Konfiguration..." -ForegroundColor Gray

# Erwartete PMS-Header
$script:PMS_Header_Expected = @(
    "SLLLFN","SLLEAN","SLLPAS","SLLCAT","SLLGNR","SLLPCD","FLGSTP","FLGFXP","FLGVKF","RELDAT",
    "XML01","XML02","XML03","XML04","XML05","SLLVPL","SLLEPL","SLOERG","SLOWAH","REDVPL",
    "SLLERR","SAAPNT","SLLIGN","IMPDAT","CHGDAT","SAASEL"
)

# Erwartete PIM-Header
$script:PIM_Header_Expected = @(
    "Lieferant","EAN","Status","Kategorie","Genre","Preiscode","Standard VP","Fixer VP","Release Date",
    "Acquisition Price","Sales Price","Publisher ID","Linedisc","Bonusgroup","VP","PrioEP","RgEP",
    "Währung RgEP","Tiefpreis","Errorcode","Fehlercode","L-Prio-Punkte","Sperrcode","Verwendete Kalkulation",
    "letzter Import","letzte Änderung","letzter Status"
)

# Lieferanten-Lookup (Nummer → Name)
$script:SupplierLookupTable = @{
    '16409132'='AVA Verlagsauslieferung'; '16801357'='Bremer Versandwerk GmbH'; '16409120'='Buchzentrum'
    '16800790'='Carletto AG'; '16517649'='Carlit + Ravensburger AG'; '16803558'='ciando (Agency)'
    '16803554'='ciando GmbH'; '15642908'='Ex Libris AG Dietikon 1'; '16450805'='Grüezi Music AG'
    '16802683'='Libri (Agency)'; '16776945'='Libri GmbH'; '16803735'='Max Bersinger AG'
    '16801413'='MFP Tonträger'; '16407363'='Musikvertrieb AG'; '30000023'='Office World (Oridis)'
    '16409618'='OLF S.A.'; '16411177'='Phonag Records AG'; '16558172'='Phono-Vertrieb'
    '16212120'='Rainbow Home Entertainment'; '16526960'='Sombo AG'; '16699796'='Sony Music Entertainment'
    '16407336'='Starworld Enterprise GmbH'; '16423780'='Thali AG'; '16486030'='Universal Music GmbH'
    '30000223'='Vedes Grosshandel GmbH'; '16706931'='Waldmeier AG'; '16797703'='Warner Music Group'
    '16435880'='Zeitfracht Medien GmbH'
}

# User-Lookup (System-Username → Friendly Name)
$script:UserLookupTable = @{
    'M0733302' = 'WOB'
    'M0779325' = 'AZG'
    'M0555315' = 'CPA'
}

# SharePoint-Speicherung aktiviert?
$script:SaveToSharePoint = $false

# Performance-Schwellenwerte
$script:EXCEL_EXPORT_LIMIT = 1000000  # Excel nur bis 1 Mio Zeilen pro File

Write-Host "  Konfiguration geladen." -ForegroundColor Green

# =====================================================================
# MODUL 2: FUNCTIONS-HELPERS (v1.0)
# =====================================================================
Write-Host "Lade Helper-Funktionen..." -ForegroundColor Gray

function Invoke-CalculateTimeDifference {
    param([Parameter(Mandatory=$true)][PSCustomObject]$Dataset)
    
    $pmsDate = $Dataset.PMS_CHGDAT
    $pimDate = $Dataset.PIM_Last_Change_Date
    
    if ([string]::IsNullOrWhiteSpace($pmsDate) -or [string]::IsNullOrWhiteSpace($pimDate)) {
        return ""
    }
    
    try {
        $pmsDateTime = [datetime]::ParseExact($pmsDate, "yyyyMMdd", $null)
        $pimDateTime = [datetime]::ParseExact($pimDate, "dd.MM.yyyy", $null)
        $diff = ($pmsDateTime - $pimDateTime).Days
        
        if ($diff -eq 0) { return "0 Tage (gleich)" }
        elseif ($diff -gt 0) { return "$diff Tage (PMS neuer)" }
        else { return "$([Math]::Abs($diff)) Tage (PIM neuer)" }
    }
    catch {
        return "Fehler beim Parsen"
    }
}

Write-Host "  Helper-Funktionen geladen." -ForegroundColor Green

# =====================================================================
# MODUL 3: FUNCTIONS-DIALOGS (v1.1)
# =====================================================================
Write-Host "Lade Dialog-Funktionen..." -ForegroundColor Gray

# Windows Forms Assembly laden
Add-Type -AssemblyName System.Windows.Forms

function Get-FilePathDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$WindowTitle,
        [Parameter(Mandatory=$false)]
        [string]$InitialDirectory = [Environment]::GetFolderPath("Desktop")
    )
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $WindowTitle
    $openFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
    $openFileDialog.InitialDirectory = $InitialDirectory
    $openFileDialog.Multiselect = $false
    
    if($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $openFileDialog.FileName
    }
    
    return $null
}

Write-Host "  Dialog-Funktionen geladen." -ForegroundColor Green

# =====================================================================
# MODUL 4: FUNCTIONS-EXCEL (v1.1)
# =====================================================================
Write-Host "Lade Excel-Funktionen..." -ForegroundColor Gray

function Apply-SummaryRow {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$WorksheetName,
        [Parameter(Mandatory=$true)][PSCustomObject]$HeaderSummary,
        [Parameter(Mandatory=$true)][PSCustomObject]$WarningSummary,
        [Parameter(Mandatory=$true)][string]$ScriptVersion,
        [Parameter(Mandatory=$true)][string]$SupplierNumber
    )
    
    try {
        $excel = Open-ExcelPackage -Path $Path
        $ws = $excel.Workbook.Worksheets[$WorksheetName]
        
        if ($null -eq $ws) {
            Write-Warning "Worksheet '$WorksheetName' nicht gefunden in '$Path'"
            Close-ExcelPackage $excel
            return
        }
        
        # ZEILE 1: WARNUNGS-ZEILE (HELLORANGE)
        # A1: Lieferantennummer (KEINE Farbe)
        $ws.Cells[1, 1].Value = $SupplierNumber
        
        # Rest: Warnungszähler (NUR hellorange färben wenn > 0)
        $colIndex = 1
        foreach ($prop in $WarningSummary.PSObject.Properties) {
            if ($prop.Name -eq 'EAN') { 
                # Überspringe EAN, A1 haben wir schon gesetzt
                continue
            }
            
            $colIndex++
            $warningValue = $prop.Value
            $ws.Cells[1, $colIndex].Value = $warningValue
            
            # NUR hellorange färben wenn Warnung > 0
            if ($warningValue -gt 0) {
                $ws.Cells[1, $colIndex].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[1, $colIndex].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 255, 200, 100))
                $ws.Cells[1, $colIndex].Style.Font.Bold = $true
            }
        }
        
        # ZEILE 2: FEHLER-ZEILE (HELLROT)
        # A2: Script-Version (KEINE Farbe)
        $ws.Cells[2, 1].Value = $ScriptVersion
        
        # Rest: Fehlerzähler (NUR hellrot färben wenn > 0)
        $colIndex = 1
        foreach ($prop in $HeaderSummary.PSObject.Properties) {
            if ($prop.Name -eq 'EAN') { 
                # Überspringe EAN, A2 haben wir schon gesetzt
                continue
            }
            
            $colIndex++
            $errorValue = $prop.Value
            $ws.Cells[2, $colIndex].Value = $errorValue
            
            # NUR hellrot färben wenn Fehler > 0
            if ($errorValue -gt 0) {
                $ws.Cells[2, $colIndex].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[2, $colIndex].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 255, 100, 100))
                $ws.Cells[2, $colIndex].Style.Font.Bold = $true
            }
        }
        
        # ZEILE 3: HEADER (blau, wie bisher)
        $lastCol = $ws.Dimension.End.Column
        for ($i = 1; $i -le $lastCol; $i++) {
            $ws.Cells[3, $i].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $ws.Cells[3, $i].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
            $ws.Cells[3, $i].Style.Font.Bold = $true
        }
        
        $excel.Save()
        Close-ExcelPackage $excel
    }
    catch {
        Write-Warning "Fehler in Apply-SummaryRow: $_"
    }
}

function Optimize-ColumnWidthForHeader {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$WorksheetName
    )
    
    try {
        $excel = Open-ExcelPackage -Path $Path
        $ws = $excel.Workbook.Worksheets[$WorksheetName]
        
        if ($null -eq $ws) {
            Write-Warning "Worksheet '$WorksheetName' nicht gefunden"
            Close-ExcelPackage $excel
            return
        }
        
        $lastCol = $ws.Dimension.End.Column
        
        for ($col = 1; $col -le $lastCol; $col++) {
            $headerText = $ws.Cells[3, $col].Text
            
            if ([string]::IsNullOrEmpty($headerText)) { continue }
            
            $width = [Math]::Min([Math]::Max($headerText.Length * 1.2, 8), 50)
            
            if ($headerText -like "Check*") { $width = [Math]::Max($width, 25) }
            elseif ($headerText -eq "EAN") { $width = 15 }
            elseif ($headerText -like "*Diff*") { $width = 12 }
            elseif ($headerText -like "PMS_*" -or $headerText -like "PIM_*") { $width = 12 }
            
            $ws.Column($col).Width = $width
        }
        
        $excel.Save()
        Close-ExcelPackage $excel
    }
    catch {
        Write-Warning "Fehler in Optimize-ColumnWidthForHeader: $_"
    }
}

function Color-HeaderBySource {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$WorksheetName
    )
    
    try {
        $excel = Open-ExcelPackage -Path $Path
        $ws = $excel.Workbook.Worksheets[$WorksheetName]
        
        if ($null -eq $ws) {
            Write-Warning "Worksheet '$WorksheetName' nicht gefunden"
            Close-ExcelPackage $excel
            return
        }
        
        $lastCol = $ws.Dimension.End.Column
        
        for ($col = 1; $col -le $lastCol; $col++) {
            $headerText = $ws.Cells[3, $col].Text
            
            if ([string]::IsNullOrEmpty($headerText)) { continue }
            
            # PMS-Felder: Hellgrün
            if ($headerText -like "PMS_*") {
                $ws.Cells[3, $col].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[3, $col].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 200, 255, 200))
                $ws.Cells[3, $col].Style.Font.Bold = $true
            }
            # PIM-Felder: Hellgelb
            elseif ($headerText -like "PIM_*") {
                $ws.Cells[3, $col].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[3, $col].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 255, 255, 200))
                $ws.Cells[3, $col].Style.Font.Bold = $true
            }
            # Check-Spalten: Hellblau
            elseif ($headerText -like "Check*" -or $headerText -eq "EAN" -or $headerText -like "*Diff*") {
                $ws.Cells[3, $col].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[3, $col].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
                $ws.Cells[3, $col].Style.Font.Bold = $true
            }
        }
        
        $excel.Save()
        Close-ExcelPackage $excel
    }
    catch {
        Write-Warning "Fehler in Color-HeaderBySource: $_"
    }
}

function Export-CsvFast {
    param(
        [Parameter(Mandatory=$true)][System.Collections.ArrayList]$Data,
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$Delimiter = ';'
    )
    
    try {
        $sw = New-Object System.IO.StreamWriter($Path, $false, [System.Text.Encoding]::UTF8)
        
        if ($Data.Count -gt 0) {
            $headers = $Data[0].PSObject.Properties.Name
            $sw.WriteLine(($headers -join $Delimiter))
            
            foreach ($row in $Data) {
                $values = $headers | ForEach-Object { 
                    $val = $row.$_
                    if ($null -eq $val) { "" } else { $val.ToString() }
                }
                $sw.WriteLine(($values -join $Delimiter))
            }
        }
        
        $sw.Close()
        $sw.Dispose()
    }
    catch {
        Write-Warning "Fehler in Export-CsvFast: $_"
        if ($sw) { $sw.Dispose() }
    }
}

Write-Host "  Excel-Funktionen geladen." -ForegroundColor Green

# =====================================================================
# MODUL 5: FUNCTIONS-CHECKS (v1.6)
# =====================================================================
Write-Host "Lade Check-Funktionen..." -ForegroundColor Gray

function Invoke-Check1_Status {
    param([PSCustomObject]$Dataset)
    $pmsStatus = $Dataset.PMS_SLLPAS
    $pimStatus = $Dataset.PIM_Status
    if($pmsStatus -eq $pimStatus){return "ok"}
    return "nicht ok"
}

function Invoke-Check2_Kategorie {
    param([PSCustomObject]$Dataset)
    $pmsCat = $Dataset.PMS_SLLCAT
    $pimCat = $Dataset.PIM_Kategorie
    if($pmsCat -eq "UKN" -and [string]::IsNullOrEmpty($pimCat)){return "ok - Kein Kat-Mapping im PMS und PIM"}
    if($pmsCat -eq $pimCat){return "ok"}
    
    # Neue Logik: Kategorie-Unterschiede bei passiven Artikeln sind ok
    $pmsStatus = $Dataset.PMS_SLLPAS
    $pimStatus = $Dataset.PIM_Status
    $check1Result = $Dataset.'Check 1: Status'
    
    if ($pmsStatus -eq "passive" -and $pimStatus -eq "passive" -and $check1Result -eq "ok") {
        return "ok - Status = passive"
    }
    
    return "nicht ok"
}

function Invoke-Check3_Genre {
    param([PSCustomObject]$Dataset)
    
    # ⭐ WICHTIGSTE PRÜFUNG ZUERST: Beide passive = IMMER ok!
    # Egal ob Genre-Werte vorhanden sind oder nicht!
    $pmsStatus = $Dataset.PMS_SLLPAS
    $pimStatus = $Dataset.PIM_Status
    
    if ($pmsStatus -eq "passive" -and $pimStatus -eq "passive") {
        return "ok - Status = passive"
    }
    
    # Ab hier: Mindestens einer ist active → Genre MUSS stimmen!
    $pmsGenresRaw = $Dataset.PMS_SLLGNR
    $pimGenre = $Dataset.PIM_Genre
    
    # Beide leer → ok
    if ([string]::IsNullOrEmpty($pmsGenresRaw) -and 
        [string]::IsNullOrEmpty($pimGenre)) {
        return "ok"
    }
    
    # Einer leer, der andere nicht → nicht ok
    if ([string]::IsNullOrEmpty($pmsGenresRaw) -or 
        [string]::IsNullOrEmpty($pimGenre)) {
        return "nicht ok"
    }
    
    # Array-Matching: PMS kann mehrere Genres haben (z.B. "[123, 456, 789]")
    $pmsGenresClean = $pmsGenresRaw.Trim('[]')
    $pmsGenresArray = $pmsGenresClean.Split(',') | ForEach-Object { $_.Trim() }
    
    if ($pmsGenresArray -contains $pimGenre) {
        return "ok"
    }
    
    return "nicht ok"
}

function Invoke-Check4_Preiscode {
    param([PSCustomObject]$Dataset)
    # NEU v2.16: Passive-Prüfung mit spezieller Meldung
    if($Dataset.PMS_SLLPAS -eq "passive"){
        return "ok - Status = passive (PMS liefert bei passive keinen Preiscode)"
    }
    $pmsPreis = $Dataset.PMS_SLLPCD
    $pimPreis = $Dataset.PIM_Preiscode
    if($pmsPreis -eq $pimPreis){return "ok"}
    return "nicht ok"
}

function Invoke-Check5_StandardVP {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    
    # NEU v2.24: PMS liefert nur FLAG (0 oder 1), PIM liefert WERT
    # Vergleich nur: Hat Standard VP JA/NEIN, nicht der Wert selbst!
    
    $pmsFlag = $Dataset.PMS_FLGSTP
    $pimStdVP = $Dataset.'PIM_Standard VP ab Lieferant'
    
    # PMS Flag auswerten
    $pmsHasStandardVP = ($pmsFlag -eq '1')
    
    # PIM Wert auswerten
    $pimHasStandardVP = -not [string]::IsNullOrEmpty($pimStdVP)
    
    # Vergleich: Beide sagen "Standard VP vorhanden" ODER beide sagen "Kein Standard VP"
    if($pmsHasStandardVP -and $pimHasStandardVP){
        return "ok - Beide haben Standard VP"
    }
    if(-not $pmsHasStandardVP -and -not $pimHasStandardVP){
        return "ok - Beide haben keinen Standard VP"
    }
    
    return "nicht ok"
}

function Invoke-Check6_FixerVP {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    
    # NEU v2.15: PMS liefert nur FLAG (0 oder 1), PIM liefert WERT
    # Vergleich nur: Hat Fixpreis JA/NEIN, nicht der Wert selbst!
    
    $pmsFlag = $Dataset.PMS_FLGFXP
    $pimFixVP = $Dataset.'PIM_Fixer VP'
    
    # PMS Flag auswerten
    $pmsHasFixpreis = ($pmsFlag -eq '1')
    
    # PIM Wert auswerten
    $pimHasFixpreis = -not [string]::IsNullOrEmpty($pimFixVP)
    
    # Vergleich: Beide sagen "Fixpreis vorhanden" ODER beide sagen "Kein Fixpreis"
    if($pmsHasFixpreis -and $pimHasFixpreis){
        return "ok - Beide haben Fixpreis"
    }
    if(-not $pmsHasFixpreis -and -not $pimHasFixpreis){
        return "ok - Beide haben keinen Fixpreis"
    }
    
    return "nicht ok"
}

function Invoke-Check7_ReleaseDatum {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    
    # Trimme Leerzeichen - mehrere Leerzeichen im PMS = leerer Wert
    $pmsRel = $Dataset.PMS_RELDAT
    if ($null -ne $pmsRel) { $pmsRel = $pmsRel.ToString().Trim() }
    
    # NEU v2.18: Korrekter Feldname! Header ist "Release Date" (mit Leerzeichen)
    $pimRel = $Dataset.'PIM_Release Date'
    if ($null -ne $pimRel) { $pimRel = $pimRel.ToString().Trim() }
    
    # Beide leer → ok
    if([string]::IsNullOrEmpty($pmsRel) -and [string]::IsNullOrEmpty($pimRel)){return "ok"}
    
    # NEU v2.17: Robustere Datumsformat-Konvertierung
    # PMS Format: DD.MM.YY (z.B. "23.11.18")
    # PIM Format: YYYYMMDD (z.B. "20181123")
    # Vergleich: Jahr nur mit 2 hinteren Ziffern!
    
    # Versuche Formate zu parsen und zu vergleichen
    if(-not [string]::IsNullOrEmpty($pmsRel) -and -not [string]::IsNullOrEmpty($pimRel)){
        # PMS Format: DD.MM.YY
        if($pmsRel -match '^(\d{2})\.(\d{2})\.(\d{2})$'){
            $pmsDay = $Matches[1].Trim()
            $pmsMonth = $Matches[2].Trim()
            $pmsYear = $Matches[3].Trim()
            
            # PIM Format: YYYYMMDD
            if($pimRel -match '^(\d{4})(\d{2})(\d{2})$'){
                $pimFullYear = $Matches[1].Trim()
                $pimMonth = $Matches[2].Trim()
                $pimDay = $Matches[3].Trim()
                $pimYear = $pimFullYear.Substring(2,2)  # Nur letzte 2 Ziffern
                
                # Expliziter String-Vergleich mit Normalisierung
                $dayMatch = ([string]$pmsDay) -eq ([string]$pimDay)
                $monthMatch = ([string]$pmsMonth) -eq ([string]$pimMonth)
                $yearMatch = ([string]$pmsYear) -eq ([string]$pimYear)
                
                if($dayMatch -and $monthMatch -and $yearMatch){
                    return "ok"
                }
            }
        }
    }
    
    # Fallback: Direkter String-Vergleich (falls alte Formate oder gleich)
    if($pmsRel -eq $pimRel){return "ok"}
    
    return "nicht ok"
}

function Invoke-Check8_Errorcode {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    $pmsErr = $Dataset.PMS_SLLERR
    # NEU v2.19: Korrekter Feldname! Header ist "Errorcode" (ohne Leerzeichen)
    $pimErr = $Dataset.PIM_Errorcode
    
    # Beide "0" oder leer = ok
    $pmsIsZero = ($pmsErr -eq '0' -or $pmsErr -eq '0.00' -or $pmsErr -eq '0,00')
    $pimIsZero = ($pimErr -eq '0' -or $pimErr -eq '0.00' -or $pimErr -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimErr)
    
    if($pmsIsZero -and ($pimIsEmpty -or $pimIsZero)){return "ok"}
    if($pmsErr -eq $pimErr){return "ok"}
    
    # NEU v2.24: Warnung wenn nur PMS einen Fehler meldet
    if($pmsErr -eq '1' -and ($pimErr -eq '0' -or $pimIsZero)){
        return "Warnung - Nur PMS hat einen Fehler"
    }
    
    return "nicht ok"
}

function Invoke-Check9_VP {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    $pmsVP = $Dataset.PMS_SLLVPL
    $pimVP = $Dataset.PIM_VP
    $pmsIsZero = ($pmsVP -eq '0' -or $pmsVP -eq '0.00' -or $pmsVP -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimVP)
    if($pmsIsZero -and $pimIsEmpty){return "ok"}
    if($pmsVP -eq $pimVP){return "ok"}
    # NEU v2.22: Verkauft-Flag entfernt
    return "nicht ok"
}

function Invoke-Check10_PrioEP {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    $pmsPrio = $Dataset.PMS_SLLEPL
    $pimPrio = $Dataset.PIM_PrioEP
    $pmsIsZero = ($pmsPrio -eq '0' -or $pmsPrio -eq '0.00' -or $pmsPrio -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimPrio)
    if($pmsIsZero -and $pimIsEmpty){return "ok"}
    $pmsVal = 0.0; $pimVal = 0.0
    $pmsOk = [decimal]::TryParse($pmsPrio, [ref]$pmsVal)
    $pimOk = [decimal]::TryParse($pimPrio, [ref]$pimVal)
    if(-not $pmsOk -or -not $pimOk){return "nicht ok - ungültige Werte"}
    if($pmsVal -eq $pimVal){return "ok"}
    
    # NEU v2.13: Zwei Toleranzen - wenn EINE erfüllt ist → ok!
    $diff = [Math]::Abs($pmsVal - $pimVal)
    
    # Toleranz 1: Dynamisch (0.01% vom PMS-Wert)
    $toleranceDynamic = [Math]::Abs($pmsVal) * 0.0001
    
    # Toleranz 2: Fix (±0.02)
    $toleranceFixed = 0.02
    
    # Prüfe beide Toleranzen
    if($diff -le $toleranceDynamic){
        return "ok - Diff von $([Math]::Round($diff, 2)) innerhalb dynamischer Toleranz (0.01%)"
    }
    if($diff -le $toleranceFixed){
        return "ok - Diff von $([Math]::Round($diff, 2)) innerhalb fixer Toleranz (±0.02)"
    }
    
    return "nicht ok"
}

function Invoke-Check11_RgEP {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    $pmsRgEP = $Dataset.PMS_SLOERG
    $pimRgEP = $Dataset.PIM_RgEP
    $pmsIsZero = ($pmsRgEP -eq '0' -or $pmsRgEP -eq '0.00' -or $pmsRgEP -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimRgEP)
    if($pmsIsZero -and $pimIsEmpty){return "ok"}
    if($pmsRgEP -eq $pimRgEP){return "ok"}
    return "nicht ok"
}

function Invoke-Check12_Tiefpreis {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    
    # NEU v2.12: Wenn SLLVPL = REDVPL → Kein Tiefpreis vorhanden
    $sllvpl = $Dataset.PMS_SLLVPL
    $redvpl = $Dataset.PMS_REDVPL
    if (-not [string]::IsNullOrEmpty($sllvpl) -and 
        -not [string]::IsNullOrEmpty($redvpl) -and 
        $sllvpl -eq $redvpl) {
        # Beide Preise gleich = kein Tiefpreis
        $pimTief = $Dataset.PIM_Tiefpreis
        if ([string]::IsNullOrEmpty($pimTief)) {
            return "ok - Kein Tiefpreis (SLLVPL = REDVPL)"
        }
    }
    
    $pmsTief = $Dataset.PMS_SLOWAH
    $pimTief = $Dataset.PIM_Tiefpreis
    $pmsIsZero = ($pmsTief -eq '0' -or $pmsTief -eq '0.00' -or $pmsTief -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimTief)
    if($pmsIsZero -and $pimIsEmpty){return "ok"}
    if($pmsTief -eq $pimTief){return "ok"}
    
    # Zusätzliche Checks bei Tiefpreis-Differenz
    $pmsErr = $Dataset.PMS_SLLERR
    if($pmsErr -eq "999906"){return "Warnung - Tiefpreis unterschiedlich - anderer Lf priorisiert (999906)"}
    # NEU v2.22: Verkauft-Flag entfernt
    
    # NEU: SAASEL-Check für nicht-priorisierte Lieferanten
    $saasel = $Dataset.PMS_SAASEL
    if($saasel -eq "0"){
        return "Warnung - TP vorhanden - dieser Lieferant ist nicht priorisierter Lf - TP kommt vermutlich von anderem Lf - dort vermutlich in einer B-Kategorie"
    }
    
    return "nicht ok"
}

function Invoke-Check13_LPrioFehlercode {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    
    # NEU v2.14: PMS Fehlercode kommt aus SAAPNT (nur wenn >= 900000)
    $pmsSaapnt = $Dataset.PMS_SAAPNT
    $pmsErrCode = ""
    
    # Wenn SAAPNT >= 900000 → Das ist der Fehlercode
    if (-not [string]::IsNullOrEmpty($pmsSaapnt)) {
        $saapntVal = 0
        if ([long]::TryParse($pmsSaapnt, [ref]$saapntVal)) {
            if ($saapntVal -ge 900000) {
                $pmsErrCode = $pmsSaapnt
            }
        }
    }
    
    # NEU v2.14: PIM Fehlercode aus PIM_Fehlercode (nicht PIM_Error Code)
    $pimErrCode = $Dataset.PIM_Fehlercode
    
    # Beide leer oder gleich "0" → ok
    $pmsIsEmpty = [string]::IsNullOrEmpty($pmsErrCode) -or $pmsErrCode -eq "0"
    $pimIsEmpty = [string]::IsNullOrEmpty($pimErrCode) -or $pimErrCode -eq "0"
    
    if($pmsIsEmpty -and $pimIsEmpty){return "ok"}
    if($pmsErrCode -eq $pimErrCode){return "ok"}
    return "nicht ok"
}

function Invoke-Check13_Extended {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    $basicResult = Invoke-Check13_LPrioFehlercode -Dataset $Dataset
    if($basicResult -eq "ok"){return $basicResult}
    
    # NEU v2.23: Wenn "nicht ok", prüfe ob PIM_Fehlercode = 999914
    $pimFehlercode = $Dataset.PIM_Fehlercode
    if($pimFehlercode -eq "999914"){
        return "Warnung - Titel fehlt (wird im PMS ggf. nicht sauber verarbeitet)"
    }
    
    # NEU v2.14: PMS Fehlercode kommt aus SAAPNT (nur wenn >= 900000)
    $pmsSaapnt = $Dataset.PMS_SAAPNT
    $pmsErrCode = ""
    
    if (-not [string]::IsNullOrEmpty($pmsSaapnt)) {
        $saapntVal = 0
        if ([long]::TryParse($pmsSaapnt, [ref]$saapntVal)) {
            if ($saapntVal -ge 900000) {
                $pmsErrCode = $pmsSaapnt
            }
        }
    }
    
    # Prüfe spezielle Fehlercodes im PMS
    if($pmsErrCode -eq "999914"){return "Warnung - Fehlercode 999914 (Title fehlt)"}
    if($pmsErrCode -eq "999906"){return "Warnung - Fehlercode 999906 (Anderer Lf priorisiert)"}
    
    # NEU v2.20: Verkauft-Flag Warnung entfernt
    return $basicResult
}

function Invoke-Check14_LPrio {
    param([PSCustomObject]$Dataset)
    if($Dataset.PMS_SLLPAS -eq "passive"){return "ok - Status = passive"}
    $pmsLPrio = $Dataset.PMS_SAAPNT
    $pimLPrio = $Dataset.'PIM_L-Prio-Punkte'
    
    # Direkter Vergleich
    if($pmsLPrio -eq $pimLPrio){return "ok"}
    
    # NEU v2.22: Wenn nicht identisch, prüfe ob SAAPNT = Fehlercode
    # Diese Prüfung kommt VOR allen anderen Checks!
    $pimFehlercode = $Dataset.PIM_Fehlercode
    if(-not [string]::IsNullOrEmpty($pmsLPrio) -and $pmsLPrio -eq $pimFehlercode){
        return "ok - L-Prio-Fehlercode ist identisch"
    }
    
    # Werte parsen
    $pmsVal = 0; $pimVal = 0
    $pmsOk = [long]::TryParse($pmsLPrio, [ref]$pmsVal)
    $pimOk = [long]::TryParse($pimLPrio, [ref]$pimVal)
    if(-not $pmsOk -or -not $pimOk){return "nicht ok - ungültige Werte"}
    $lPrioDiff = $pmsVal - $pimVal
    $prioEPDiffRaw = $Dataset.'PrioEP Diff'
    if([string]::IsNullOrEmpty($prioEPDiffRaw) -or $prioEPDiffRaw -eq "ungültige Werte"){
        return "nicht ok"
    }
    $prioEPDiffVal = 0.0
    $parsedOk = [decimal]::TryParse($prioEPDiffRaw, [ref]$prioEPDiffVal)
    if(-not $parsedOk){
        return "nicht ok"
    }
    
    # Toleranz-Check - Differenz darf BIS ZU 250x PrioEP Diff sein
    $maxAllowedDiff = [Math]::Abs([Math]::Round($prioEPDiffVal * 250, 0))
    $absLPrioDiff = [Math]::Abs($lPrioDiff)
    
    if($absLPrioDiff -le $maxAllowedDiff){
        return "ok - Differenz im Toleranzbereich"
    }
    
    # NEU v2.23: Wenn immer noch "nicht ok", prüfe ob PIM_Fehlercode = 999914
    if($pimFehlercode -eq "999914"){
        return "Warnung - Titel fehlt (wird im PMS ggf. nicht sauber verarbeitet)"
    }
    
    return "nicht ok"
}

Write-Host "  Check-Funktionen geladen." -ForegroundColor Green

# =====================================================================
# MODUL 6: MAIN LOGIC (v1.9)
# =====================================================================
Write-Host "Lade Hauptlogik..." -ForegroundColor Gray

# Globale Variablen für Zusammenfassung
$scriptSuccessfullyCompleted = $false
$createdOutputFiles = [System.Collections.Generic.List[string]]::new()
$pmsEanCount = 0
$pimEanCount = 0
$script:supplierNameForSummary = ""
$script:foundPimDuplicates = $false
$presenceErrorCount = 0
$statusErrorCount = 0
$categoryErrorCount = 0
$genreErrorCount = 0
$preiscodeErrorCount = 0
$standardVPErrorCount = 0
$fixerVPErrorCount = 0
$releaseDatumErrorCount = 0
$errorCodeErrorCount = 0
$vpErrorCount = 0
$vpWarningCount = 0
$prioEPErrorCount = 0
$rgEPErrorCount = 0
$tiefpreisErrorCount = 0
$tiefpreisWarningCount = 0
$lprioFehlercodeErrorCount = 0
$lprioFehlercodeWarningCount = 0
$lprioErrorCount = 0
$lprioWarningCount = 0
$presenceSoldCount = 0
$statusSoldCount = 0
$categorySoldCount = 0
$genreSoldCount = 0
$preiscodeSoldCount = 0
$standardVPSoldCount = 0
$fixerVPSoldCount = 0
$releaseDatumSoldCount = 0
$errorCodeSoldCount = 0
$vpSoldCount = 0
$prioEPSoldCount = 0
$rgEPSoldCount = 0
$tiefpreisSoldCount = 0
$lprioFehlercodeSoldCount = 0
$lprioSoldCount = 0

function Pause-Ende {
    Write-Host ""
    $ColorOk = "Green"
    $ColorNok = "Red"
    $Line = "=" * 60
    $matchedEanCount = ($All_Datasets | Where-Object { $_.'Gefunden ...' -eq 'im PMS und im PIM' }).Count
    $errorCount = $Error_Datasets.Count
    $successCount = ($All_Datasets.Count - $errorCount)
    $supplierDisplayString = $pmsSupplier
    if ($script:supplierNameForSummary -ne $pmsSupplier) { $supplierDisplayString = "$($script:supplierNameForSummary) ($($pmsSupplier))" }
    
    Write-Host $Line -ForegroundColor Cyan
    Write-Host "Vergleich Phase X - Berechnung (Script Version $($global:ScriptVersion))" -ForegroundColor White
    Write-Host ""
    Write-Host "Input-Files:" -ForegroundColor Yellow
    Write-Host "  PMS: $(Split-Path $pmsFilePath -Leaf)"
    Write-Host "  PIM: $(Split-Path $pimFilePath -Leaf)"
    Write-Host ""
    Write-Host "Output-Files (in '.\VergleichsErgebnisseBerechnung\$script:sanitizedSupplierName\'):" -ForegroundColor Yellow
    if ($createdOutputFiles.Count -gt 0) { 
        foreach ($file in $createdOutputFiles) { Write-Host "  $file" } 
    } else { 
        Write-Host "  Es wurden keine Output-Files erstellt." 
    }
    Write-Host ""
    $runDateTime = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    Write-Host "Datum: $runDateTime" -ForegroundColor Yellow
    Write-Host "Dauer: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Zusammenfassung:" -ForegroundColor Yellow
    Write-Host "  Header-Überprüfung: OK" -ForegroundColor $ColorOk
    Write-Host "  Überprüfung der Lieferanten-Nummern: OK" -ForegroundColor $ColorOk
    Write-Host "  Anzahl EANs im PMS-File: $pmsEanCount"
    Write-Host "  Anzahl EANs im PIM-File: $pimEanCount"
    Write-Host "  Anzahl EANs in beiden Files: $matchedEanCount"
    Write-Host "  Anzahl fehlerfreie EANs: $successCount" -ForegroundColor $ColorOk
    
    if ($warningOnlyCount -gt 0) {
        Write-Host "  Anzahl EANs mit Warnungen (nur): $warningOnlyCount" -ForegroundColor Yellow
    } else {
        Write-Host "  Anzahl EANs mit Warnungen (nur): 0" -ForegroundColor $ColorOk
    }
    
    $finalStatusColor = $ColorOk
    if ($errorCount -gt 0) {
        Write-Host "  Anzahl EANs mit Fehlern: $errorCount" -ForegroundColor $ColorNok
        $finalStatusColor = $ColorNok
        $finalStatusText = "Nicht ok - hat Fehler"
    } else {
        Write-Host "  Anzahl EANs mit Fehlern: 0" -ForegroundColor $ColorOk
        $finalStatusText = "OK - fehlerfrei"
    }
    
    if ($script:foundPimDuplicates) { 
        Write-Host "  Doppelte EANs im PIM File gefunden" -ForegroundColor Red 
    } else { 
        Write-Host "  Keine doppelten EANs im PIM File gefunden" -ForegroundColor Green 
    }
    
    Write-Host ""
    Write-Host "Berechnung von Lieferant $supplierDisplayString ist $finalStatusText" -ForegroundColor $finalStatusColor
    
    if ($errorCount -gt 0 -or $totalWarningCount -gt 0) {
        Write-Host ""
        Write-Host "Fehler-Übersicht:" -ForegroundColor Yellow
        Write-Host ""
        
        $headerFormat = "{0,-10} {1,-30} {2,15} {3,18} {4,16}"
        $separatorLine = "-" * 90
        
        Write-Host ($headerFormat -f "Check", "Titel", "Anzahl Fehler", "Anzahl Warnungen", "Fehler+Verkauft") -ForegroundColor Cyan
        Write-Host $separatorLine -ForegroundColor Cyan
        
        $checks = @(
            @{Num = "Check 0"; Titel = "Vorhanden in beiden Quellen"; Fehler = $presenceErrorCount; Warnung = 0; Sold = $presenceSoldCount }
            @{Num = "Check 1"; Titel = "Status"; Fehler = $statusErrorCount; Warnung = 0; Sold = $statusSoldCount }
            @{Num = "Check 2"; Titel = "Kategorie"; Fehler = $categoryErrorCount; Warnung = 0; Sold = $categorySoldCount }
            @{Num = "Check 3"; Titel = "Genre"; Fehler = $genreErrorCount; Warnung = 0; Sold = $genreSoldCount }
            @{Num = "Check 4"; Titel = "Preiscode"; Fehler = $preiscodeErrorCount; Warnung = 0; Sold = $preiscodeSoldCount }
            @{Num = "Check 5"; Titel = "Standard VP"; Fehler = $standardVPErrorCount; Warnung = 0; Sold = $standardVPSoldCount }
            @{Num = "Check 6"; Titel = "Fixer VP"; Fehler = $fixerVPErrorCount; Warnung = 0; Sold = $fixerVPSoldCount }
            @{Num = "Check 7"; Titel = "Release-Datum"; Fehler = $releaseDatumErrorCount; Warnung = 0; Sold = $releaseDatumSoldCount }
            @{Num = "Check 8"; Titel = "Errorcode"; Fehler = $errorCodeErrorCount; Warnung = 0; Sold = $errorCodeSoldCount }
            @{Num = "Check 9"; Titel = "VP"; Fehler = $vpErrorCount; Warnung = $vpWarningCount; Sold = $vpSoldCount }
            @{Num = "Check 10"; Titel = "PrioEP"; Fehler = $prioEPErrorCount; Warnung = 0; Sold = $prioEPSoldCount }
            @{Num = "Check 11"; Titel = "RgEP"; Fehler = $rgEPErrorCount; Warnung = 0; Sold = $rgEPSoldCount }
            @{Num = "Check 12"; Titel = "Tiefpreis"; Fehler = $tiefpreisErrorCount; Warnung = $tiefpreisWarningCount; Sold = $tiefpreisSoldCount }
            @{Num = "Check 13"; Titel = "L-Prio Fehlercode"; Fehler = $lprioFehlercodeErrorCount; Warnung = $lprioFehlercodeWarningCount; Sold = $lprioFehlercodeSoldCount }
            @{Num = "Check 14"; Titel = "L-Prio"; Fehler = $lprioErrorCount; Warnung = $lprioWarningCount; Sold = $lprioSoldCount }
        )
        
        foreach ($check in $checks) {
            $warnungAnzeige = if ($check.Warnung -gt 0) { $check.Warnung } else { "-" }
            $soldAnzeige = if ($check.Sold -gt 0) { $check.Sold } else { "-" }
            $line = $headerFormat -f $check.Num, $check.Titel, $check.Fehler, $warnungAnzeige, $soldAnzeige
            
            if ($check.Fehler -gt 0) {
                Write-Host $line -ForegroundColor Red
            } elseif ($check.Warnung -gt 0) {
                Write-Host $line -ForegroundColor Yellow
            } else {
                Write-Host $line -ForegroundColor Green
            }
        }
        
        Write-Host $separatorLine -ForegroundColor Cyan
    }
    
    [void](Read-Host "Drücke ENTER um das Fenster zu schliessen")
}

function Invoke-MainLogic {
    try {
        # Fenster vergrössern
        try {
            if (-not $psISE) {
                $cur = $Host.UI.RawUI.WindowSize
                $newH = [int]($cur.Height * 1.5)
                if ($Host.UI.RawUI.BufferSize.Height -lt $newH) {
                    $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($Host.UI.RawUI.BufferSize.Width, $newH)
                }
                $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size($cur.Width, $newH)
            }
        } catch { }

        $scriptDir = $PSScriptRoot
        if ([string]::IsNullOrEmpty($scriptDir)) {
            $scriptDir = Split-Path -Parent $script:MyInvocation.MyCommand.Path
        }
        if ([string]::IsNullOrEmpty($scriptDir)) {
            $scriptDir = (Get-Location).Path
        }
        
        $parentDir = Split-Path -Parent $scriptDir
        $rootDir = Split-Path -Parent $parentDir
        
        $pimDir = Join-Path $rootDir "PIM"
        $script:InputDirectory = Join-Path $pimDir "PhaseX_Berechnung"
        
        Write-Host "    Script-Verzeichnis: $scriptDir" -ForegroundColor Gray
        Write-Host "    Input-Verzeichnis:  $($script:InputDirectory)" -ForegroundColor Gray

        Write-Host "--- Skript-Version $($global:ScriptVersion) ---`n" -ForegroundColor Gray
        $script:stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        Write-Host "1. Prüfe Eingabe-Verzeichnis..."
        if (-not (Test-Path $script:InputDirectory -PathType Container)) { 
            throw "Eingabeverzeichnis existiert nicht: '$($script:InputDirectory)'`n`nErwartet wird der Ordner 'PIM\PhaseX_Berechnung' im übergeordneten Verzeichnis des Script-Ordners."
        }
        Write-Host "    Verzeichnis ist vorhanden."

        Write-Host "2. Bitte Dateien auswählen..."
        $absInput = Convert-Path $script:InputDirectory
        $script:pmsFilePath = Get-FilePathDialog -WindowTitle "Bitte die PMS-Datei auswählen" -InitialDirectory $absInput
        if (-not $script:pmsFilePath) { Write-Host "Aktion abgebrochen."; return }
        $script:pimFilePath = Get-FilePathDialog -WindowTitle "Bitte die PIM-Datei auswählen" -InitialDirectory $absInput
        if (-not $script:pimFilePath) { Write-Host "Aktion abgebrochen."; return }
        Write-Host "    PMS-Datei: $(Split-Path $script:pmsFilePath -Leaf)"
        Write-Host "    PIM-Datei: $(Split-Path $script:pimFilePath -Leaf)"

        Write-Host "3. Prüfe Header der CSV-Dateien..."
        Write-Host "    - Prüfe PMS-Datei..."
        $pmsHeaderLine = (Get-Content -Path $script:pmsFilePath -TotalCount 1).TrimEnd(';')
        if ([string]::IsNullOrWhiteSpace($pmsHeaderLine)) { throw "PMS-Datei '$($script:pmsFilePath)' ist leer oder Header fehlt." }
        $pmsActualHeader = $pmsHeaderLine.Split(';')
        if ($null -ne (Compare-Object $script:PMS_Header_Expected $pmsActualHeader -CaseSensitive)) {
            throw "Header PMS nicht korrekt.`nErwartet: $($script:PMS_Header_Expected -join ';')`nGefunden: $($pmsActualHeader -join ';')"
        }
        Write-Host "      -> Header in PMS-Datei ist korrekt." -ForegroundColor Green

        Write-Host "    - Prüfe PIM-Datei..."
        $pimHeaderLine = Get-Content -Path $script:pimFilePath -TotalCount 1 -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($pimHeaderLine)) { throw "PIM-Datei '$($script:pimFilePath)' ist leer oder Header fehlt." }
        $pimActualHeader = ($pimHeaderLine.Replace('"', '')).Split(';')
        if ($null -ne (Compare-Object $script:PIM_Header_Expected $pimActualHeader -CaseSensitive)) {
            throw "Header PIM nicht korrekt.`nErwartet: $($script:PIM_Header_Expected -join ';')`nGefunden: $($pimActualHeader -join ';')"
        }
        Write-Host "      -> Header in PIM-Datei ist korrekt." -ForegroundColor Green

        Write-Host "4. Führe Lieferanten-Check durch..."
        $pmsFirstDataRow = (Get-Content -Path $script:pmsFilePath -TotalCount 2 | Select-Object -Last 1).TrimEnd(';')
        $pmsFirstRecord = $pmsFirstDataRow | ConvertFrom-Csv -Header $script:PMS_Header_Expected -Delimiter ';'
        $script:pmsSupplier = $pmsFirstRecord.SLLLFN

        $pimFirstDataRow = (Get-Content -Path $script:pimFilePath -TotalCount 2 -Encoding UTF8 | Select-Object -Last 1)
        $pimFirstRecord = $pimFirstDataRow | ConvertFrom-Csv -Header $script:PIM_Header_Expected -Delimiter ';'
        $pimSupplier = $pimFirstRecord.Lieferant

        if ($script:pmsSupplier -ne $pimSupplier) {
            throw "Lieferantennummern stimmen NICHT überein!`nPMS: '$($script:pmsSupplier)'`nPIM: '$pimSupplier'"
        }

        Write-Host "    Lieferantennummern stimmen überein: '$($script:pmsSupplier)'." -ForegroundColor Green
        $supplierName = $script:pmsSupplier
        if ($script:SupplierLookupTable.ContainsKey($script:pmsSupplier)) { $supplierName = $script:SupplierLookupTable[$script:pmsSupplier] }
        $script:supplierNameForSummary = $supplierName
        $script:sanitizedSupplierName = $supplierName.Replace(' ', '-').Replace('+', '') -replace '[\\/:*?"<>|]', ''

        if ($script:SaveToSharePoint) {
            $script:OutputDirectory = ".\VergleichsErgebnisseBerechnung"
            if (-not (Test-Path $script:OutputDirectory -PathType Container)) { New-Item -Path $script:OutputDirectory -ItemType Directory | Out-Null }
            $script:OutputDirectory = Join-Path $script:OutputDirectory $script:sanitizedSupplierName
            if (-not (Test-Path $script:OutputDirectory -PathType Container)) {
                Write-Host "    - Erstelle Lieferanten-Unterverzeichnis: $($script:OutputDirectory)" -ForegroundColor Gray
                New-Item -Path $script:OutputDirectory -ItemType Directory -ErrorAction Stop | Out-Null
            }
        } else {
            $script:OutputDirectory = Split-Path -Path $script:pmsFilePath -Parent
            if (-not (Test-Path $script:OutputDirectory -PathType Container)) { throw "Output-Ordner (lokal) existiert nicht: '$($script:OutputDirectory)'" }
        }

        $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $SystemUserName = $env:USERNAME
        $FriendlyUserName = $SystemUserName
        if ($script:UserLookupTable.ContainsKey($SystemUserName)) { $FriendlyUserName = $script:UserLookupTable[$SystemUserName] }

        $OutputFileName_All = "PhaseX_Vergl_Berechnung__$($script:sanitizedSupplierName)_$($script:pmsSupplier)__$($FriendlyUserName)__ALLE__$($Timestamp).csv"
        $OutputFileName_Errors = "PhaseX_Vergl_Berechnung__$($script:sanitizedSupplierName)_$($script:pmsSupplier)__$($FriendlyUserName)__ERRORS__$($Timestamp).csv"
        $OutputFilePath_All = Join-Path $script:OutputDirectory $OutputFileName_All
        $OutputFilePath_Errors = Join-Path $script:OutputDirectory $OutputFileName_Errors

        Write-Host "5. Lese und verarbeite Dateien... (Dies kann einige Minuten dauern)"
        $All_Datasets_Hashtable = @{}
        $pmsSkippedCounter = 0
        $pimSkippedCounter = 0

        Write-Host "    - Verarbeite PMS-Datei..."
        $reader = $null
        try {
            $reader = [System.IO.File]::OpenText($script:pmsFilePath)
            $null = $reader.ReadLine()
            while ($reader.Peek() -ge 0) {
                $line = $reader.ReadLine()
                $values = $line.Split(';')
                $pmsRowProps = [ordered]@{}
                for ($i = 0; $i -lt $script:PMS_Header_Expected.Count; $i++) { $pmsRowProps[$script:PMS_Header_Expected[$i]] = $values[$i] }
                $pmsRow = [PSCustomObject]$pmsRowProps
                $ean = $pmsRow.SLLEAN
                if (-not $ean) { $pmsSkippedCounter++; continue }
                if ($All_Datasets_Hashtable.ContainsKey($ean)) {
                    Write-Warning "Doppelte EAN '$ean' in PMS-Datei. Nur erster Eintrag wird berücksichtigt."
                    continue
                }
                $newObjectProps = [ordered]@{
                    EAN                                    = "'$ean"
                    'Gefunden ...'                         = "nur im PMS"
                    'Check Summary'                        = ""
                    'Check 0: Vorhanden in beiden Quellen' = ""
                    'Check 1: Status'                      = ""
                    'Check 2: Kategorie'                   = ""
                    'Check 3: Genre'                       = ""
                    'Check 4: Preiscode'                   = ""
                    'Check 5: Standard VP ab Lieferant'    = ""
                    'Check 6: Fixer VP'                    = ""
                    'Check 7: Release-Datum'               = ""
                    'Check 8: Errorcode'                   = ""
                    'Check 9: VP'                          = ""
                    'VP Diff'                              = ""
                    'Check 10: PrioEP'                     = ""
                    'PrioEP Diff'                          = ""
                    'Check 11: RgEP'                       = ""
                    'RgEP Diff'                            = ""
                    'Check 12: Tiefpreis'                  = ""
                    'Check 13: L-Prio Fehlercode'          = ""
                    'Check 14: L-Prio'                     = ""
                    'L-Prio Diff'                          = ""
                    'ZeitDiff letzte Änderung'             = ""
                    'ZeitDiff Bewertung'                   = ""
                }
                foreach ($header in $script:PMS_Header_Expected) { if ($header -ne 'SLLLFN') { $newObjectProps["PMS_$header"] = $pmsRow.$header } }
                foreach ($header in $script:PIM_Header_Expected) { $newObjectProps["PIM_$header"] = $null }
                $All_Datasets_Hashtable.Add($ean, [PSCustomObject]$newObjectProps)
            }
        } finally { if ($reader) { $reader.Close(); $reader.Dispose() } }
        $script:pmsEanCount = $All_Datasets_Hashtable.Count
        Write-Host "    - PMS-Datei eingelesen. $($All_Datasets_Hashtable.Count) eindeutige Datensätze gefunden."
        if ($pmsSkippedCounter -gt 0) { Write-Warning "$pmsSkippedCounter Zeilen ohne EAN im PMS-File wurden übersprungen." }

        Write-Host "    - Verarbeite PIM-Datei..."
        $pimSeenEans = @{}
        $reader = $null
        try {
            $reader = New-Object System.IO.StreamReader($script:pimFilePath, [System.Text.Encoding]::UTF8)
            $null = $reader.ReadLine()
            while ($reader.Peek() -ge 0) {
                $line = $reader.ReadLine()
                $values = ($line.Replace('"', '')).Split(';')
                $pimRowProps = [ordered]@{}
                
                for ($i = 0; $i -lt $script:PIM_Header_Expected.Count; $i++) {
                    $value = $values[$i]
                    $fieldName = $script:PIM_Header_Expected[$i]
                    
                    if ($fieldName -in @('Fixer VP', 'Acquisition Price', 'Sales Price', 'VP', 'PrioEP', 'RgEP', 'Tiefpreis')) {
                        $value = $value.Replace(',', '')
                    }
                    
                    $pimRowProps[$fieldName] = $value
                }
                
                $pimRow = [PSCustomObject]$pimRowProps
                $ean = $pimRow.EAN
                if (-not $ean) { $pimSkippedCounter++; continue }
                $script:pimEanCount++
                if ($pimSeenEans.ContainsKey($ean)) {
                    if ($All_Datasets_Hashtable.ContainsKey($ean)) {
                        $existing = $All_Datasets_Hashtable[$ean]
                        $existing.'Gefunden ...' = "mehrfach im PIM"
                        $existing.'Check Summary' = "nicht ok - EAN mehrfach im PIM"
                        $script:foundPimDuplicates = $true
                    }
                    continue
                } else { $pimSeenEans.Add($ean, $true) }

                if ($All_Datasets_Hashtable.ContainsKey($ean)) {
                    $existing = $All_Datasets_Hashtable[$ean]
                    $existing.'Gefunden ...' = "im PMS und im PIM"
                    foreach ($header in $script:PIM_Header_Expected) { $existing."PIM_$header" = $pimRow.$header }
                } else {
                    $newObjectProps = [ordered]@{
                        EAN                                    = "'$ean"
                        'Gefunden ...'                         = "nur im PIM"
                        'Check Summary'                        = ""
                        'Check 0: Vorhanden in beiden Quellen' = ""
                        'Check 1: Status'                      = ""
                        'Check 2: Kategorie'                   = ""
                        'Check 3: Genre'                       = ""
                        'Check 4: Preiscode'                   = ""
                        'Check 5: Standard VP ab Lieferant'    = ""
                        'Check 6: Fixer VP'                    = ""
                        'Check 7: Release-Datum'               = ""
                        'Check 8: Errorcode'                   = ""
                        'Check 9: VP'                          = ""
                        'VP Diff'                              = ""
                        'Check 10: PrioEP'                     = ""
                        'PrioEP Diff'                          = ""
                        'Check 11: RgEP'                       = ""
                        'RgEP Diff'                            = ""
                        'Check 12: Tiefpreis'                  = ""
                        'Check 13: L-Prio Fehlercode'          = ""
                        'Check 14: L-Prio'                     = ""
                        'L-Prio Diff'                          = ""
                        'ZeitDiff letzte Änderung'             = ""
                        'ZeitDiff Bewertung'                   = ""
                    }
                    foreach ($header in $script:PMS_Header_Expected) { if ($header -ne 'SLLLFN') { $newObjectProps["PMS_$header"] = $null } }
                    foreach ($header in $script:PIM_Header_Expected) { $newObjectProps["PIM_$header"] = $pimRow.$header }
                    $All_Datasets_Hashtable.Add($ean, [PSCustomObject]$newObjectProps)
                }
            }
        } finally { if ($reader) { $reader.Close(); $reader.Dispose() } }
        Write-Host "    - PIM-Datei verarbeitet."
        if ($pimSkippedCounter -gt 0) { Write-Warning "$pimSkippedCounter Zeilen ohne EAN im PIM-File wurden übersprungen." }
        Write-Host "Beide Files eingelesen. Dauer $($script:stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
        Write-Host "Gesamtanzahl eindeutiger Datensätze: $($All_Datasets_Hashtable.Count)"

        $script:All_Datasets = [System.Collections.ArrayList]@($All_Datasets_Hashtable.Values)
        $totalDatasets = $script:All_Datasets.Count
        $i = 0
        Write-Host "6. Führe Checks durch..."

        foreach ($dataset in $script:All_Datasets) {
            $i++
            if ($i % 5000 -eq 0) {
                $perc = [Math]::Floor(($i / $totalDatasets) * 100)
                Write-Progress -Activity "Schritt 6: Führe Checks durch" -Status "$perc% abgeschlossen ($i von $totalDatasets EANs)" -PercentComplete $perc
            }

            switch ($dataset.'Gefunden ...') {
                'im PMS und im PIM' { $dataset.'Check 0: Vorhanden in beiden Quellen' = 'ok - EAN in beiden Quellen' }
                'nur im PIM' { $dataset.'Check 0: Vorhanden in beiden Quellen' = 'ok - EAN nur im PIM' }
                'nur im PMS' {
                    if ($dataset.PMS_SLLPAS -eq 'passive') { $dataset.'Check 0: Vorhanden in beiden Quellen' = 'ok - EAN fehlt im PIM - passive im PMS' }
                    else { $dataset.'Check 0: Vorhanden in beiden Quellen' = 'nicht ok - EAN fehlt im PIM' }
                }
                'mehrfach im PIM' { $dataset.'Check 0: Vorhanden in beiden Quellen' = 'nicht ok - EAN mehrfach im PIM' }
                default { $dataset.'Check 0: Vorhanden in beiden Quellen' = 'nicht ok' }
            }

            if ($dataset.'Check Summary' -like 'nicht ok - EAN mehrfach im PIM') { continue }

            if ($dataset.'Gefunden ...' -eq "im PMS und im PIM") {
                $dataset.'ZeitDiff letzte Änderung' = Invoke-CalculateTimeDifference -Dataset $dataset
                $dataset.'Check 1: Status' = Invoke-Check1_Status -Dataset $dataset

                if ($dataset.'Check 1: Status' -like 'ok*') {
                    $dataset.'Check 2: Kategorie' = Invoke-Check2_Kategorie -Dataset $dataset
                    
                    if ($dataset.'Check 2: Kategorie' -eq 'ok - Kein Kat-Mapping im PMS und PIM') {
                        $dataset.'Check 3: Genre' = ''
                        $dataset.'Check 4: Preiscode' = ''
                        $dataset.'Check 5: Standard VP ab Lieferant' = ''
                        $dataset.'Check 6: Fixer VP' = ''
                        $dataset.'Check 7: Release-Datum' = ''
                        $dataset.'Check 8: Errorcode' = ''
                        $dataset.'Check 9: VP' = ''
                        $dataset.'VP Diff' = ''
                        $dataset.'Check 10: PrioEP' = ''
                        $dataset.'PrioEP Diff' = ''
                        $dataset.'Check 11: RgEP' = ''
                        $dataset.'RgEP Diff' = ''
                        $dataset.'Check 12: Tiefpreis' = ''
                        $dataset.'Check 13: L-Prio Fehlercode' = ''
                        $dataset.'Check 14: L-Prio' = ''
                        $dataset.'L-Prio Diff' = ''
                        $dataset.'Check Summary' = 'ok'
                        continue
                    }
                    
                    $dataset.'Check 3: Genre' = Invoke-Check3_Genre -Dataset $dataset
                    $dataset.'Check 4: Preiscode' = Invoke-Check4_Preiscode -Dataset $dataset

                    $whitelistSkipMessage = "ok - Lieferant bei Kategorie nicht auf Whitelist"
                    if ($dataset.PMS_SAAPNT -eq "999905") {
                        $dataset.'Check 5: Standard VP ab Lieferant' = $whitelistSkipMessage
                        $dataset.'Check 6: Fixer VP' = $whitelistSkipMessage
                        $dataset.'Check 7: Release-Datum' = $whitelistSkipMessage
                        $dataset.'Check 8: Errorcode' = $whitelistSkipMessage
                        $dataset.'Check 9: VP' = $whitelistSkipMessage
                        $dataset.'Check 10: PrioEP' = $whitelistSkipMessage
                        $dataset.'Check 11: RgEP' = $whitelistSkipMessage
                        $dataset.'Check 12: Tiefpreis' = $whitelistSkipMessage
                        $dataset.'Check 14: L-Prio' = $whitelistSkipMessage
                    } else {
                        $dataset.'Check 5: Standard VP ab Lieferant' = Invoke-Check5_StandardVP -Dataset $dataset
                        $dataset.'Check 6: Fixer VP' = Invoke-Check6_FixerVP -Dataset $dataset
                        $dataset.'Check 7: Release-Datum' = Invoke-Check7_ReleaseDatum -Dataset $dataset
                        $dataset.'Check 8: Errorcode' = Invoke-Check8_Errorcode -Dataset $dataset
                        $dataset.'Check 9: VP' = Invoke-Check9_VP -Dataset $dataset

                        if ($dataset.'Check 9: VP' -eq 'nicht ok') {
                            $pmsVal = 0.0; $pimVal = 0.0
                            $pmsOk = [decimal]::TryParse($dataset.PMS_SLLVPL, [ref]$pmsVal)
                            $pimOk = [decimal]::TryParse($dataset.PIM_VP, [ref]$pimVal)
                            $dataset.'VP Diff' = if ($pmsOk -and $pimOk) { $pmsVal - $pimVal } else { "ungültige Werte" }
                        }

                        $dataset.'Check 10: PrioEP' = Invoke-Check10_PrioEP -Dataset $dataset
                        $pmsVal = 0.0; $pimVal = 0.0
                        $pmsOk = [decimal]::TryParse($dataset.PMS_SLLEPL, [ref]$pmsVal)
                        $pimOk = [decimal]::TryParse($dataset.PIM_PrioEP, [ref]$pimVal)
                        if ($pmsOk -and $pimOk -and $pmsVal -ne $pimVal) {
                            $dataset.'PrioEP Diff' = $pmsVal - $pimVal
                        }

                        $dataset.'Check 11: RgEP' = Invoke-Check11_RgEP -Dataset $dataset
                        if ($dataset.'Check 11: RgEP' -eq 'nicht ok') {
                            $pmsVal = 0.0; $pimVal = 0.0
                            $pmsOk = [decimal]::TryParse($dataset.PMS_SLOERG, [ref]$pmsVal)
                            $pimOk = [decimal]::TryParse($dataset.PIM_RgEP, [ref]$pimVal)
                            $dataset.'RgEP Diff' = if ($pmsOk -and $pimOk) { $pmsVal - $pimVal } else { "ungültige Werte" }
                        }

                        $dataset.'Check 12: Tiefpreis' = Invoke-Check12_Tiefpreis -Dataset $dataset
                        $dataset.'Check 13: L-Prio Fehlercode' = Invoke-Check13_LPrioFehlercode -Dataset $dataset
                        $dataset.'Check 13: L-Prio Fehlercode' = Invoke-Check13_Extended -Dataset $dataset

                        $dataset.'Check 14: L-Prio' = Invoke-Check14_LPrio -Dataset $dataset
                        if ($dataset.'Check 14: L-Prio' -eq 'nicht ok') {
                            $pmsVal = 0; $pimVal = 0
                            $pmsOk = [long]::TryParse($dataset.PMS_SAAPNT, [ref]$pmsVal)
                            $pimOk = [long]::TryParse($dataset.'PIM_L-Prio-Punkte', [ref]$pimVal)
                            $dataset.'L-Prio Diff' = if ($pmsOk -and $pimOk) { $pmsVal - $pimVal } else { "ungültige Werte" }
                        }
                    }

                    if (($dataset.'Check 1: Status' -like 'ok*') -and 
                        ($dataset.'Check 2: Kategorie' -like 'ok*') -and 
                        ($dataset.'Check 3: Genre' -like 'ok*') -and
                        ($dataset.'Check 4: Preiscode' -like 'ok*') -and
                        ($dataset.'Check 5: Standard VP ab Lieferant' -like 'ok*') -and
                        ($dataset.'Check 6: Fixer VP' -like 'ok*') -and
                        ($dataset.'Check 7: Release-Datum' -like 'ok*') -and
                        ($dataset.'Check 8: Errorcode' -like 'ok*') -and
                        ($dataset.'Check 9: VP' -like 'ok*' -or $dataset.'Check 9: VP' -like 'Warnung*') -and
                        ($dataset.'Check 10: PrioEP' -like 'ok*') -and
                        ($dataset.'Check 11: RgEP' -like 'ok*') -and
                        ($dataset.'Check 12: Tiefpreis' -like 'ok*' -or $dataset.'Check 12: Tiefpreis' -like 'Warnung*') -and
                        ($dataset.'Check 13: L-Prio Fehlercode' -like 'ok*' -or $dataset.'Check 13: L-Prio Fehlercode' -like 'Warnung*') -and
                        ($dataset.'Check 14: L-Prio' -like 'ok*' -or $dataset.'Check 14: L-Prio' -like 'Warnung*')) {
                        $dataset.'Check Summary' = 'ok'
                    } else {
                        $dataset.'Check Summary' = 'nicht ok'
                    }
                } else {
                    $dataset.'Check 2: Kategorie' = '---'
                    $dataset.'Check 3: Genre' = '---'
                    $dataset.'Check 4: Preiscode' = '---'
                    $dataset.'Check 5: Standard VP ab Lieferant' = '---'
                    $dataset.'Check 6: Fixer VP' = '---'
                    $dataset.'Check 7: Release-Datum' = '---'
                    $dataset.'Check 8: Errorcode' = '---'
                    $dataset.'Check 9: VP' = '---'
                    $dataset.'VP Diff' = '---'
                    $dataset.'Check 10: PrioEP' = '---'
                    $dataset.'PrioEP Diff' = '---'
                    $dataset.'Check 11: RgEP' = '---'
                    $dataset.'RgEP Diff' = '---'
                    $dataset.'Check 12: Tiefpreis' = '---'
                    $dataset.'Check 13: L-Prio Fehlercode' = '---'
                    $dataset.'Check 14: L-Prio' = '---'
                    $dataset.'L-Prio Diff' = '---'
                }
            } elseif ($dataset.'Gefunden ...' -eq "nur im PIM") {
                $dataset.'Check Summary' = 'ok - EAN nur im PIM'
            } else {
                if ($dataset.PMS_SLLPAS -eq 'passive') { $dataset.'Check Summary' = 'ok - EAN fehlt im PIM - passive im PMS' }
                else { $dataset.'Check Summary' = 'nicht ok - EAN fehlt im PIM' }
            }
        }
        Write-Progress -Activity "Schritt 6: Führe Checks durch" -Completed
        Write-Host "    Checks abgeschlossen." -ForegroundColor Green

        Write-Host "7. Bereite Export vor (OPTIMIERT: Single-Pass)..."
        $totalRowCount = $script:All_Datasets.Count
        
        # Initialisiere alle Counters
        $script:vpWarningCount = 0
        $script:tiefpreisWarningCount = 0
        $script:lprioFehlercodeWarningCount = 0
        $script:lprioWarningCount = 0
        $script:totalWarningCount = 0
        $script:warningOnlyCount = 0
        
        $script:presenceErrorCount = 0
        $script:statusErrorCount = 0
        $script:categoryErrorCount = 0
        $script:genreErrorCount = 0
        $script:preiscodeErrorCount = 0
        $script:standardVPErrorCount = 0
        $script:fixerVPErrorCount = 0
        $script:releaseDatumErrorCount = 0
        $script:errorCodeErrorCount = 0
        $script:vpErrorCount = 0
        $script:prioEPErrorCount = 0
        $script:rgEPErrorCount = 0
        $script:tiefpreisErrorCount = 0
        $script:lprioFehlercodeErrorCount = 0
        $script:lprioErrorCount = 0
        
        $script:presenceSoldCount = 0
        $script:statusSoldCount = 0
        $script:categorySoldCount = 0
        $script:genreSoldCount = 0
        $script:preiscodeSoldCount = 0
        $script:standardVPSoldCount = 0
        $script:fixerVPSoldCount = 0
        $script:releaseDatumSoldCount = 0
        $script:errorCodeSoldCount = 0
        $script:vpSoldCount = 0
        $script:prioEPSoldCount = 0
        $script:rgEPSoldCount = 0
        $script:tiefpreisSoldCount = 0
        $script:lprioFehlercodeSoldCount = 0
        $script:lprioSoldCount = 0
        
        # ArrayLists für gefilterte Datasets (schneller als Array += )
        $script:Error_Datasets = [System.Collections.ArrayList]::new()
        $ErrorsAndWarnings_List = [System.Collections.ArrayList]::new()
        
        # SINGLE-PASS COUNTING: Ein Durchlauf durch alle Daten
        foreach ($dataset in $script:All_Datasets) {
            $checkSummary = $dataset.'Check Summary'
            $isError = ($checkSummary -notlike 'ok*')
            $isSold = ($dataset.PMS_FLGVKF -eq '1')
            
            # Check für Warnungen
            $hasVpWarning = ($dataset.'Check 9: VP' -like 'Warnung*')
            $hasTiefpreisWarning = ($dataset.'Check 12: Tiefpreis' -like 'Warnung*')
            $hasLprioFehlercodeWarning = ($dataset.'Check 13: L-Prio Fehlercode' -like 'Warnung*')
            $hasLprioWarning = ($dataset.'Check 14: L-Prio' -like 'Warnung*')
            $hasAnyWarning = ($hasVpWarning -or $hasTiefpreisWarning -or $hasLprioFehlercodeWarning -or $hasLprioWarning)
            
            # Warnungs-Counts
            if ($hasVpWarning) { $script:vpWarningCount++ }
            if ($hasTiefpreisWarning) { $script:tiefpreisWarningCount++ }
            if ($hasLprioFehlercodeWarning) { $script:lprioFehlercodeWarningCount++ }
            if ($hasLprioWarning) { $script:lprioWarningCount++ }
            
            # Nur-Warnungen (ok aber mit Warnung)
            if (-not $isError -and $hasAnyWarning) {
                $script:warningOnlyCount++
            }
            
            # Fehler-Datasets und Error-Counts
            if ($isError) {
                [void]$script:Error_Datasets.Add($dataset)
                
                # Error-Counts für jeden Check
                if ($dataset.'Check 0: Vorhanden in beiden Quellen' -like 'nicht ok*') {
                    $script:presenceErrorCount++
                    if ($isSold) { $script:presenceSoldCount++ }
                }
                if ($dataset.'Check 1: Status' -like 'nicht ok*') {
                    $script:statusErrorCount++
                    if ($isSold) { $script:statusSoldCount++ }
                }
                if ($dataset.'Check 2: Kategorie' -like 'nicht ok*') {
                    $script:categoryErrorCount++
                    if ($isSold) { $script:categorySoldCount++ }
                }
                if ($dataset.'Check 3: Genre' -like 'nicht ok*') {
                    $script:genreErrorCount++
                    if ($isSold) { $script:genreSoldCount++ }
                }
                if ($dataset.'Check 4: Preiscode' -like 'nicht ok*') {
                    $script:preiscodeErrorCount++
                    if ($isSold) { $script:preiscodeSoldCount++ }
                }
                if ($dataset.'Check 5: Standard VP ab Lieferant' -like 'nicht ok*') {
                    $script:standardVPErrorCount++
                    if ($isSold) { $script:standardVPSoldCount++ }
                }
                if ($dataset.'Check 6: Fixer VP' -like 'nicht ok*') {
                    $script:fixerVPErrorCount++
                    if ($isSold) { $script:fixerVPSoldCount++ }
                }
                if ($dataset.'Check 7: Release-Datum' -like 'nicht ok*') {
                    $script:releaseDatumErrorCount++
                    if ($isSold) { $script:releaseDatumSoldCount++ }
                }
                if ($dataset.'Check 8: Errorcode' -like 'nicht ok*') {
                    $script:errorCodeErrorCount++
                    if ($isSold) { $script:errorCodeSoldCount++ }
                }
                if ($dataset.'Check 9: VP' -like 'nicht ok*') {
                    $script:vpErrorCount++
                    if ($isSold) { $script:vpSoldCount++ }
                }
                if ($dataset.'Check 10: PrioEP' -like 'nicht ok*') {
                    $script:prioEPErrorCount++
                    if ($isSold) { $script:prioEPSoldCount++ }
                }
                if ($dataset.'Check 11: RgEP' -like 'nicht ok*') {
                    $script:rgEPErrorCount++
                    if ($isSold) { $script:rgEPSoldCount++ }
                }
                if ($dataset.'Check 12: Tiefpreis' -like 'nicht ok*') {
                    $script:tiefpreisErrorCount++
                    if ($isSold) { $script:tiefpreisSoldCount++ }
                }
                if ($dataset.'Check 13: L-Prio Fehlercode' -like 'nicht ok*') {
                    $script:lprioFehlercodeErrorCount++
                    if ($isSold) { $script:lprioFehlercodeSoldCount++ }
                }
                if ($dataset.'Check 14: L-Prio' -like 'nicht ok*') {
                    $script:lprioErrorCount++
                    if ($isSold) { $script:lprioSoldCount++ }
                }
            }
            
            # ErrorsAndWarnings Liste (Fehler ODER Warnung)
            if ($isError -or $hasAnyWarning) {
                [void]$ErrorsAndWarnings_List.Add($dataset)
            }
        }
        
        # Total Warning Count
        $script:totalWarningCount = $script:vpWarningCount + $script:tiefpreisWarningCount + $script:lprioFehlercodeWarningCount + $script:lprioWarningCount
        
        # ArrayLists zu Arrays konvertieren (für Kompatibilität mit bestehendem Code)
        $script:Error_Datasets = @($script:Error_Datasets)
        $ErrorsAndWarnings_Datasets = @($ErrorsAndWarnings_List)
        
        # Export-Arrays OHNE Select-Object (Properties werden beim Export gefiltert)
        $exportAll = $script:All_Datasets
        $exportErrors = $ErrorsAndWarnings_Datasets

        Write-Host "    Datensätze für 'Alle':   $($exportAll.Count)"
        Write-Host "    Datensätze für 'Fehler': $($exportErrors.Count) (inkl. $($script:warningOnlyCount) nur mit Warnungen)"
        Write-Host ""

        # Prüfe ImportExcel-Modul (wird für Excel-Export benötigt)
        $script:ExcelModuleAvailable = $false
        try {
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel -ErrorAction Stop
                $script:ExcelModuleAvailable = $true
                Write-Host "    'ImportExcel'-Modul gefunden und geladen." -ForegroundColor Green
            } else {
                Write-Warning "'ImportExcel'-Modul nicht gefunden."
                $choice = Read-Host "Möchtest du es für den Benutzer '$env:USERNAME' installieren (Internetverbindung nötig)? (j/n)"
                if ($choice -eq 'j') {
                    Write-Host "Installiere 'ImportExcel'..."
                    Install-Module ImportExcel -Scope CurrentUser -AllowClobber -Force -Confirm:$false
                    Import-Module ImportExcel -ErrorAction Stop
                    Write-Host "'ImportExcel' erfolgreich installiert und geladen." -ForegroundColor Green
                    $script:ExcelModuleAvailable = $true
                } else {
                    Write-Warning "Installation übersprungen. Alle Exports erfolgen als CSV."
                }
            }
        } catch {
            Write-Warning "Fehler bei ImportExcel: $($_.Exception.Message)"
            Write-Warning "Fallback: Alle Exports erfolgen als CSV."
            $script:ExcelModuleAvailable = $false
        }
        
        Write-Host ""
        Write-Host "Format-Entscheidung:" -ForegroundColor Yellow
        
        # INTELLIGENTE FORMAT-WAHL PRO OUTPUT-FILE
        # ALLE-File: Excel wenn < 1M Zeilen UND ImportExcel verfügbar
        if ($exportAll.Count -lt $script:EXCEL_EXPORT_LIMIT -and $script:ExcelModuleAvailable) {
            $useExcelForAll = $true
            $fileExtensionAll = ".xlsx"
            Write-Host "  - ALLE-File:   $($exportAll.Count) Zeilen → EXCEL (.xlsx)" -ForegroundColor Green
        } else {
            $useExcelForAll = $false
            $fileExtensionAll = ".csv"
            if ($exportAll.Count -ge $script:EXCEL_EXPORT_LIMIT) {
                Write-Host "  - ALLE-File:   $($exportAll.Count) Zeilen → CSV (>1M Limit)" -ForegroundColor Yellow
            } else {
                Write-Host "  - ALLE-File:   $($exportAll.Count) Zeilen → CSV (kein Excel-Modul)" -ForegroundColor Yellow
            }
        }
        
        # ERRORS-File: Excel wenn < 1M Zeilen UND ImportExcel verfügbar (UNABHÄNGIG vom ALLE-File!)
        if ($exportErrors.Count -lt $script:EXCEL_EXPORT_LIMIT -and $script:ExcelModuleAvailable) {
            $useExcelForErrors = $true
            $fileExtensionErrors = ".xlsx"
            Write-Host "  - ERRORS-File: $($exportErrors.Count) Zeilen → EXCEL (.xlsx)" -ForegroundColor Green
        } else {
            $useExcelForErrors = $false
            $fileExtensionErrors = ".csv"
            if ($exportErrors.Count -ge $script:EXCEL_EXPORT_LIMIT) {
                Write-Host "  - ERRORS-File: $($exportErrors.Count) Zeilen → CSV (>1M Limit)" -ForegroundColor Yellow
            } else {
                Write-Host "  - ERRORS-File: $($exportErrors.Count) Zeilen → CSV (kein Excel-Modul)" -ForegroundColor Yellow
            }
        }

        # Dateinamen mit korrekten Extensions erstellen
        $OutputFilePath_All = $OutputFilePath_All.Replace(".csv", $fileExtensionAll)
        $OutputFilePath_Errors = $OutputFilePath_Errors.Replace(".csv", $fileExtensionErrors)
        $OutputFileName_All = $OutputFileName_All.Replace(".csv", $fileExtensionAll)
        $OutputFileName_Errors = $OutputFileName_Errors.Replace(".csv", $fileExtensionErrors)

        Write-Host ""
        Write-Host "8. Schreibe Ergebnisdateien (nach '$($script:OutputDirectory)')..."
        Write-Host "    Ausgabe-Datei (alle):   '$OutputFileName_All'" -ForegroundColor Cyan
        Write-Host "    Ausgabe-Datei (Fehler): '$OutputFileName_Errors'" -ForegroundColor Cyan
        Write-Host ""

        # Summary-Objekte vorbereiten (für Excel-Exports)
        if ($useExcelForAll -or $useExcelForErrors) {
            $headerSummaryProps = [ordered]@{
                'EAN'                                  = $script:pmsSupplier
                'Check Summary'                        = $script:Error_Datasets.Count
                'Check 0: Vorhanden in beiden Quellen' = $script:presenceErrorCount
                'Check 1: Status'                      = $script:statusErrorCount
                'Check 2: Kategorie'                   = $script:categoryErrorCount
                'Check 3: Genre'                       = $script:genreErrorCount
                'Check 4: Preiscode'                   = $script:preiscodeErrorCount
                'Check 5: Standard VP ab Lieferant'    = $script:standardVPErrorCount
                'Check 6: Fixer VP'                    = $script:fixerVPErrorCount
                'Check 7: Release-Datum'               = $script:releaseDatumErrorCount
                'Check 8: Errorcode'                   = $script:errorCodeErrorCount
                'Check 9: VP'                          = $script:vpErrorCount
                'VP Diff'                              = ' '
                'Check 10: PrioEP'                     = $script:prioEPErrorCount
                'PrioEP Diff'                          = ' '
                'Check 11: RgEP'                       = $script:rgEPErrorCount
                'RgEP Diff'                            = ' '
                'Check 12: Tiefpreis'                  = $script:tiefpreisErrorCount
                'Check 13: L-Prio Fehlercode'          = $script:lprioFehlercodeErrorCount
                'Check 14: L-Prio'                     = $script:lprioErrorCount
                'L-Prio Diff'                          = ' '
                'ZeitDiff letzte Änderung'             = ' '
                'ZeitDiff Bewertung'                   = ' '
            }
            foreach ($header in $script:PMS_Header_Expected) {
                if ($header -ne 'SLLLFN' -and $header -ne 'SLLEAN') { $headerSummaryProps["PMS_$header"] = ' ' }
            }
            foreach ($header in $script:PIM_Header_Expected) { $headerSummaryProps["PIM_$header"] = ' ' }
            $headerSummary = [PSCustomObject]$headerSummaryProps
            
            $totalWarnings = $script:vpWarningCount + $script:tiefpreisWarningCount + $script:lprioFehlercodeWarningCount + $script:lprioWarningCount
            $warningSummaryProps = [ordered]@{
                'EAN'                                  = "Warnungen"
                'Check Summary'                        = $totalWarnings
                'Check 0: Vorhanden in beiden Quellen' = 0
                'Check 1: Status'                      = 0
                'Check 2: Kategorie'                   = 0
                'Check 3: Genre'                       = 0
                'Check 4: Preiscode'                   = 0
                'Check 5: Standard VP ab Lieferant'    = 0
                'Check 6: Fixer VP'                    = 0
                'Check 7: Release-Datum'               = 0
                'Check 8: Errorcode'                   = 0
                'Check 9: VP'                          = $script:vpWarningCount
                'VP Diff'                              = ' '
                'Check 10: PrioEP'                     = 0
                'PrioEP Diff'                          = ' '
                'Check 11: RgEP'                       = 0
                'RgEP Diff'                            = ' '
                'Check 12: Tiefpreis'                  = $script:tiefpreisWarningCount
                'Check 13: L-Prio Fehlercode'          = $script:lprioFehlercodeWarningCount
                'Check 14: L-Prio'                     = $script:lprioWarningCount
                'L-Prio Diff'                          = ' '
                'ZeitDiff letzte Änderung'             = ' '
                'ZeitDiff Bewertung'                   = ' '
            }
            foreach ($header in $script:PMS_Header_Expected) {
                if ($header -ne 'SLLLFN' -and $header -ne 'SLLEAN') { $warningSummaryProps["PMS_$header"] = ' ' }
            }
            foreach ($header in $script:PIM_Header_Expected) { $warningSummaryProps["PIM_$header"] = ' ' }
            $warningSummary = [PSCustomObject]$warningSummaryProps
        }

        # EXPORT: ALLE-File
        Write-Host "    - Schreibe Datei mit allen Datensätzen nach '$OutputFilePath_All'..."
        if ($useExcelForAll) {
            $exportAll | Select-Object * -ExcludeProperty 'Gefunden ...', 'LfNr', 'PMS_SLLEAN' | Export-Excel -Path $OutputFilePath_All -WorksheetName "Vergleich" -ClearSheet -StartRow 3 -AutoFilter -FreezePane 4, 2
            Apply-SummaryRow -Path $OutputFilePath_All -WorksheetName "Vergleich" -HeaderSummary $headerSummary -WarningSummary $warningSummary -ScriptVersion $global:ScriptVersion -SupplierNumber $script:pmsSupplier
            Optimize-ColumnWidthForHeader -Path $OutputFilePath_All -WorksheetName "Vergleich"
            Color-HeaderBySource -Path $OutputFilePath_All -WorksheetName "Vergleich"
            Write-Host "      Erfolgreich geschrieben (Excel)." -ForegroundColor Green
        } else {
            # Für CSV: Filtern vor Export
            $exportAllFiltered = $exportAll | Select-Object * -ExcludeProperty 'Gefunden ...', 'LfNr', 'PMS_SLLEAN'
            Export-CsvFast -Data ([System.Collections.ArrayList]::new($exportAllFiltered)) -Path $OutputFilePath_All -Delimiter ';'
            Write-Host "      Erfolgreich geschrieben (CSV)." -ForegroundColor Green
        }
        $createdOutputFiles.Add($OutputFileName_All)

        # EXPORT: ERRORS-File
        Write-Host "    - Filtere und schreibe Datei mit fehlerhaften Datensätzen nach '$OutputFilePath_Errors'..."
        if ($exportErrors.Count -gt 0) {
            if ($useExcelForErrors) {
                $exportErrors | Select-Object * -ExcludeProperty 'Gefunden ...', 'LfNr', 'PMS_SLLEAN' | Export-Excel -Path $OutputFilePath_Errors -WorksheetName "Fehler" -ClearSheet -StartRow 3 -AutoFilter -FreezePane 4, 2
                Apply-SummaryRow -Path $OutputFilePath_Errors -WorksheetName "Fehler" -HeaderSummary $headerSummary -WarningSummary $warningSummary -ScriptVersion $global:ScriptVersion -SupplierNumber $script:pmsSupplier
                Optimize-ColumnWidthForHeader -Path $OutputFilePath_Errors -WorksheetName "Fehler"
                Color-HeaderBySource -Path $OutputFilePath_Errors -WorksheetName "Fehler"
                Write-Host "      Erfolgreich geschrieben (Excel). ($($script:Error_Datasets.Count) Fehler, $($script:warningOnlyCount) nur Warnungen)" -ForegroundColor Green
            } else {
                # Für CSV: Filtern vor Export
                $exportErrorsFiltered = $exportErrors | Select-Object * -ExcludeProperty 'Gefunden ...', 'LfNr', 'PMS_SLLEAN'
                Export-CsvFast -Data ([System.Collections.ArrayList]::new($exportErrorsFiltered)) -Path $OutputFilePath_Errors -Delimiter ';'
                Write-Host "      Erfolgreich geschrieben (CSV). ($($script:Error_Datasets.Count) Fehler, $($script:warningOnlyCount) nur Warnungen)" -ForegroundColor Green
            }
            $createdOutputFiles.Add($OutputFileName_Errors)
        } else {
            Write-Host "      Keine fehlerhaften Datensätze oder Warnungen gefunden, Fehler-Datei wird nicht erstellt." -ForegroundColor Green
        }

        Write-Host ""
        Write-Host "--------------------------------------------------------" -ForegroundColor Green
        Write-Host "Verarbeitung abgeschlossen."
        
        if ($createdOutputFiles.Count -gt 0) {
            Write-Host ""
            Write-Host "Verschiebe Output-Files nach SharePoint..." -ForegroundColor Yellow
            $sharePointDir = ".\VergleichsErgebnisseBerechnung"
            if (-not (Test-Path $sharePointDir -PathType Container)) {
                New-Item -Path $sharePointDir -ItemType Directory | Out-Null
            }
            $sharePointSubDir = Join-Path $sharePointDir $script:sanitizedSupplierName
            if (-not (Test-Path $sharePointSubDir -PathType Container)) {
                Write-Host "  - Erstelle Lieferanten-Unterverzeichnis: $sharePointSubDir" -ForegroundColor Gray
                New-Item -Path $sharePointSubDir -ItemType Directory | Out-Null
            }
            
            foreach ($fileName in $createdOutputFiles) {
                $sourcePath = Join-Path $script:OutputDirectory $fileName
                $destPath = Join-Path $sharePointSubDir $fileName
                try {
                    Move-Item -Path $sourcePath -Destination $destPath -Force
                    Write-Host "  ✓ $fileName" -ForegroundColor Green
                } catch {
                    Write-Host "  ✗ $fileName - Fehler: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            Write-Host "Verschieben abgeschlossen." -ForegroundColor Green
        }
        
        $script:scriptSuccessfullyCompleted = $true
    }
    catch {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        Write-Host "EIN FEHLER IST AUFGETRETEN:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Skript-Fehler", "OK", "Error")
    }
    finally {
        if ($script:stopwatch -and $script:stopwatch.IsRunning) { $script:stopwatch.Stop() }
        if ($script:scriptSuccessfullyCompleted) {
            Pause-Ende
        } else {
            Write-Host "`nDrücke ENTER um das Fenster zu schliessen."
            Read-Host
        }
    }
}

Write-Host "  Hauptlogik geladen." -ForegroundColor Green

# =====================================================================
# MODUL 7: START-LOGIK (v1.18)
# =====================================================================
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  PhaseX Berechnung - Version $($global:ScriptVersion)" -ForegroundColor White
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Script läuft direkt - kein Relaunch nötig

# Hauptlogik ausführen
try {
    Invoke-MainLogic
} catch {
    Write-Host "KRITISCHER FEHLER: $_" -ForegroundColor Red
    Read-Host "Drücke ENTER"
}
