<#
.SYNOPSIS
    Hauptlogik für PMS/PIM Vergleich - wird von Start.ps1 geladen

.NOTES
    File:           main_v1.3.ps1
    Version:        1.3
    Änderungshistorie:
        1.3 - Fix: Verwendet $global:ScriptVersion (nicht $script:ScriptVersion)
            - Damit die Version von Start.ps1 korrekt angezeigt wird
        1.2 - PrioEP Diff wird immer berechnet wenn Werte unterschiedlich
            - Auch bei "ok - Diff von X" (Toleranz erfuellt) wird Diff angezeigt
            - Wichtig fuer Check 14 Korrelationspruefung
        1.1 - ScriptVersion wird von Start.ps1 uebernommen (nicht mehr lokal definiert)
            - Verwendet $script:ScriptVersion aus Start.ps1
        1.0 - Initiale Version (aus V1.103 extrahiert)
            - Hauptlogik aus ursprünglichem Script
            - Bootstrap und Versionsprüfung in Start.ps1 ausgelagert
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Main = "1.3"

# HINWEIS: $global:ScriptVersion wird von Start.ps1 gesetzt (z.B. "Berechnung_V1.7")

# =====================================================================
# GLOBALE VARIABLEN FÜR ZUSAMMENFASSUNG
# =====================================================================
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

# Sold-Counts
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

# =====================================================================
# PAUSE-FUNKTION
# =====================================================================
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
    Write-Host "Output-Files (im Ordner '$OutputDirectory'):" -ForegroundColor Yellow
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
    
    # Fehler-Übersicht Tabelle
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
    
    if ($global:__ForcePauseAtEnd) { 
        Write-Host "Hinweis: Relaunch mit eigenem Fenster war nicht möglich. Fenster bleibt hier offen." -ForegroundColor Yellow 
    }
    [void](Read-Host "Drücke ENTER um das Fenster zu schliessen")
}

# =====================================================================
# HAUPTLOGIK
# =====================================================================
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

        # OneDrive-Pfad ermitteln
        try {
            $oneDrive = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive\Accounts\Business1" -ErrorAction Stop).UserFolder
            if (-not $oneDrive) { throw "UserFolder ist leer." }
            $script:InputDirectory = Join-Path $oneDrive "PIM\PhaseX_Berechnung"
        } catch {
            throw "KRITISCHER FEHLER: Der OneDrive-Pfad konnte nicht ermittelt werden. Details: $($_.Exception.Message)"
        }

        Write-Host "--- Skript-Version $($global:ScriptVersion) ---`n" -ForegroundColor Gray
        $script:stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        Write-Host "1. Prüfe Eingabe-Verzeichnis..."
        if (-not (Test-Path $script:InputDirectory -PathType Container)) { throw "Eingabeverzeichnis existiert nicht: '$($script:InputDirectory)'" }
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
                    
                    # Entferne Komma als Tausender-Trennzeichen
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
                    
                    # Wenn Check 2 "Kein Kat-Mapping" ergibt, alle weiteren Checks überspringen
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
                        # V1.2: PrioEP Diff immer berechnen wenn Werte unterschiedlich (auch bei Toleranz-OK)
                        $pmsVal = 0.0; $pimVal = 0.0
                        $pmsOk = [decimal]::TryParse($dataset.PMS_SLLEPL, [ref]$pmsVal)
                        $pimOk = [decimal]::TryParse($dataset.PIM_PrioEP, [ref]$pimVal)
                        if ($pmsOk -and $pimOk -and $pmsVal -ne $pimVal) {
                            $dataset.'PrioEP Diff' = $pmsVal - $pimVal
                        }

                        $dataset.'Check 11: RgEP' = Invoke-Check11_RgEP -Dataset $dataset
                        if ($dataset.'Check 11: RgEP' -eq 'nicht ok') {
                            $pmsVal = 0.0; $pimVal = 0.0
                            $pmsOk = [decimal]::TryParse($dataset.PMS_SLOEPF, [ref]$pmsVal)
                            $pimOk = [decimal]::TryParse($dataset.PIM_RgEP, [ref]$pimVal)
                            $dataset.'RgEP Diff' = if ($pmsOk -and $pimOk) { $pmsVal - $pimVal } else { "ungültige Werte" }
                        }

                        $dataset.'Check 12: Tiefpreis' = Invoke-Check12_Tiefpreis -Dataset $dataset

                        # Check 13 mit erweiterter Logik
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

        Write-Host "7. Bereite Export vor..."
        $script:Error_Datasets = @($script:All_Datasets | Where-Object { $_.'Check Summary' -notlike 'ok*' })
        $totalRowCount = $script:All_Datasets.Count
        $script:UseExcelExport = $false
        $fileExtension = ".csv"

        # Warnungen zählen
        $script:vpWarningCount = @($script:All_Datasets | Where-Object { $_.'Check 9: VP' -like 'Warnung*' }).Count
        $script:tiefpreisWarningCount = @($script:All_Datasets | Where-Object { $_.'Check 12: Tiefpreis' -like 'Warnung*' }).Count
        $script:lprioFehlercodeWarningCount = @($script:All_Datasets | Where-Object { $_.'Check 13: L-Prio Fehlercode' -like 'Warnung*' }).Count
        $script:lprioWarningCount = @($script:All_Datasets | Where-Object { $_.'Check 14: L-Prio' -like 'Warnung*' }).Count
        $script:totalWarningCount = $script:vpWarningCount + $script:tiefpreisWarningCount + $script:lprioFehlercodeWarningCount + $script:lprioWarningCount
        
        # Datasets mit Warnungen (Check Summary = ok, aber hat Warnungen)
        $Warning_Datasets = @($script:All_Datasets | Where-Object { 
            $_.'Check Summary' -like 'ok*' -and (
                $_.'Check 9: VP' -like 'Warnung*' -or 
                $_.'Check 12: Tiefpreis' -like 'Warnung*' -or
                $_.'Check 13: L-Prio Fehlercode' -like 'Warnung*' -or 
                $_.'Check 14: L-Prio' -like 'Warnung*'
            )
        })
        $script:warningOnlyCount = $Warning_Datasets.Count

        if ($script:Error_Datasets.Count -gt 0) {
            $script:presenceErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 0: Vorhanden in beiden Quellen' -like 'nicht ok*' }).Count
            $script:statusErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 1: Status' -like 'nicht ok*' }).Count
            $script:categoryErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 2: Kategorie' -like 'nicht ok*' }).Count
            $script:genreErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 3: Genre' -like 'nicht ok*' }).Count
            $script:preiscodeErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 4: Preiscode' -like 'nicht ok*' }).Count
            $script:standardVPErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 5: Standard VP ab Lieferant' -like 'nicht ok*' }).Count
            $script:fixerVPErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 6: Fixer VP' -like 'nicht ok*' }).Count
            $script:releaseDatumErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 7: Release-Datum' -like 'nicht ok*' }).Count
            $script:errorCodeErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 8: Errorcode' -like 'nicht ok*' }).Count
            $script:vpErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 9: VP' -like 'nicht ok*' }).Count
            $script:prioEPErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 10: PrioEP' -like 'nicht ok*' }).Count
            $script:rgEPErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 11: RgEP' -like 'nicht ok*' }).Count
            $script:tiefpreisErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 12: Tiefpreis' -like 'nicht ok*' }).Count
            $script:lprioFehlercodeErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 13: L-Prio Fehlercode' -like 'nicht ok*' }).Count
            $script:lprioErrorCount = @($script:Error_Datasets | Where-Object { $_.'Check 14: L-Prio' -like 'nicht ok*' }).Count
        }

        $exportAll = @($script:All_Datasets | Select-Object * -ExcludeProperty 'Gefunden ...', 'LfNr', 'PMS_SLLEAN')
        $ErrorsAndWarnings_Datasets = @($script:All_Datasets | Where-Object { 
            $_.'Check Summary' -notlike 'ok*' -or 
            $_.'Check 9: VP' -like 'Warnung*' -or 
            $_.'Check 12: Tiefpreis' -like 'Warnung*' -or
            $_.'Check 13: L-Prio Fehlercode' -like 'Warnung*' -or 
            $_.'Check 14: L-Prio' -like 'Warnung*'
        })
        
        # Sold-Counts
        $script:presenceSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 0: Vorhanden in beiden Quellen' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:statusSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 1: Status' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:categorySoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 2: Kategorie' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:genreSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 3: Genre' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:preiscodeSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 4: Preiscode' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:standardVPSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 5: Standard VP ab Lieferant' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:fixerVPSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 6: Fixer VP' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:releaseDatumSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 7: Release-Datum' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:errorCodeSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 8: Errorcode' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:vpSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 9: VP' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:prioEPSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 10: PrioEP' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:rgEPSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 11: RgEP' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:tiefpreisSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 12: Tiefpreis' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:lprioFehlercodeSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 13: L-Prio Fehlercode' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        $script:lprioSoldCount = @($script:Error_Datasets | Where-Object { ($_.'Check 14: L-Prio' -like 'nicht ok*') -and $_.PMS_FLGVKF -eq '1' }).Count
        
        $exportErrors = @($ErrorsAndWarnings_Datasets | Select-Object * -ExcludeProperty 'Gefunden ...', 'LfNr', 'PMS_SLLEAN')

        Write-Host "    Datensätze für 'Alle':   $($exportAll.Count)"
        Write-Host "    Datensätze für 'Fehler': $($exportErrors.Count) (inkl. $($script:warningOnlyCount) nur mit Warnungen)"

        if ($totalRowCount -ge 1000000) {
            Write-Warning "Mehr als 1 Million Zeilen ($totalRowCount) gefunden. Export erfolgt als CSV."
        } else {
            try {
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    Import-Module ImportExcel -ErrorAction Stop
                    $script:UseExcelExport = $true
                    $fileExtension = ".xlsx"
                    Write-Host "    'ImportExcel'-Modul gefunden und geladen. Erstelle .xlsx-Datei(en)." -ForegroundColor Green
                } else {
                    Write-Warning "'ImportExcel'-Modul nicht gefunden."
                    $choice = Read-Host "Möchtest du es für den Benutzer '$env:USERNAME' installieren (Internetverbindung nötig)? (j/n)"
                    if ($choice -eq 'j') {
                        Write-Host "Installiere 'ImportExcel'..."
                        Install-Module ImportExcel -Scope CurrentUser -AllowClobber -Force -Confirm:$false
                        Import-Module ImportExcel -ErrorAction Stop
                        Write-Host "'ImportExcel' erfolgreich installiert und geladen." -ForegroundColor Green
                        $script:UseExcelExport = $true
                        $fileExtension = ".xlsx"
                    } else {
                        Write-Warning "Installation übersprungen. Fallback auf CSV-Export."
                    }
                }
            } catch {
                Write-Warning "Fehler bei ImportExcel: $($_.Exception.Message)"
                Write-Warning "Fallback auf CSV-Export."
                $script:UseExcelExport = $false
                $fileExtension = ".csv"
            }
        }

        $OutputFilePath_All = $OutputFilePath_All.Replace(".csv", $fileExtension)
        $OutputFilePath_Errors = $OutputFilePath_Errors.Replace(".csv", $fileExtension)
        $OutputFileName_All = $OutputFileName_All.Replace(".csv", $fileExtension)
        $OutputFileName_Errors = $OutputFileName_Errors.Replace(".csv", $fileExtension)

        Write-Host ""
        Write-Host "8. Schreibe Ergebnisdateien (nach '$($script:OutputDirectory)')..."
        Write-Host "    Ausgabe-Datei (alle):   '$OutputFileName_All'" -ForegroundColor Cyan
        Write-Host "    Ausgabe-Datei (Fehler): '$OutputFileName_Errors'" -ForegroundColor Cyan
        Write-Host ""

        if ($script:UseExcelExport) {
            # Fehler-Zeile (Zeile 2, rot)
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
            
            # Warnungs-Zeile (Zeile 1, hell-orange)
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

            Write-Host "    - Schreibe Datei mit allen Datensätzen nach '$OutputFilePath_All'..."
            $exportAll | Export-Excel -Path $OutputFilePath_All -WorksheetName "Vergleich" -ClearSheet -StartRow 3 -AutoFilter -FreezePane 4, 2
            Apply-SummaryRow -Path $OutputFilePath_All -WorksheetName "Vergleich" -HeaderSummary $headerSummary -WarningSummary $warningSummary -ScriptVersion $global:ScriptVersion -SupplierNumber $script:pmsSupplier
            $createdOutputFiles.Add($OutputFileName_All)
            Write-Host "      Erfolgreich geschrieben." -ForegroundColor Green

            Write-Host "    - Filtere und schreibe Datei mit fehlerhaften Datensätzen nach '$OutputFilePath_Errors'..."
            if ($exportErrors.Count -gt 0) {
                $exportErrors | Export-Excel -Path $OutputFilePath_Errors -WorksheetName "Fehler" -ClearSheet -StartRow 3 -AutoFilter -FreezePane 4, 2
                Apply-SummaryRow -Path $OutputFilePath_Errors -WorksheetName "Fehler" -HeaderSummary $headerSummary -WarningSummary $warningSummary -ScriptVersion $global:ScriptVersion -SupplierNumber $script:pmsSupplier
                $createdOutputFiles.Add($OutputFileName_Errors)
                Write-Host "      Erfolgreich geschrieben. ($($script:Error_Datasets.Count) Fehler, $($script:warningOnlyCount) nur Warnungen)" -ForegroundColor Green
            } else {
                Write-Host "      Keine fehlerhaften Datensätze oder Warnungen gefunden, Fehler-Datei wird nicht erstellt." -ForegroundColor Green
            }

            Optimize-ColumnWidthForHeader -Path $OutputFilePath_All -WorksheetName "Vergleich"
            Color-HeaderBySource -Path $OutputFilePath_All -WorksheetName "Vergleich"
            if ($exportErrors.Count -gt 0) {
                Optimize-ColumnWidthForHeader -Path $OutputFilePath_Errors -WorksheetName "Fehler"
                Color-HeaderBySource -Path $OutputFilePath_Errors -WorksheetName "Fehler"
            }
        } else {
            Write-Host "    - Schreibe Datei mit allen Datensätzen nach '$OutputFilePath_All'..."
            Export-CsvFast -Data $exportAll -Path $OutputFilePath_All -Delimiter ';'
            $createdOutputFiles.Add($OutputFileName_All)
            Write-Host "      Erfolgreich geschrieben." -ForegroundColor Green
            Write-Host "    - Filtere und schreibe Datei mit fehlerhaften Datensätzen nach '$OutputFilePath_Errors'..."
            if ($exportErrors.Count -gt 0) {
                Export-CsvFast -Data $exportErrors -Path $OutputFilePath_Errors -Delimiter ';'
                $createdOutputFiles.Add($OutputFileName_Errors)
                Write-Host "      Erfolgreich geschrieben. ($($script:Error_Datasets.Count) Fehler, $($script:warningOnlyCount) nur Warnungen)" -ForegroundColor Green
            } else {
                Write-Host "      Keine fehlerhaften Datensätze oder Warnungen gefunden, Fehler-Datei wird nicht erstellt." -ForegroundColor Green
            }
        }

        Write-Host ""
        Write-Host "--------------------------------------------------------" -ForegroundColor Green
        Write-Host "Verarbeitung abgeschlossen."
        
        # Verschiebe Files nach SharePoint
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
