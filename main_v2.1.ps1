<#
.SYNOPSIS
    Hauptlogik für PMS/PIM Vergleich - wird von Start.ps1 geladen
    PERFORMANCE-OPTIMIERT für 12+ Mio Zeilen

.NOTES
    File:           main_v2.1.ps1
    Version:        2.1
    Änderungshistorie:
        2.1 - Fortschrittsausgabe alle 100k Zeilen (mit Tausendertrennzeichen)
        2.0 - PERFORMANCE: Radikale Optimierung fuer 12+ Mio Zeilen
            - String-Arrays statt PSCustomObject (~80% weniger RAM)
            - Generic Dictionary mit initialer Kapazitaet
            - Streaming CSV-Export
            - Inline-Zaehler (keine Where-Object am Ende)
            - GC alle 100k Zeilen
            - Erwartet ~9 GB RAM statt ~48 GB
        1.9 - PMS-Feld SLOEPF umbenannt zu SLOERG
        1.8 - DEBUG-Ausgaben entfernt
        1.7 - Fix: $PSScriptRoot statt $MyInvocation
#>

# =====================================================================
# MODUL-VERSION
# =====================================================================
$script:ModuleVersion_Main = "2.1"

# =====================================================================
# GLOBALE VARIABLEN
# =====================================================================
$script:pmsEanCount = 0
$script:pimEanCount = 0
$script:supplierNameForSummary = ""
$script:foundPimDuplicates = $false
$script:matchedCount = 0
$script:errorCount = 0
$script:warningOnlyCount = 0

# Error-Counts (werden inline gezaehlt)
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
$script:vpWarningCount = 0
$script:prioEPErrorCount = 0
$script:rgEPErrorCount = 0
$script:tiefpreisErrorCount = 0
$script:tiefpreisWarningCount = 0
$script:lprioFehlercodeErrorCount = 0
$script:lprioFehlercodeWarningCount = 0
$script:lprioErrorCount = 0
$script:lprioWarningCount = 0

# Sold-Counts
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

# =====================================================================
# HELPER: Neues Result-Array erstellen
# =====================================================================
function New-ResultArray {
    return [string[]]::new($script:RES_ARRAY_SIZE)
}

# =====================================================================
# HELPER: Fehler/Warnungen inline zaehlen
# =====================================================================
function Update-ErrorCounts {
    param([string[]]$Data, [bool]$IsError)
    
    if (-not $IsError) { return }
    
    $R = $script:RES_IDX
    $isSold = ($Data[$R.PMS_FLGVKF] -eq '1')
    
    if ($Data[$R.Check0] -like 'nicht ok*') { 
        $script:presenceErrorCount++
        if ($isSold) { $script:presenceSoldCount++ }
    }
    if ($Data[$R.Check1] -like 'nicht ok*') { 
        $script:statusErrorCount++
        if ($isSold) { $script:statusSoldCount++ }
    }
    if ($Data[$R.Check2] -like 'nicht ok*') { 
        $script:categoryErrorCount++
        if ($isSold) { $script:categorySoldCount++ }
    }
    if ($Data[$R.Check3] -like 'nicht ok*') { 
        $script:genreErrorCount++
        if ($isSold) { $script:genreSoldCount++ }
    }
    if ($Data[$R.Check4] -like 'nicht ok*') { 
        $script:preiscodeErrorCount++
        if ($isSold) { $script:preiscodeSoldCount++ }
    }
    if ($Data[$R.Check5] -like 'nicht ok*') { 
        $script:standardVPErrorCount++
        if ($isSold) { $script:standardVPSoldCount++ }
    }
    if ($Data[$R.Check6] -like 'nicht ok*') { 
        $script:fixerVPErrorCount++
        if ($isSold) { $script:fixerVPSoldCount++ }
    }
    if ($Data[$R.Check7] -like 'nicht ok*') { 
        $script:releaseDatumErrorCount++
        if ($isSold) { $script:releaseDatumSoldCount++ }
    }
    if ($Data[$R.Check8] -like 'nicht ok*') { 
        $script:errorCodeErrorCount++
        if ($isSold) { $script:errorCodeSoldCount++ }
    }
    if ($Data[$R.Check9] -like 'nicht ok*') { 
        $script:vpErrorCount++
        if ($isSold) { $script:vpSoldCount++ }
    }
    if ($Data[$R.Check10] -like 'nicht ok*') { 
        $script:prioEPErrorCount++
        if ($isSold) { $script:prioEPSoldCount++ }
    }
    if ($Data[$R.Check11] -like 'nicht ok*') { 
        $script:rgEPErrorCount++
        if ($isSold) { $script:rgEPSoldCount++ }
    }
    if ($Data[$R.Check12] -like 'nicht ok*') { 
        $script:tiefpreisErrorCount++
        if ($isSold) { $script:tiefpreisSoldCount++ }
    }
    if ($Data[$R.Check13] -like 'nicht ok*') { 
        $script:lprioFehlercodeErrorCount++
        if ($isSold) { $script:lprioFehlercodeSoldCount++ }
    }
    if ($Data[$R.Check14] -like 'nicht ok*') { 
        $script:lprioErrorCount++
        if ($isSold) { $script:lprioSoldCount++ }
    }
}

function Update-WarningCounts {
    param([string[]]$Data)
    $R = $script:RES_IDX
    
    if ($Data[$R.Check9] -like 'Warnung*') { $script:vpWarningCount++ }
    if ($Data[$R.Check12] -like 'Warnung*') { $script:tiefpreisWarningCount++ }
    if ($Data[$R.Check13] -like 'Warnung*') { $script:lprioFehlercodeWarningCount++ }
    if ($Data[$R.Check14] -like 'Warnung*') { $script:lprioWarningCount++ }
}

# =====================================================================
# HELPER: Array zu CSV-Zeile konvertieren
# =====================================================================
function ConvertTo-CsvLine {
    param([string[]]$Data, [string]$Delimiter = ';')
    
    $R = $script:RES_IDX
    $values = [System.Collections.Generic.List[string]]::new(76)
    
    # Output-Reihenfolge (ohne "Gefunden...")
    $values.Add($Data[$R.EAN])
    $values.Add($Data[$R.CheckSummary])
    $values.Add($Data[$R.Check0])
    $values.Add($Data[$R.Check1])
    $values.Add($Data[$R.Check2])
    $values.Add($Data[$R.Check3])
    $values.Add($Data[$R.Check4])
    $values.Add($Data[$R.Check5])
    $values.Add($Data[$R.Check6])
    $values.Add($Data[$R.Check7])
    $values.Add($Data[$R.Check8])
    $values.Add($Data[$R.Check9])
    $values.Add($Data[$R.VPDiff])
    $values.Add($Data[$R.Check10])
    $values.Add($Data[$R.PrioEPDiff])
    $values.Add($Data[$R.Check11])
    $values.Add($Data[$R.RgEPDiff])
    $values.Add($Data[$R.Check12])
    $values.Add($Data[$R.Check13])
    $values.Add($Data[$R.Check14])
    $values.Add($Data[$R.LPrioDiff])
    $values.Add($Data[$R.ZeitDiff])
    $values.Add($Data[$R.ZeitDiffBewertung])
    
    # PMS-Felder (24-48)
    for ($i = 24; $i -le 48; $i++) {
        $values.Add($Data[$i])
    }
    
    # PIM-Felder (49-75)
    for ($i = 49; $i -le 75; $i++) {
        $values.Add($Data[$i])
    }
    
    return $values -join $Delimiter
}

# =====================================================================
# PAUSE-FUNKTION
# =====================================================================
function Pause-Ende {
    param(
        [string]$pmsFilePath,
        [string]$pimFilePath,
        [string]$OutputDirectory,
        [System.Collections.Generic.List[string]]$createdOutputFiles,
        [System.Diagnostics.Stopwatch]$stopwatch
    )
    
    Write-Host ""
    $ColorOk = "Green"
    $ColorNok = "Red"
    $Line = "=" * 60
    $supplierDisplayString = $script:pmsSupplier
    if ($script:supplierNameForSummary -ne $script:pmsSupplier) { 
        $supplierDisplayString = "$($script:supplierNameForSummary) ($($script:pmsSupplier))" 
    }
    
    $totalCount = $script:pmsEanCount + ($script:pimEanCount - $script:matchedCount)
    $successCount = $totalCount - $script:errorCount
    
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
    Write-Host "  Anzahl EANs im PMS-File: $($script:pmsEanCount)"
    Write-Host "  Anzahl EANs im PIM-File: $($script:pimEanCount)"
    Write-Host "  Anzahl EANs in beiden Files: $($script:matchedCount)"
    Write-Host "  Anzahl fehlerfreie EANs: $successCount" -ForegroundColor $ColorOk
    
    $totalWarnings = $script:vpWarningCount + $script:tiefpreisWarningCount + $script:lprioFehlercodeWarningCount + $script:lprioWarningCount
    if ($script:warningOnlyCount -gt 0) {
        Write-Host "  Anzahl EANs mit Warnungen (nur): $($script:warningOnlyCount)" -ForegroundColor Yellow
    } else {
        Write-Host "  Anzahl EANs mit Warnungen (nur): 0" -ForegroundColor $ColorOk
    }
    
    $finalStatusColor = $ColorOk
    if ($script:errorCount -gt 0) {
        Write-Host "  Anzahl EANs mit Fehlern: $($script:errorCount)" -ForegroundColor $ColorNok
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
    
    # Fehler-Übersicht
    if ($script:errorCount -gt 0 -or $totalWarnings -gt 0) {
        Write-Host ""
        Write-Host "Fehler-Übersicht:" -ForegroundColor Yellow
        Write-Host ""
        
        $headerFormat = "{0,-10} {1,-30} {2,15} {3,18} {4,16}"
        $separatorLine = "-" * 90
        
        Write-Host ($headerFormat -f "Check", "Titel", "Anzahl Fehler", "Anzahl Warnungen", "Fehler+Verkauft") -ForegroundColor Cyan
        Write-Host $separatorLine -ForegroundColor Cyan
        
        $checks = @(
            @{Num = "Check 0"; Titel = "Vorhanden in beiden Quellen"; Fehler = $script:presenceErrorCount; Warnung = 0; Sold = $script:presenceSoldCount }
            @{Num = "Check 1"; Titel = "Status"; Fehler = $script:statusErrorCount; Warnung = 0; Sold = $script:statusSoldCount }
            @{Num = "Check 2"; Titel = "Kategorie"; Fehler = $script:categoryErrorCount; Warnung = 0; Sold = $script:categorySoldCount }
            @{Num = "Check 3"; Titel = "Genre"; Fehler = $script:genreErrorCount; Warnung = 0; Sold = $script:genreSoldCount }
            @{Num = "Check 4"; Titel = "Preiscode"; Fehler = $script:preiscodeErrorCount; Warnung = 0; Sold = $script:preiscodeSoldCount }
            @{Num = "Check 5"; Titel = "Standard VP"; Fehler = $script:standardVPErrorCount; Warnung = 0; Sold = $script:standardVPSoldCount }
            @{Num = "Check 6"; Titel = "Fixer VP"; Fehler = $script:fixerVPErrorCount; Warnung = 0; Sold = $script:fixerVPSoldCount }
            @{Num = "Check 7"; Titel = "Release-Datum"; Fehler = $script:releaseDatumErrorCount; Warnung = 0; Sold = $script:releaseDatumSoldCount }
            @{Num = "Check 8"; Titel = "Errorcode"; Fehler = $script:errorCodeErrorCount; Warnung = 0; Sold = $script:errorCodeSoldCount }
            @{Num = "Check 9"; Titel = "VP"; Fehler = $script:vpErrorCount; Warnung = $script:vpWarningCount; Sold = $script:vpSoldCount }
            @{Num = "Check 10"; Titel = "PrioEP"; Fehler = $script:prioEPErrorCount; Warnung = 0; Sold = $script:prioEPSoldCount }
            @{Num = "Check 11"; Titel = "RgEP"; Fehler = $script:rgEPErrorCount; Warnung = 0; Sold = $script:rgEPSoldCount }
            @{Num = "Check 12"; Titel = "Tiefpreis"; Fehler = $script:tiefpreisErrorCount; Warnung = $script:tiefpreisWarningCount; Sold = $script:tiefpreisSoldCount }
            @{Num = "Check 13"; Titel = "L-Prio Fehlercode"; Fehler = $script:lprioFehlercodeErrorCount; Warnung = $script:lprioFehlercodeWarningCount; Sold = $script:lprioFehlercodeSoldCount }
            @{Num = "Check 14"; Titel = "L-Prio"; Fehler = $script:lprioErrorCount; Warnung = $script:lprioWarningCount; Sold = $script:lprioSoldCount }
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
        Write-Host "Hinweis: Relaunch mit eigenem Fenster war nicht möglich." -ForegroundColor Yellow 
    }
    [void](Read-Host "Drücke ENTER um das Fenster zu schliessen")
}

# =====================================================================
# HAUPTLOGIK
# =====================================================================
function Invoke-MainLogic {
    $scriptSuccessfullyCompleted = $false
    $createdOutputFiles = [System.Collections.Generic.List[string]]::new()
    
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

        # Input-Pfad ermitteln
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
        $InputDirectory = Join-Path $pimDir "PhaseX_Berechnung"
        
        Write-Host "    Script-Verzeichnis: $scriptDir" -ForegroundColor Gray
        Write-Host "    Input-Verzeichnis:  $InputDirectory" -ForegroundColor Gray
        Write-Host "--- Skript-Version $($global:ScriptVersion) ---`n" -ForegroundColor Gray
        
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        Write-Host "1. Prüfe Eingabe-Verzeichnis..."
        if (-not (Test-Path $InputDirectory -PathType Container)) { 
            throw "Eingabeverzeichnis existiert nicht: '$InputDirectory'"
        }
        Write-Host "    Verzeichnis ist vorhanden."

        Write-Host "2. Bitte Dateien auswählen..."
        $absInput = Convert-Path $InputDirectory
        $pmsFilePath = Get-FilePathDialog -WindowTitle "Bitte die PMS-Datei auswählen" -InitialDirectory $absInput
        if (-not $pmsFilePath) { Write-Host "Aktion abgebrochen."; return }
        $pimFilePath = Get-FilePathDialog -WindowTitle "Bitte die PIM-Datei auswählen" -InitialDirectory $absInput
        if (-not $pimFilePath) { Write-Host "Aktion abgebrochen."; return }
        Write-Host "    PMS-Datei: $(Split-Path $pmsFilePath -Leaf)"
        Write-Host "    PIM-Datei: $(Split-Path $pimFilePath -Leaf)"

        Write-Host "3. Prüfe Header der CSV-Dateien..."
        Write-Host "    - Prüfe PMS-Datei..."
        $pmsHeaderLine = (Get-Content -Path $pmsFilePath -TotalCount 1).TrimEnd(';')
        if ([string]::IsNullOrWhiteSpace($pmsHeaderLine)) { throw "PMS-Datei ist leer oder Header fehlt." }
        $pmsActualHeader = $pmsHeaderLine.Split(';')
        if ($null -ne (Compare-Object $script:PMS_Header_Expected $pmsActualHeader -CaseSensitive)) {
            throw "Header PMS nicht korrekt.`nErwartet: $($script:PMS_Header_Expected -join ';')`nGefunden: $($pmsActualHeader -join ';')"
        }
        Write-Host "      -> Header in PMS-Datei ist korrekt." -ForegroundColor Green

        Write-Host "    - Prüfe PIM-Datei..."
        $pimHeaderLine = Get-Content -Path $pimFilePath -TotalCount 1 -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($pimHeaderLine)) { throw "PIM-Datei ist leer oder Header fehlt." }
        $pimActualHeader = ($pimHeaderLine.Replace('"', '')).Split(';')
        if ($null -ne (Compare-Object $script:PIM_Header_Expected $pimActualHeader -CaseSensitive)) {
            throw "Header PIM nicht korrekt.`nErwartet: $($script:PIM_Header_Expected -join ';')`nGefunden: $($pimActualHeader -join ';')"
        }
        Write-Host "      -> Header in PIM-Datei ist korrekt." -ForegroundColor Green

        Write-Host "4. Führe Lieferanten-Check durch..."
        $pmsFirstDataRow = (Get-Content -Path $pmsFilePath -TotalCount 2 | Select-Object -Last 1).TrimEnd(';')
        $pmsFields = $pmsFirstDataRow.Split(';')
        $script:pmsSupplier = $pmsFields[$script:PMS_IDX.SLLLFN]

        $pimFirstDataRow = (Get-Content -Path $pimFilePath -TotalCount 2 -Encoding UTF8 | Select-Object -Last 1)
        $pimFields = ($pimFirstDataRow.Replace('"', '')).Split(';')
        $pimSupplier = $pimFields[$script:PIM_IDX.Lieferant]

        if ($script:pmsSupplier -ne $pimSupplier) {
            throw "Lieferantennummern stimmen NICHT überein!`nPMS: '$($script:pmsSupplier)'`nPIM: '$pimSupplier'"
        }

        Write-Host "    Lieferantennummern stimmen überein: '$($script:pmsSupplier)'." -ForegroundColor Green
        $supplierName = $script:pmsSupplier
        if ($script:SupplierLookupTable.ContainsKey($script:pmsSupplier)) { 
            $supplierName = $script:SupplierLookupTable[$script:pmsSupplier] 
        }
        $script:supplierNameForSummary = $supplierName
        $sanitizedSupplierName = $supplierName.Replace(' ', '-').Replace('+', '') -replace '[\\/:*?"<>|]', ''

        # Output-Verzeichnis
        $OutputDirectory = Split-Path -Path $pmsFilePath -Parent
        
        $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $SystemUserName = $env:USERNAME
        $FriendlyUserName = $SystemUserName
        if ($script:UserLookupTable.ContainsKey($SystemUserName)) { 
            $FriendlyUserName = $script:UserLookupTable[$SystemUserName] 
        }

        $OutputFileName_All = "PhaseX_Vergl_Berechnung__$($sanitizedSupplierName)_$($script:pmsSupplier)__$($FriendlyUserName)__ALLE__$($Timestamp).csv"
        $OutputFileName_Errors = "PhaseX_Vergl_Berechnung__$($sanitizedSupplierName)_$($script:pmsSupplier)__$($FriendlyUserName)__ERRORS__$($Timestamp).csv"
        $OutputFilePath_All = Join-Path $OutputDirectory $OutputFileName_All
        $OutputFilePath_Errors = Join-Path $OutputDirectory $OutputFileName_Errors

        # =====================================================================
        # SCHRITT 5: PMS-Datei einlesen (nur PMS-Daten in Dictionary)
        # =====================================================================
        Write-Host "5. Lese und verarbeite Dateien... (Dies kann einige Minuten dauern)"
        Write-Host "    PERFORMANCE-MODUS: Optimiert fuer grosse Dateien (12+ Mio Zeilen)" -ForegroundColor Cyan
        
        # Geschaetzte Kapazitaet: 5 Mio als Start
        $DataDict = [System.Collections.Generic.Dictionary[string,string[]]]::new(5000000)
        $pmsSkippedCounter = 0
        $R = $script:RES_IDX
        $P = $script:PMS_IDX

        Write-Host "    - Verarbeite PMS-Datei..."
        $reader = $null
        $lineCounter = 0
        try {
            $reader = [System.IO.StreamReader]::new($pmsFilePath, [System.Text.Encoding]::Default)
            $null = $reader.ReadLine()  # Header überspringen
            
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                $lineCounter++
                
                # Fortschritt alle 100k Zeilen + GC
                if ($lineCounter % 100000 -eq 0) {
                    Write-Host "      PMS: $($lineCounter.ToString('N0')) Zeilen verarbeitet..." -ForegroundColor Gray
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
                
                $pmsFields = $line.Split(';')
                $ean = Get-SafeField $pmsFields $P.SLLEAN
                
                if ([string]::IsNullOrEmpty($ean)) { 
                    $pmsSkippedCounter++
                    continue 
                }
                
                if ($DataDict.ContainsKey($ean)) {
                    Write-Warning "Doppelte EAN '$ean' in PMS-Datei. Nur erster Eintrag wird berücksichtigt."
                    continue
                }
                
                # Neues Result-Array erstellen
                $data = New-ResultArray
                
                # Meta-Felder
                $data[$R.EAN] = "'$ean"
                $data[$R.Gefunden] = "nur im PMS"
                
                # PMS-Felder kopieren (ohne SLLLFN)
                $data[$R.PMS_SLLEAN] = $ean
                $data[$R.PMS_SLLPAS] = Get-SafeField $pmsFields $P.SLLPAS
                $data[$R.PMS_SLLCAT] = Get-SafeField $pmsFields $P.SLLCAT
                $data[$R.PMS_SLLGNR] = Get-SafeField $pmsFields $P.SLLGNR
                $data[$R.PMS_SLLPCD] = Get-SafeField $pmsFields $P.SLLPCD
                $data[$R.PMS_FLGSTP] = Get-SafeField $pmsFields $P.FLGSTP
                $data[$R.PMS_FLGFXP] = Get-SafeField $pmsFields $P.FLGFXP
                $data[$R.PMS_FLGVKF] = Get-SafeField $pmsFields $P.FLGVKF
                $data[$R.PMS_RELDAT] = Get-SafeField $pmsFields $P.RELDAT
                $data[$R.PMS_XML01] = Get-SafeField $pmsFields $P.XML01
                $data[$R.PMS_XML02] = Get-SafeField $pmsFields $P.XML02
                $data[$R.PMS_XML03] = Get-SafeField $pmsFields $P.XML03
                $data[$R.PMS_XML04] = Get-SafeField $pmsFields $P.XML04
                $data[$R.PMS_XML05] = Get-SafeField $pmsFields $P.XML05
                $data[$R.PMS_SLLVPL] = Get-SafeField $pmsFields $P.SLLVPL
                $data[$R.PMS_SLLEPL] = Get-SafeField $pmsFields $P.SLLEPL
                $data[$R.PMS_SLOERG] = Get-SafeField $pmsFields $P.SLOERG
                $data[$R.PMS_SLOWAH] = Get-SafeField $pmsFields $P.SLOWAH
                $data[$R.PMS_REDVPL] = Get-SafeField $pmsFields $P.REDVPL
                $data[$R.PMS_SLLERR] = Get-SafeField $pmsFields $P.SLLERR
                $data[$R.PMS_SAAPNT] = Get-SafeField $pmsFields $P.SAAPNT
                $data[$R.PMS_SLLIGN] = Get-SafeField $pmsFields $P.SLLIGN
                $data[$R.PMS_IMPDAT] = Get-SafeField $pmsFields $P.IMPDAT
                $data[$R.PMS_CHGDAT] = Get-SafeField $pmsFields $P.CHGDAT
                $data[$R.PMS_SAASEL] = Get-SafeField $pmsFields $P.SAASEL
                
                $DataDict.Add($ean, $data)
            }
        } finally { 
            if ($reader) { $reader.Close(); $reader.Dispose() } 
        }
        
        $script:pmsEanCount = $DataDict.Count
        Write-Host "    - PMS-Datei eingelesen. $($DataDict.Count.ToString('N0')) eindeutige Datensätze gefunden."
        if ($pmsSkippedCounter -gt 0) { Write-Warning "$pmsSkippedCounter Zeilen ohne EAN im PMS-File wurden übersprungen." }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        # =====================================================================
        # SCHRITT 5b: PIM-Datei einlesen und mit PMS matchen
        # =====================================================================
        Write-Host "    - Verarbeite PIM-Datei..."
        $pimSeenEans = [System.Collections.Generic.HashSet[string]]::new()
        $pimOnlyList = [System.Collections.Generic.List[string[]]]::new()
        $pimSkippedCounter = 0
        $PI = $script:PIM_IDX
        $lineCounter = 0
        
        # Felder die Komma-Bereinigung brauchen
        $commaFields = @($PI.FixerVP, $PI.AcquisitionPrice, $PI.SalesPrice, $PI.VP, $PI.PrioEP, $PI.RgEP, $PI.Tiefpreis)
        
        $reader = $null
        try {
            $reader = [System.IO.StreamReader]::new($pimFilePath, [System.Text.Encoding]::UTF8)
            $null = $reader.ReadLine()  # Header überspringen
            
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                $lineCounter++
                
                if ($lineCounter % 100000 -eq 0) {
                    Write-Host "      PIM: $($lineCounter.ToString('N0')) Zeilen verarbeitet..." -ForegroundColor Gray
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
                
                $pimFields = ($line.Replace('"', '')).Split(';')
                $ean = Get-SafeField $pimFields $PI.EAN
                
                if ([string]::IsNullOrEmpty($ean)) { 
                    $pimSkippedCounter++
                    continue 
                }
                
                $script:pimEanCount++
                
                # Duplikat-Check
                if ($pimSeenEans.Contains($ean)) {
                    if ($DataDict.ContainsKey($ean)) {
                        $existing = $DataDict[$ean]
                        $existing[$R.Gefunden] = "mehrfach im PIM"
                        $existing[$R.CheckSummary] = "nicht ok - EAN mehrfach im PIM"
                        $script:foundPimDuplicates = $true
                    }
                    continue
                }
                [void]$pimSeenEans.Add($ean)
                
                # Komma-Bereinigung für numerische Felder
                foreach ($idx in $commaFields) {
                    if ($pimFields.Count -gt $idx) {
                        $pimFields[$idx] = $pimFields[$idx].Replace(',', '')
                    }
                }
                
                if ($DataDict.ContainsKey($ean)) {
                    # Match gefunden - PIM-Daten hinzufügen
                    $script:matchedCount++
                    $data = $DataDict[$ean]
                    $data[$R.Gefunden] = "im PMS und im PIM"
                    
                    # PIM-Felder kopieren
                    $data[$R.PIM_Lieferant] = Get-SafeField $pimFields $PI.Lieferant
                    $data[$R.PIM_EAN] = Get-SafeField $pimFields $PI.EAN
                    $data[$R.PIM_Status] = Get-SafeField $pimFields $PI.Status
                    $data[$R.PIM_Kategorie] = Get-SafeField $pimFields $PI.Kategorie
                    $data[$R.PIM_Genre] = Get-SafeField $pimFields $PI.Genre
                    $data[$R.PIM_Preiscode] = Get-SafeField $pimFields $PI.Preiscode
                    $data[$R.PIM_StandardVP] = Get-SafeField $pimFields $PI.StandardVP
                    $data[$R.PIM_FixerVP] = Get-SafeField $pimFields $PI.FixerVP
                    $data[$R.PIM_ReleaseDate] = Get-SafeField $pimFields $PI.ReleaseDate
                    $data[$R.PIM_AcquisitionPrice] = Get-SafeField $pimFields $PI.AcquisitionPrice
                    $data[$R.PIM_SalesPrice] = Get-SafeField $pimFields $PI.SalesPrice
                    $data[$R.PIM_PublisherID] = Get-SafeField $pimFields $PI.PublisherID
                    $data[$R.PIM_Linedisc] = Get-SafeField $pimFields $PI.Linedisc
                    $data[$R.PIM_Bonusgroup] = Get-SafeField $pimFields $PI.Bonusgroup
                    $data[$R.PIM_VP] = Get-SafeField $pimFields $PI.VP
                    $data[$R.PIM_PrioEP] = Get-SafeField $pimFields $PI.PrioEP
                    $data[$R.PIM_RgEP] = Get-SafeField $pimFields $PI.RgEP
                    $data[$R.PIM_WaehrungRgEP] = Get-SafeField $pimFields $PI.WaehrungRgEP
                    $data[$R.PIM_Tiefpreis] = Get-SafeField $pimFields $PI.Tiefpreis
                    $data[$R.PIM_Errorcode] = Get-SafeField $pimFields $PI.Errorcode
                    $data[$R.PIM_Fehlercode] = Get-SafeField $pimFields $PI.Fehlercode
                    $data[$R.PIM_LPrioPunkte] = Get-SafeField $pimFields $PI.LPrioPunkte
                    $data[$R.PIM_Sperrcode] = Get-SafeField $pimFields $PI.Sperrcode
                    $data[$R.PIM_VerwendeteKalk] = Get-SafeField $pimFields $PI.VerwendeteKalk
                    $data[$R.PIM_LetzterImport] = Get-SafeField $pimFields $PI.LetzterImport
                    $data[$R.PIM_LetzteAenderung] = Get-SafeField $pimFields $PI.LetzteAenderung
                    $data[$R.PIM_LetzterStatus] = Get-SafeField $pimFields $PI.LetzterStatus
                } else {
                    # Nur im PIM - neuen Datensatz erstellen
                    $data = New-ResultArray
                    $data[$R.EAN] = "'$ean"
                    $data[$R.Gefunden] = "nur im PIM"
                    
                    # PIM-Felder kopieren
                    $data[$R.PIM_Lieferant] = Get-SafeField $pimFields $PI.Lieferant
                    $data[$R.PIM_EAN] = Get-SafeField $pimFields $PI.EAN
                    $data[$R.PIM_Status] = Get-SafeField $pimFields $PI.Status
                    $data[$R.PIM_Kategorie] = Get-SafeField $pimFields $PI.Kategorie
                    $data[$R.PIM_Genre] = Get-SafeField $pimFields $PI.Genre
                    $data[$R.PIM_Preiscode] = Get-SafeField $pimFields $PI.Preiscode
                    $data[$R.PIM_StandardVP] = Get-SafeField $pimFields $PI.StandardVP
                    $data[$R.PIM_FixerVP] = Get-SafeField $pimFields $PI.FixerVP
                    $data[$R.PIM_ReleaseDate] = Get-SafeField $pimFields $PI.ReleaseDate
                    $data[$R.PIM_AcquisitionPrice] = Get-SafeField $pimFields $PI.AcquisitionPrice
                    $data[$R.PIM_SalesPrice] = Get-SafeField $pimFields $PI.SalesPrice
                    $data[$R.PIM_PublisherID] = Get-SafeField $pimFields $PI.PublisherID
                    $data[$R.PIM_Linedisc] = Get-SafeField $pimFields $PI.Linedisc
                    $data[$R.PIM_Bonusgroup] = Get-SafeField $pimFields $PI.Bonusgroup
                    $data[$R.PIM_VP] = Get-SafeField $pimFields $PI.VP
                    $data[$R.PIM_PrioEP] = Get-SafeField $pimFields $PI.PrioEP
                    $data[$R.PIM_RgEP] = Get-SafeField $pimFields $PI.RgEP
                    $data[$R.PIM_WaehrungRgEP] = Get-SafeField $pimFields $PI.WaehrungRgEP
                    $data[$R.PIM_Tiefpreis] = Get-SafeField $pimFields $PI.Tiefpreis
                    $data[$R.PIM_Errorcode] = Get-SafeField $pimFields $PI.Errorcode
                    $data[$R.PIM_Fehlercode] = Get-SafeField $pimFields $PI.Fehlercode
                    $data[$R.PIM_LPrioPunkte] = Get-SafeField $pimFields $PI.LPrioPunkte
                    $data[$R.PIM_Sperrcode] = Get-SafeField $pimFields $PI.Sperrcode
                    $data[$R.PIM_VerwendeteKalk] = Get-SafeField $pimFields $PI.VerwendeteKalk
                    $data[$R.PIM_LetzterImport] = Get-SafeField $pimFields $PI.LetzterImport
                    $data[$R.PIM_LetzteAenderung] = Get-SafeField $pimFields $PI.LetzteAenderung
                    $data[$R.PIM_LetzterStatus] = Get-SafeField $pimFields $PI.LetzterStatus
                    
                    $pimOnlyList.Add($data)
                }
            }
        } finally { 
            if ($reader) { $reader.Close(); $reader.Dispose() } 
        }
        
        # PIM-only Datensätze zum Dictionary hinzufügen
        foreach ($data in $pimOnlyList) {
            $ean = $data[$R.EAN].TrimStart("'")
            if (-not $DataDict.ContainsKey($ean)) {
                $DataDict.Add($ean, $data)
            }
        }
        $pimOnlyList.Clear()
        $pimOnlyList = $null
        
        Write-Host "    - PIM-Datei verarbeitet."
        if ($pimSkippedCounter -gt 0) { Write-Warning "$pimSkippedCounter Zeilen ohne EAN im PIM-File wurden übersprungen." }
        Write-Host "Beide Files eingelesen. Dauer $($stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
        Write-Host "Gesamtanzahl eindeutiger Datensätze: $($DataDict.Count)"
        
        $pimSeenEans.Clear()
        $pimSeenEans = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        # =====================================================================
        # SCHRITT 6: Checks durchführen
        # =====================================================================
        Write-Host "6. Führe Checks durch..."
        $totalDatasets = $DataDict.Count
        $i = 0
        
        foreach ($ean in $DataDict.Keys) {
            $i++
            if ($i % 10000 -eq 0) {
                $perc = [Math]::Floor(($i / $totalDatasets) * 100)
                Write-Progress -Activity "Schritt 6: Führe Checks durch" -Status "$perc% ($i von $totalDatasets)" -PercentComplete $perc
            }
            if ($i % 100000 -eq 0) {
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            $data = $DataDict[$ean]
            $gefunden = $data[$R.Gefunden]
            
            # Check 0
            switch ($gefunden) {
                'im PMS und im PIM' { $data[$R.Check0] = 'ok - EAN in beiden Quellen' }
                'nur im PIM' { $data[$R.Check0] = 'ok - EAN nur im PIM' }
                'nur im PMS' {
                    if ($data[$R.PMS_SLLPAS] -eq 'passive') { 
                        $data[$R.Check0] = 'ok - EAN fehlt im PIM - passive im PMS' 
                    } else { 
                        $data[$R.Check0] = 'nicht ok - EAN fehlt im PIM' 
                    }
                }
                'mehrfach im PIM' { $data[$R.Check0] = 'nicht ok - EAN mehrfach im PIM' }
                default { $data[$R.Check0] = 'nicht ok' }
            }
            
            if ($data[$R.CheckSummary] -like 'nicht ok - EAN mehrfach im PIM') { continue }
            
            if ($gefunden -eq "im PMS und im PIM") {
                $data[$R.ZeitDiff] = Invoke-CalculateTimeDifference -Data $data
                $data[$R.Check1] = Invoke-Check1_Status -Data $data
                
                if ($data[$R.Check1] -like 'ok*') {
                    $data[$R.Check2] = Invoke-Check2_Kategorie -Data $data
                    
                    if ($data[$R.Check2] -eq 'ok - Kein Kat-Mapping im PMS und PIM') {
                        $data[$R.CheckSummary] = 'ok'
                        continue
                    }
                    
                    $data[$R.Check3] = Invoke-Check3_Genre -Data $data
                    $data[$R.Check4] = Invoke-Check4_Preiscode -Data $data
                    
                    $whitelistSkipMessage = "ok - Lieferant bei Kategorie nicht auf Whitelist"
                    if ($data[$R.PMS_SAAPNT] -eq "999905") {
                        $data[$R.Check5] = $whitelistSkipMessage
                        $data[$R.Check6] = $whitelistSkipMessage
                        $data[$R.Check7] = $whitelistSkipMessage
                        $data[$R.Check8] = $whitelistSkipMessage
                        $data[$R.Check9] = $whitelistSkipMessage
                        $data[$R.Check10] = $whitelistSkipMessage
                        $data[$R.Check11] = $whitelistSkipMessage
                        $data[$R.Check12] = $whitelistSkipMessage
                        $data[$R.Check14] = $whitelistSkipMessage
                    } else {
                        $data[$R.Check5] = Invoke-Check5_StandardVP -Data $data
                        $data[$R.Check6] = Invoke-Check6_FixerVP -Data $data
                        $data[$R.Check7] = Invoke-Check7_ReleaseDatum -Data $data
                        $data[$R.Check8] = Invoke-Check8_Errorcode -Data $data
                        $data[$R.Check9] = Invoke-Check9_VP -Data $data
                        
                        if ($data[$R.Check9] -eq 'nicht ok') {
                            $pmsVal = 0.0; $pimVal = 0.0
                            $pmsOk = [decimal]::TryParse($data[$R.PMS_SLLVPL], [ref]$pmsVal)
                            $pimOk = [decimal]::TryParse($data[$R.PIM_VP], [ref]$pimVal)
                            $data[$R.VPDiff] = if ($pmsOk -and $pimOk) { ($pmsVal - $pimVal).ToString() } else { "ungültige Werte" }
                        }
                        
                        $data[$R.Check10] = Invoke-Check10_PrioEP -Data $data
                        $pmsVal = 0.0; $pimVal = 0.0
                        $pmsOk = [decimal]::TryParse($data[$R.PMS_SLLEPL], [ref]$pmsVal)
                        $pimOk = [decimal]::TryParse($data[$R.PIM_PrioEP], [ref]$pimVal)
                        if ($pmsOk -and $pimOk -and $pmsVal -ne $pimVal) {
                            $data[$R.PrioEPDiff] = ($pmsVal - $pimVal).ToString()
                        }
                        
                        $data[$R.Check11] = Invoke-Check11_RgEP -Data $data
                        if ($data[$R.Check11] -eq 'nicht ok') {
                            $pmsVal = 0.0; $pimVal = 0.0
                            $pmsOk = [decimal]::TryParse($data[$R.PMS_SLOERG], [ref]$pmsVal)
                            $pimOk = [decimal]::TryParse($data[$R.PIM_RgEP], [ref]$pimVal)
                            $data[$R.RgEPDiff] = if ($pmsOk -and $pimOk) { ($pmsVal - $pimVal).ToString() } else { "ungültige Werte" }
                        }
                        
                        $data[$R.Check12] = Invoke-Check12_Tiefpreis -Data $data
                        $data[$R.Check13] = Invoke-Check13_LPrioFehlercode -Data $data
                        $data[$R.Check13] = Invoke-Check13_Extended -Data $data
                        $data[$R.Check14] = Invoke-Check14_LPrio -Data $data
                        
                        if ($data[$R.Check14] -eq 'nicht ok') {
                            $pmsVal = 0; $pimVal = 0
                            $pmsOk = [long]::TryParse($data[$R.PMS_SAAPNT], [ref]$pmsVal)
                            $pimOk = [long]::TryParse($data[$R.PIM_LPrioPunkte], [ref]$pimVal)
                            $data[$R.LPrioDiff] = if ($pmsOk -and $pimOk) { ($pmsVal - $pimVal).ToString() } else { "ungültige Werte" }
                        }
                    }
                    
                    # Check Summary berechnen
                    if (($data[$R.Check1] -like 'ok*') -and 
                        ($data[$R.Check2] -like 'ok*') -and 
                        ($data[$R.Check3] -like 'ok*') -and
                        ($data[$R.Check4] -like 'ok*') -and
                        ($data[$R.Check5] -like 'ok*') -and
                        ($data[$R.Check6] -like 'ok*') -and
                        ($data[$R.Check7] -like 'ok*') -and
                        ($data[$R.Check8] -like 'ok*') -and
                        ($data[$R.Check9] -like 'ok*' -or $data[$R.Check9] -like 'Warnung*') -and
                        ($data[$R.Check10] -like 'ok*') -and
                        ($data[$R.Check11] -like 'ok*') -and
                        ($data[$R.Check12] -like 'ok*' -or $data[$R.Check12] -like 'Warnung*') -and
                        ($data[$R.Check13] -like 'ok*' -or $data[$R.Check13] -like 'Warnung*') -and
                        ($data[$R.Check14] -like 'ok*' -or $data[$R.Check14] -like 'Warnung*')) {
                        $data[$R.CheckSummary] = 'ok'
                    } else {
                        $data[$R.CheckSummary] = 'nicht ok'
                    }
                } else {
                    # Status nicht ok - alle anderen Checks auf ---
                    $data[$R.Check2] = '---'
                    $data[$R.Check3] = '---'
                    $data[$R.Check4] = '---'
                    $data[$R.Check5] = '---'
                    $data[$R.Check6] = '---'
                    $data[$R.Check7] = '---'
                    $data[$R.Check8] = '---'
                    $data[$R.Check9] = '---'
                    $data[$R.VPDiff] = '---'
                    $data[$R.Check10] = '---'
                    $data[$R.PrioEPDiff] = '---'
                    $data[$R.Check11] = '---'
                    $data[$R.RgEPDiff] = '---'
                    $data[$R.Check12] = '---'
                    $data[$R.Check13] = '---'
                    $data[$R.Check14] = '---'
                    $data[$R.LPrioDiff] = '---'
                }
            } elseif ($gefunden -eq "nur im PIM") {
                $data[$R.CheckSummary] = 'ok - EAN nur im PIM'
            } else {
                if ($data[$R.PMS_SLLPAS] -eq 'passive') { 
                    $data[$R.CheckSummary] = 'ok - EAN fehlt im PIM - passive im PMS' 
                } else { 
                    $data[$R.CheckSummary] = 'nicht ok - EAN fehlt im PIM' 
                }
            }
            
            # Inline Error/Warning Counting
            $isError = ($data[$R.CheckSummary] -notlike 'ok*')
            if ($isError) { $script:errorCount++ }
            Update-ErrorCounts -Data $data -IsError $isError
            Update-WarningCounts -Data $data
            
            # Warning-Only zählen
            if (-not $isError) {
                $hasWarning = ($data[$R.Check9] -like 'Warnung*' -or 
                               $data[$R.Check12] -like 'Warnung*' -or 
                               $data[$R.Check13] -like 'Warnung*' -or 
                               $data[$R.Check14] -like 'Warnung*')
                if ($hasWarning) { $script:warningOnlyCount++ }
            }
        }
        
        Write-Progress -Activity "Schritt 6: Führe Checks durch" -Completed
        Write-Host "    Checks abgeschlossen." -ForegroundColor Green

        # =====================================================================
        # SCHRITT 7+8: Export (Streaming)
        # =====================================================================
        Write-Host "7. Bereite Export vor..."
        Write-Host "    Datensätze gesamt:  $($DataDict.Count)"
        Write-Host "    Datensätze Fehler:  $($script:errorCount)"
        Write-Host "    Datensätze Warnung: $($script:warningOnlyCount)"
        
        Write-Host ""
        Write-Host "8. Schreibe Ergebnisdateien (Streaming-Export)..."
        Write-Host "    Ausgabe-Datei (alle):   '$OutputFileName_All'" -ForegroundColor Cyan
        Write-Host "    Ausgabe-Datei (Fehler): '$OutputFileName_Errors'" -ForegroundColor Cyan
        
        # CSV Header erstellen
        $csvHeader = $script:OUTPUT_HEADER -join ';'
        
        # Streaming Export - ALLE
        Write-Host "    - Schreibe Datei mit allen Datensätzen..."
        $writerAll = $null
        $writerErrors = $null
        $errorFileNeeded = ($script:errorCount -gt 0 -or $script:warningOnlyCount -gt 0)
        
        try {
            $writerAll = [System.IO.StreamWriter]::new($OutputFilePath_All, $false, [System.Text.Encoding]::UTF8)
            $writerAll.WriteLine($csvHeader)
            
            if ($errorFileNeeded) {
                $writerErrors = [System.IO.StreamWriter]::new($OutputFilePath_Errors, $false, [System.Text.Encoding]::UTF8)
                $writerErrors.WriteLine($csvHeader)
            }
            
            $exportCounter = 0
            foreach ($ean in $DataDict.Keys) {
                $exportCounter++
                if ($exportCounter % 100000 -eq 0) {
                    Write-Host "      Export: $exportCounter Zeilen geschrieben..." -ForegroundColor Gray
                }
                
                $data = $DataDict[$ean]
                $csvLine = ConvertTo-CsvLine -Data $data
                $writerAll.WriteLine($csvLine)
                
                # Fehler/Warnungen auch in Error-File schreiben
                if ($errorFileNeeded) {
                    $isErrorOrWarning = ($data[$R.CheckSummary] -notlike 'ok*' -or 
                                         $data[$R.Check9] -like 'Warnung*' -or 
                                         $data[$R.Check12] -like 'Warnung*' -or 
                                         $data[$R.Check13] -like 'Warnung*' -or 
                                         $data[$R.Check14] -like 'Warnung*')
                    if ($isErrorOrWarning) {
                        $writerErrors.WriteLine($csvLine)
                    }
                }
            }
        } finally {
            if ($writerAll) { $writerAll.Close(); $writerAll.Dispose() }
            if ($writerErrors) { $writerErrors.Close(); $writerErrors.Dispose() }
        }
        
        $createdOutputFiles.Add($OutputFileName_All)
        Write-Host "      Erfolgreich geschrieben." -ForegroundColor Green
        
        if ($errorFileNeeded) {
            $createdOutputFiles.Add($OutputFileName_Errors)
            Write-Host "      Fehler-Datei erfolgreich geschrieben." -ForegroundColor Green
        }
        
        # Aufräumen
        $DataDict.Clear()
        $DataDict = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host ""
        Write-Host "--------------------------------------------------------" -ForegroundColor Green
        Write-Host "Verarbeitung abgeschlossen."
        
        $scriptSuccessfullyCompleted = $true
    }
    catch {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        Write-Host "EIN FEHLER IST AUFGETRETEN:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor Gray
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Skript-Fehler", "OK", "Error")
    }
    finally {
        if ($stopwatch -and $stopwatch.IsRunning) { $stopwatch.Stop() }
        if ($scriptSuccessfullyCompleted) {
            Pause-Ende -pmsFilePath $pmsFilePath -pimFilePath $pimFilePath -OutputDirectory $OutputDirectory -createdOutputFiles $createdOutputFiles -stopwatch $stopwatch
        } else {
            Write-Host "`nDrücke ENTER um das Fenster zu schliessen."
            Read-Host
        }
    }
}
