<#
.SYNOPSIS
    Check-Funktionen (Check 1-14) für PMS/PIM Vergleich
    PERFORMANCE-OPTIMIERT für 12+ Mio Zeilen

.NOTES
    File:           functions-checks_v1.7.ps1
    Version:        1.7
    Änderungshistorie:
        1.7 - PERFORMANCE: Komplett auf String-Array-Zugriff umgestellt
            - Verwendet $script:RES_IDX fuer Feld-Zugriff
            - ~80% weniger RAM, ~50% schneller
        1.6 - Check 11 (RgEP): PMS-Feld SLOEPF umbenannt zu SLOERG
        1.5 - Check 7 (Release-Datum): Vergleich auf 2-stelliges Jahr
        1.4 - Check 14 (L-Prio): Fix Korrelationspruefung
        1.3 - Check 14 (L-Prio): Korrelation mit PrioEP-Diff
        1.2 - Checks 9, 10, 11, 12: PMS-Wert "0" entspricht leerem PIM-Feld
        1.1 - Check 10 (PrioEP): Relative Toleranz hinzugefuegt
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Checks = "1.7"

# =====================================================================
# HINWEIS: Alle Funktionen erwarten jetzt ein String-Array $Data
# Zugriff erfolgt über $script:RES_IDX (definiert in config)
# =====================================================================

function Invoke-Check1_Status {
    param([string[]]$Data)
    $R = $script:RES_IDX
    $pmsStatus = $Data[$R.PMS_SLLPAS]
    $pimStatus = $Data[$R.PIM_Status]
    if ($pmsStatus -eq $pimStatus) { return "ok" }
    return "nicht ok"
}

function Invoke-Check2_Kategorie {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    $pmsCat = $Data[$R.PMS_SLLCAT]
    $pimCat = $Data[$R.PIM_Kategorie]
    if ($pmsCat -eq "UKN" -and [string]::IsNullOrWhiteSpace($pimCat)) { return "ok - Kein Kat-Mapping im PMS und PIM" }
    if ($pmsCat -eq $pimCat) { return "ok" }
    return "nicht ok"
}

function Invoke-Check3_Genre {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq 'passive') { return "ok - Status = passive" }
    $pmsGenresRaw = $Data[$R.PMS_SLLGNR]
    $pimGenre = $Data[$R.PIM_Genre]
    if (-not [string]::IsNullOrEmpty($pmsGenresRaw) -and $pmsGenresRaw -notlike '*' -and [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok - Kein Genre im PIM vorhanden" }
    if ([string]::IsNullOrEmpty($pmsGenresRaw) -and [string]::IsNullOrEmpty($pimGenre)) { return "ok" }
    if ($pmsGenresRaw -like '*' -and [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok - Genre fehlt im PIM" }
    if ([string]::IsNullOrEmpty($pmsGenresRaw) -or [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok" }
    $pmsGenresClean = $pmsGenresRaw.Trim('[]')
    $pmsGenresArray = $pmsGenresClean.Split(',') | ForEach-Object { $_.Trim() }
    if ($pmsGenresArray -contains $pimGenre) { return "ok" }
    return "nicht ok"
}

function Invoke-Check4_Preiscode {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    if ($Data[$R.PMS_SLLPCD] -eq $Data[$R.PIM_Preiscode]) { return "ok" }
    return "nicht ok"
}

function Invoke-Check5_StandardVP {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    $pms = $Data[$R.PMS_FLGSTP]
    $pim = $Data[$R.PIM_StandardVP]
    $pimNull = [string]::IsNullOrEmpty($pim)
    $pmsNull = [string]::IsNullOrEmpty($pms)
    
    if ($pms -eq '1' -and -not $pimNull) { return "ok" }
    if ($pms -eq $pim) { return "ok" }
    if ($pimNull -and $pms -eq '0') { return "ok" }
    if ($pimNull -and $pms -ne '0') { return "nicht ok - Standard VP nur im PMS" }
    if (($pms -eq '0' -or $pmsNull) -and -not $pimNull) { return "nicht ok - Standard VP nur im PIM" }
    if (-not $pmsNull -and -not $pimNull -and $pms -ne $pim) { return "nicht ok - Werte unterschiedlich (PMS: '$($pms)', PIM: '$($pim)')" }
    return "nicht ok - (undefinierter Unterschied)"
}

function Invoke-Check6_FixerVP {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    $pms = $Data[$R.PMS_FLGFXP]
    $pim = $Data[$R.PIM_FixerVP]
    if ($pms -eq '1') { return "ok" }
    $pimNull = [string]::IsNullOrEmpty($pim)
    $pmsNull = [string]::IsNullOrEmpty($pms)
    if ($pimNull -and $pms -eq '0') { return "ok" }
    if ($pms -eq '0' -and $pim -eq '0') { return "ok" }
    if ($pimNull -and $pmsNull) { return "ok" }
    if (($pms -eq '0' -or $pmsNull) -and -not $pimNull -and $pim -ne '0') { return "nicht ok - Fixer VP nur im PIM ($($pim)), PMS ist 0/leer" }
    return "ok"
}

function Invoke-Check7_ReleaseDatum {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }

    $pmsDateString = $Data[$R.PMS_RELDAT]
    $pimDateString = $Data[$R.PIM_ReleaseDate]

    if ($null -ne $pimDateString) {
        $pimDateString = $pimDateString.Trim()
        if ($pimDateString -match '^(0+)$') { $pimDateString = '' }
    }

    $pmsValid = (-not [string]::IsNullOrWhiteSpace($pmsDateString)) -and $pmsDateString -ne '0'
    $pimValid = -not [string]::IsNullOrEmpty($pimDateString)
    if (-not $pmsValid -and -not $pimValid) { return "ok" }
    if ($pmsValid -and -not $pimValid) { return "nicht ok - Datum fehlt im PIM" }
    if (-not $pmsValid -and $pimValid) { return "nicht ok - Datum fehlt im PMS" }

    try { $pmsDate = [datetime]::ParseExact($pmsDateString.Trim(), 'dd.MM.yy', $null) }
    catch { return "nicht ok - Datumsformat ungueltig im PMS ('$($pmsDateString)')" }

    try {
        $pimDate = [datetime]::ParseExact($pimDateString, 'yyyyMMdd', $null)
    } catch {
        try { $pimDate = [datetime]$pimDateString }
        catch { return "nicht ok - Datumsformat ungueltig im PIM ('$($pimDateString)')" }
    }

    $pmsYear2Digit = $pmsDate.Year % 100
    $pimYear2Digit = $pimDate.Year % 100
    
    if ($pmsDate.Day -eq $pimDate.Day -and 
        $pmsDate.Month -eq $pimDate.Month -and 
        $pmsYear2Digit -eq $pimYear2Digit) {
        return "ok"
    }
    
    $pmsDisplay = $pmsDate.ToString('dd.MM.') + $pmsYear2Digit.ToString('00')
    $pimDisplay = $pimDate.ToString('dd.MM.') + $pimYear2Digit.ToString('00')
    return "nicht ok - Unterschiedliche Daten (PMS: $pmsDisplay, PIM: $pimDisplay)"
}

function Invoke-Check8_Errorcode {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    $pms = $Data[$R.PMS_SLLERR]
    $pim = $Data[$R.PIM_Errorcode]
    $pmsNull = [string]::IsNullOrEmpty($pms)
    $pimNull = [string]::IsNullOrEmpty($pim)
    if ($pmsNull -and $pimNull) { return "ok" }
    if ($pms -eq $pim) { return "ok" }
    return "nicht ok"
}

function Invoke-Check9_VP {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    
    $pmsVP = $Data[$R.PMS_SLLVPL]
    $pimVP = $Data[$R.PIM_VP]
    
    $pmsIsZero = ($pmsVP -eq '0' -or $pmsVP -eq '0.00' -or $pmsVP -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimVP)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    if ($pmsVP -eq $pimVP) { return "ok" }
    
    if (-not [string]::IsNullOrEmpty($pmsVP) -and -not [string]::IsNullOrEmpty($pimVP)) {
        $kategorie = $Data[$R.PIM_Kategorie]
        $standardVP = $Data[$R.PIM_StandardVP]
        
        if (($kategorie -eq 'B' -or $kategorie -eq 'B-EN' -or $kategorie -eq 'B-FR') -and 
            -not [string]::IsNullOrEmpty($standardVP)) {
            return "Warnung - Buch-Kat - VP nicht identisch"
        }
    }
    
    return "nicht ok"
}

function Invoke-Check10_PrioEP {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    
    $pmsRaw = $Data[$R.PMS_SLLEPL]
    $pimRaw = $Data[$R.PIM_PrioEP]
    
    $pmsIsZero = ($pmsRaw -eq '0' -or $pmsRaw -eq '0.00' -or $pmsRaw -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimRaw)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    $absoluteTolerance = [decimal]0.02
    $relativeTolerance = [decimal]0.0001
    
    $pms = [decimal]0
    $pim = [decimal]0
    $ok1 = [decimal]::TryParse($pmsRaw, [ref]$pms)
    $ok2 = [decimal]::TryParse($pimRaw, [ref]$pim)
    
    if (-not ($ok1 -and $ok2)) { return "nicht ok" }
    
    if ($pms -eq $pim) { return "ok" }
    
    $diff = [Math]::Abs($pms - $pim)
    $relativeThreshold = [Math]::Abs($pms) * $relativeTolerance
    
    if ($diff -le $absoluteTolerance -or $diff -le $relativeThreshold) {
        $diffStr = $diff.ToString("0.00####", [System.Globalization.CultureInfo]::InvariantCulture)
        return "ok - Diff von $diffStr"
    }
    
    return "nicht ok"
}

function Invoke-Check11_RgEP {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    
    $pmsRgEP = $Data[$R.PMS_SLOERG]
    $pimRgEP = $Data[$R.PIM_RgEP]
    
    $pmsIsZero = ($pmsRgEP -eq '0' -or $pmsRgEP -eq '0.00' -or $pmsRgEP -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimRgEP)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    if ($pmsRgEP -eq $pimRgEP) { return "ok" }
    return "nicht ok"
}

function Invoke-Check12_Tiefpreis {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    
    $pmsVP = $Data[$R.PMS_SLLVPL]
    $pmsTiefpreis = $Data[$R.PMS_REDVPL]
    $pimTiefpreis = $Data[$R.PIM_Tiefpreis]
    
    $pmsTiefpreisEffective = $pmsTiefpreis
    if (-not [string]::IsNullOrEmpty($pmsVP) -and -not [string]::IsNullOrEmpty($pmsTiefpreis) -and $pmsVP -eq $pmsTiefpreis) {
        $pmsTiefpreisEffective = $null
    }
    
    $pmsIsZero = ($pmsTiefpreisEffective -eq '0' -or $pmsTiefpreisEffective -eq '0.00' -or $pmsTiefpreisEffective -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimTiefpreis)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    $pmsNull = [string]::IsNullOrEmpty($pmsTiefpreisEffective)
    
    if ($pmsNull -and $pimIsEmpty) { return "ok" }
    if ($pmsTiefpreisEffective -eq $pimTiefpreis) { return "ok" }
    
    # Check 9 Ergebnis aus Array holen
    $check9 = $Data[$R.Check9]
    $redvpl = 0
    $sllvpl = 0
    $redvplParsed = [decimal]::TryParse($Data[$R.PMS_REDVPL], [ref]$redvpl)
    $sllvplParsed = [decimal]::TryParse($Data[$R.PMS_SLLVPL], [ref]$sllvpl)
    
    if ($check9 -eq "ok" -and 
        $redvplParsed -and $sllvplParsed -and $redvpl -lt $sllvpl -and 
        $Data[$R.PMS_SAASEL] -eq '0') {
        return "Warnung: Anderer Lf is priorisiert; TP wahrscheinlich vom anderen Lf"
    }
    
    return "nicht ok"
}

function Invoke-Check13_LPrioFehlercode {
    param([string[]]$Data)
    $R = $script:RES_IDX
    if ($Data[$R.PMS_SLLPAS] -eq "passive") { return "ok - Status = passive" }
    
    $pmsSaapnt = $Data[$R.PMS_SAAPNT]
    $pimFehlercode = $Data[$R.PIM_Fehlercode]
    
    $pmsInt = 0
    $pimInt = 0
    $pmsIsNum = [long]::TryParse($pmsSaapnt, [ref]$pmsInt)
    $pimIsNum = [long]::TryParse($pimFehlercode, [ref]$pimInt)
    
    $cond1 = [string]::IsNullOrEmpty($pimFehlercode) -and $pmsIsNum -and $pmsInt -lt 900000
    $cond2 = $pmsIsNum -and $pimIsNum -and ($pmsInt -eq $pimInt)
    
    if ($cond1 -or $cond2) { return "ok" }
    return "nicht ok"
}

function Invoke-Check14_LPrio {
    param([string[]]$Data)
    $R = $script:RES_IDX
    
    if ($Data[$R.PMS_SLLPAS] -eq "passive") {
        return "ok - Status = passive"
    }
    
    $pmsSaapnt = 0
    $pimFehlercode = 0
    $pmsParsed = [long]::TryParse($Data[$R.PMS_SAAPNT], [ref]$pmsSaapnt)
    $pimParsed = [long]::TryParse($Data[$R.PIM_Fehlercode], [ref]$pimFehlercode)
    
    if ($pmsParsed -and $pmsSaapnt -ge 900000 -and $pimParsed -and $pmsSaapnt -eq $pimFehlercode) {
        return "ok - Fehlercode identisch"
    }
    
    if ([string]::IsNullOrEmpty($Data[$R.PIM_LPrioPunkte]) -and 
        $Data[$R.Check13] -eq "Warnung - Title fehlt im PIM (Title Tag wahrscheinlich leer)") {
        return "Warnung - Title fehlt im PIM (Title Tag wahrscheinlich leer)"
    }
    
    if ($Data[$R.PMS_SAAPNT] -eq $Data[$R.PIM_LPrioPunkte]) {
        return "ok"
    }
    
    if ($Data[$R.Check9] -eq "Warnung - Buch-Kat - VP nicht identisch") {
        return "Warnung - Buch-Kat - VP nicht identisch"
    }
    
    # Korrelation mit PrioEP-Diff
    $prioEPDiffValue = $Data[$R.PrioEPDiff]
    if (-not [string]::IsNullOrEmpty($prioEPDiffValue) -and $prioEPDiffValue -ne '' -and $prioEPDiffValue -ne '---') {
        $prioEPDiff = [decimal]0
        $prioEPDiffParsed = [decimal]::TryParse($prioEPDiffValue, [ref]$prioEPDiff)
        
        if ($prioEPDiffParsed -and $prioEPDiff -ne 0) {
            $pmsLPrio = 0
            $pimLPrio = 0
            $pmsLPrioParsed = [long]::TryParse($Data[$R.PMS_SAAPNT], [ref]$pmsLPrio)
            $pimLPrioParsed = [long]::TryParse($Data[$R.PIM_LPrioPunkte], [ref]$pimLPrio)
            
            if ($pmsLPrioParsed -and $pimLPrioParsed) {
                $lPrioDiff = [Math]::Abs($pmsLPrio - $pimLPrio)
                $prioEPDiffAbs = [Math]::Abs($prioEPDiff)
                
                if ($lPrioDiff -lt (250 * $prioEPDiffAbs)) {
                    return "Warnung - L-Prio-Diff vorhanden. Kommt wahrscheinlich von PrioEP-Diff"
                }
            }
        }
    }
    
    return "nicht ok"
}

function Invoke-Check13_Extended {
    param([string[]]$Data)
    $R = $script:RES_IDX
    
    if ($Data[$R.Check13] -ne 'nicht ok') {
        return $Data[$R.Check13]
    }
    
    $checks1to12AllOk = (
        ($Data[$R.Check1] -like 'ok*') -and
        ($Data[$R.Check2] -like 'ok*') -and
        ($Data[$R.Check3] -like 'ok*') -and
        ($Data[$R.Check4] -like 'ok*') -and
        ($Data[$R.Check5] -like 'ok*') -and
        ($Data[$R.Check6] -like 'ok*') -and
        ($Data[$R.Check7] -like 'ok*') -and
        ($Data[$R.Check8] -like 'ok*') -and
        ($Data[$R.Check9] -like 'ok*' -or $Data[$R.Check9] -like 'Warnung*') -and
        ($Data[$R.Check10] -like 'ok*') -and
        ($Data[$R.Check11] -like 'ok*') -and
        ($Data[$R.Check12] -like 'ok*')
    )
    
    if ($checks1to12AllOk) {
        $pmsSaapnt = 0
        $pimFehlercode = 0
        $pmsParsed = [long]::TryParse($Data[$R.PMS_SAAPNT], [ref]$pmsSaapnt)
        $pimParsed = [long]::TryParse($Data[$R.PIM_Fehlercode], [ref]$pimFehlercode)
        
        if ($pmsParsed -and $pmsSaapnt -lt 900000 -and $pimParsed -and $pimFehlercode -eq 999914) {
            return "Warnung - Title fehlt im PIM (Title Tag wahrscheinlich leer)"
        }
    }
    
    return "nicht ok"
}

# =====================================================================
# HELPER: ZeitDifferenz berechnen
# =====================================================================
function Invoke-CalculateTimeDifference {
    param([string[]]$Data)
    $R = $script:RES_IDX
    
    $pmsDateString = $Data[$R.PMS_CHGDAT]
    $pimDateString = $Data[$R.PIM_LetzteAenderung]
    
    if ([string]::IsNullOrWhiteSpace($pmsDateString) -or [string]::IsNullOrWhiteSpace($pimDateString)) {
        return "fehlende Daten"
    }
    
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $pmsDateTime = $null
    
    try {
        $pmsDateTime = [datetime]::ParseExact("$pmsDateString 12:00:00", "dd.MM.yy HH:mm:ss", $culture)
    } catch {
        return "PMS-Datum ungültig: '$pmsDateString'"
    }
    
    $pimDateTime = $null
    $trimmedPimString = $pimDateString.Trim()
    
    try {
        $pimDateTime = [datetime]$trimmedPimString
    } catch {
        try {
            $sanitized = $trimmedPimString -replace '[–—]', '-' -replace '\s+', ' '
            $formats = @("yyyy-MM-dd HH:mm:ss", "dd.MM.yyyy HH:mm:ss")
            $pimDateTime = [datetime]::ParseExact($sanitized, $formats, $culture, [System.Globalization.DateTimeStyles]::None)
        } catch {
            $sanitizedForError = $pimDateString.Trim() -replace '[–—]', '-' -replace '\s+', ' '
            return "PIM-Datum unlesbar. Original: '$($pimDateString)', Bereinigt: '$($sanitizedForError)'"
        }
    }
    
    $timeSpan = $pimDateTime - $pmsDateTime
    return [Math]::Round($timeSpan.TotalHours, 2).ToString()
}
