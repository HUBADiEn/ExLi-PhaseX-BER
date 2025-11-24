<#
.SYNOPSIS
    Check-Funktionen (Check 1-14) für PMS/PIM Vergleich

.NOTES
    File:           functions-checks_v1.3.ps1
    Version:        1.3
    Änderungshistorie:
        1.3 - Check 14 (L-Prio): Korrelation mit PrioEP-Diff
            - Wenn L-Prio Diff vorhanden UND PrioEP Diff vorhanden
            - UND |L-Prio Diff| < 250 * |PrioEP Diff|
            - Dann: Warnung statt Fehler (L-Prio-Diff kommt wahrscheinlich von PrioEP-Diff)
        1.2 - Checks 9, 10, 11, 12: PMS-Wert "0" entspricht leerem PIM-Feld
        1.1 - Check 10 (PrioEP): Relative Toleranz hinzugefuegt
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Checks = "1.3"

function Invoke-Check1_Status {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq $Dataset.PIM_Status) { "ok" } else { "nicht ok" }
}

function Invoke-Check2_Kategorie {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    $pmsCat = $Dataset.PMS_SLLCAT
    $pimCat = $Dataset.PIM_Kategorie
    if ($pmsCat -eq "UKN" -and [string]::IsNullOrWhiteSpace($pimCat)) { return "ok - Kein Kat-Mapping im PMS und PIM" }
    if ($pmsCat -eq $pimCat) { "ok" } else { "nicht ok" }
}

function Invoke-Check3_Genre {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq 'passive') { return "ok - Status = passive" }
    $pmsGenresRaw = $Dataset.PMS_SLLGNR
    $pimGenre = $Dataset.PIM_Genre
    if (-not [string]::IsNullOrEmpty($pmsGenresRaw) -and $pmsGenresRaw -notlike '*' -and [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok - Kein Genre im PIM vorhanden" }
    if ([string]::IsNullOrEmpty($pmsGenresRaw) -and [string]::IsNullOrEmpty($pimGenre)) { return "ok" }
    if ($pmsGenresRaw -like '*' -and [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok - Genre fehlt im PIM" }
    if ([string]::IsNullOrEmpty($pmsGenresRaw) -or [string]::IsNullOrEmpty($pimGenre)) { return "nicht ok" }
    $pmsGenresClean = $pmsGenresRaw.Trim('[]')
    $pmsGenresArray = $pmsGenresClean.Split(',') | ForEach-Object { $_.Trim() }
    if ($pmsGenresArray -contains $pimGenre) { "ok" } else { "nicht ok" }
}

function Invoke-Check4_Preiscode {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    if ($Dataset.PMS_SLLPCD -eq $Dataset.PIM_Preiscode) { "ok" } else { "nicht ok" }
}

function Invoke-Check5_StandardVP {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    $pms = $Dataset.PMS_FLGSTP
    $pim = $Dataset.'PIM_Standard VP'
    $pimNull = [string]::IsNullOrEmpty($pim)
    $pmsNull = [string]::IsNullOrEmpty($pms)
    
    # Beliebiger Wert im PIM = "1" im PMS
    if ($pms -eq '1' -and -not $pimNull) { return "ok" }
    
    if ($pms -eq $pim) { return "ok" }
    if ($pimNull -and $pms -eq '0') { return "ok" }
    if ($pimNull -and $pms -ne '0') { return "nicht ok - Standard VP nur im PMS" }
    if (($pms -eq '0' -or $pmsNull) -and -not $pimNull) { return "nicht ok - Standard VP nur im PIM" }
    if (-not $pmsNull -and -not $pimNull -and $pms -ne $pim) { return "nicht ok - Werte unterschiedlich (PMS: '$($pms)', PIM: '$($pim)')" }
    return "nicht ok - (undefinierter Unterschied)"
}

function Invoke-Check6_FixerVP {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    $pms = $Dataset.PMS_FLGFXP
    $pim = $Dataset.'PIM_Fixer VP'
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
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }

    $pmsDateString = $Dataset.PMS_RELDAT
    $pimDateString = $Dataset.'PIM_Release Date'

    # PIM-"0"/Nullfolgen wie leer behandeln
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

    if ($pmsDate.Date -eq $pimDate.Date) { "ok" }
    else { "nicht ok - Unterschiedliche Daten (PMS: $($pmsDate.ToString('dd.MM.yyyy')), PIM: $($pimDate.ToString('dd.MM.yyyy')))" }
}

function Invoke-Check8_Errorcode {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    $pms = $Dataset.PMS_SLLERR
    $pim = $Dataset.PIM_Errorcode
    $pmsNull = [string]::IsNullOrEmpty($pms)
    $pimNull = [string]::IsNullOrEmpty($pim)
    if ($pmsNull -and $pimNull) { return "ok" }
    if ($pms -eq $pim) { return "ok" }
    return "nicht ok"
}

function Invoke-Check9_VP {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    
    $pmsVP = $Dataset.PMS_SLLVPL
    $pimVP = $Dataset.PIM_VP
    
    # V1.2: PMS "0" entspricht leerem PIM-Feld
    $pmsIsZero = ($pmsVP -eq '0' -or $pmsVP -eq '0.00' -or $pmsVP -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimVP)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    # Wenn VPs identisch sind
    if ($pmsVP -eq $pimVP) { return "ok" }
    
    # Spezialfall Buch-Kategorien mit Standard VP
    if (-not [string]::IsNullOrEmpty($pmsVP) -and -not [string]::IsNullOrEmpty($pimVP)) {
        $kategorie = $Dataset.PIM_Kategorie
        $standardVP = $Dataset.'PIM_Standard VP'
        
        # Pruefe ob Buch-Kategorie (B, B-EN, B-FR) UND Standard VP vorhanden
        if (($kategorie -eq 'B' -or $kategorie -eq 'B-EN' -or $kategorie -eq 'B-FR') -and 
            -not [string]::IsNullOrEmpty($standardVP)) {
            return "Warnung - Buch-Kat - VP nicht identisch"
        }
    }
    
    # Standard: nicht ok
    return "nicht ok"
}

function Invoke-Check10_PrioEP {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    
    $pmsRaw = $Dataset.PMS_SLLEPL
    $pimRaw = $Dataset.PIM_PrioEP
    
    # V1.2: PMS "0" entspricht leerem PIM-Feld
    $pmsIsZero = ($pmsRaw -eq '0' -or $pmsRaw -eq '0.00' -or $pmsRaw -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimRaw)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    # Toleranzen
    $absoluteTolerance = [decimal]0.02           # Absolute Toleranz: 0.02
    $relativeTolerance = [decimal]0.0001         # Relative Toleranz: 0.01% vom PMS-Wert
    
    $pms = [decimal]0
    $pim = [decimal]0
    $ok1 = [decimal]::TryParse($pmsRaw, [ref]$pms)
    $ok2 = [decimal]::TryParse($pimRaw, [ref]$pim)
    
    if (-not ($ok1 -and $ok2)) { return "nicht ok" }
    
    # Exakt gleich
    if ($pms -eq $pim) { return "ok" }
    
    $diff = [Math]::Abs($pms - $pim)
    
    # V1.1: Berechne relative Toleranz basierend auf PMS-Wert
    $relativeThreshold = [Math]::Abs($pms) * $relativeTolerance
    
    # OK wenn absolute ODER relative Toleranz erfuellt
    if ($diff -le $absoluteTolerance -or $diff -le $relativeThreshold) {
        $diffStr = $diff.ToString("0.00####", [System.Globalization.CultureInfo]::InvariantCulture)
        return "ok - Diff von $diffStr"
    }
    
    return "nicht ok"
}

function Invoke-Check11_RgEP {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    
    $pmsRgEP = $Dataset.PMS_SLOEPF
    $pimRgEP = $Dataset.PIM_RgEP
    
    # V1.2: PMS "0" entspricht leerem PIM-Feld
    $pmsIsZero = ($pmsRgEP -eq '0' -or $pmsRgEP -eq '0.00' -or $pmsRgEP -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimRgEP)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    if ($pmsRgEP -eq $pimRgEP) { "ok" } else { "nicht ok" }
}

function Invoke-Check12_Tiefpreis {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    
    $pmsVP = $Dataset.PMS_SLLVPL
    $pmsTiefpreis = $Dataset.PMS_REDVPL
    $pimTiefpreis = $Dataset.PIM_Tiefpreis
    
    # Wenn VP und Tiefpreis im PMS identisch sind, behandle als fehlenden Tiefpreis
    $pmsTiefpreisEffective = $pmsTiefpreis
    if (-not [string]::IsNullOrEmpty($pmsVP) -and -not [string]::IsNullOrEmpty($pmsTiefpreis) -and $pmsVP -eq $pmsTiefpreis) {
        $pmsTiefpreisEffective = $null
    }
    
    # V1.2: PMS "0" entspricht leerem PIM-Feld
    $pmsIsZero = ($pmsTiefpreisEffective -eq '0' -or $pmsTiefpreisEffective -eq '0.00' -or $pmsTiefpreisEffective -eq '0,00')
    $pimIsEmpty = [string]::IsNullOrEmpty($pimTiefpreis)
    if ($pmsIsZero -and $pimIsEmpty) { return "ok" }
    
    $pmsNull = [string]::IsNullOrEmpty($pmsTiefpreisEffective)
    
    if ($pmsNull -and $pimIsEmpty) { return "ok" }
    if ($pmsTiefpreisEffective -eq $pimTiefpreis) { return "ok" }
    
    # V1.103: Warnung wenn anderer Lieferant priorisiert
    # Bedingungen: nicht ok UND Check 9 = ok UND REDVPL < SLLVPL UND SAASEL = 0
    $check9 = $Dataset.'Check 9: VP'
    $redvpl = 0
    $sllvpl = 0
    $redvplParsed = [decimal]::TryParse($Dataset.PMS_REDVPL, [ref]$redvpl)
    $sllvplParsed = [decimal]::TryParse($Dataset.PMS_SLLVPL, [ref]$sllvpl)
    
    if ($check9 -eq "ok" -and 
        $redvplParsed -and $sllvplParsed -and $redvpl -lt $sllvpl -and 
        $Dataset.PMS_SAASEL -eq '0') {
        return "Warnung: Anderer Lf is priorisiert; TP wahrscheinlich vom anderen Lf"
    }
    
    return "nicht ok"
}

function Invoke-Check13_LPrioFehlercode {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    
    $pmsSaapnt = $Dataset.PMS_SAAPNT
    $pimFehlercode = $Dataset.PIM_Fehlercode
    
    # Versuche beide Werte als Zahlen zu parsen
    $pmsInt = 0
    $pimInt = 0
    $pmsIsNum = [long]::TryParse($pmsSaapnt, [ref]$pmsInt)
    $pimIsNum = [long]::TryParse($pimFehlercode, [ref]$pimInt)
    
    # OK-Bedingung 1: PIM Fehlercode ist leer UND PMS SAAPNT < 900000
    $cond1 = [string]::IsNullOrEmpty($pimFehlercode) -and $pmsIsNum -and $pmsInt -lt 900000
    
    # OK-Bedingung 2: Beide sind Zahlen UND identisch
    $cond2 = $pmsIsNum -and $pimIsNum -and ($pmsInt -eq $pimInt)
    
    if ($cond1 -or $cond2) { "ok" } else { "nicht ok" }
}

function Invoke-Check14_LPrio {
    param([PSCustomObject]$Dataset)
    
    # Wenn Status = passive, dann Check ueberspringen
    if ($Dataset.PMS_SLLPAS -eq "passive") {
        return "ok - Status = passive"
    }
    
    # Wenn PMS SAAPNT >= 900000 UND identisch mit PIM Fehlercode
    $pmsSaapnt = 0
    $pimFehlercode = 0
    $pmsParsed = [long]::TryParse($Dataset.PMS_SAAPNT, [ref]$pmsSaapnt)
    $pimParsed = [long]::TryParse($Dataset.PIM_Fehlercode, [ref]$pimFehlercode)
    
    if ($pmsParsed -and $pmsSaapnt -ge 900000 -and $pimParsed -and $pmsSaapnt -eq $pimFehlercode) {
        return "ok - Fehlercode identisch"
    }
    
    # Wenn PIM L-Prio-Punkte leer UND Check 13 hat Title-Warnung
    if ([string]::IsNullOrEmpty($Dataset.'PIM_L-Prio-Punkte') -and 
        $Dataset.'Check 13: L-Prio Fehlercode' -eq "Warnung - Title fehlt im PIM (Title Tag wahrscheinlich leer)") {
        return "Warnung - Title fehlt im PIM (Title Tag wahrscheinlich leer)"
    }
    
    # Normaler Vergleich
    if ($Dataset.PMS_SAAPNT -eq $Dataset.'PIM_L-Prio-Punkte') {
        return "ok"
    }
    
    # Wenn nicht ok UND Check 9 hat Buch-Kat Warnung
    if ($Dataset.'Check 9: VP' -eq "Warnung - Buch-Kat - VP nicht identisch") {
        return "Warnung - Buch-Kat - VP nicht identisch"
    }
    
    # V1.3: Pruefe Korrelation mit PrioEP-Diff
    # Wenn L-Prio Diff vorhanden UND PrioEP Diff vorhanden (Check 10 = nicht ok)
    # UND |L-Prio Diff| < 250 * |PrioEP Diff| -> Warnung
    if ($Dataset.'Check 10: PrioEP' -eq 'nicht ok') {
        # Berechne L-Prio Differenz
        $pmsLPrio = 0
        $pimLPrio = 0
        $pmsLPrioParsed = [long]::TryParse($Dataset.PMS_SAAPNT, [ref]$pmsLPrio)
        $pimLPrioParsed = [long]::TryParse($Dataset.'PIM_L-Prio-Punkte', [ref]$pimLPrio)
        
        # Berechne PrioEP Differenz
        $pmsPrioEP = [decimal]0
        $pimPrioEP = [decimal]0
        $pmsPrioEPParsed = [decimal]::TryParse($Dataset.PMS_SLLEPL, [ref]$pmsPrioEP)
        $pimPrioEPParsed = [decimal]::TryParse($Dataset.PIM_PrioEP, [ref]$pimPrioEP)
        
        if ($pmsLPrioParsed -and $pimLPrioParsed -and $pmsPrioEPParsed -and $pimPrioEPParsed) {
            $lPrioDiff = [Math]::Abs($pmsLPrio - $pimLPrio)
            $prioEPDiff = [Math]::Abs($pmsPrioEP - $pimPrioEP)
            
            # Wenn PrioEP-Diff > 0 und L-Prio-Diff < 250 * PrioEP-Diff
            if ($prioEPDiff -gt 0 -and $lPrioDiff -lt (250 * $prioEPDiff)) {
                return "Warnung - L-Prio-Diff vorhanden. Kommt wahrscheinlich von PrioEP-Diff"
            }
        }
    }
    
    return "nicht ok"
}

# Erweiterte Check 13 Logik (wird nach Check 1-12 aufgerufen)
function Invoke-Check13_Extended {
    param([PSCustomObject]$Dataset)
    
    # Nur ausfuehren wenn Check 13 "nicht ok" ist
    if ($Dataset.'Check 13: L-Prio Fehlercode' -ne 'nicht ok') {
        return $Dataset.'Check 13: L-Prio Fehlercode'
    }
    
    # Pruefe ob Check 1-12 alle ok waren
    $checks1to12AllOk = (
        ($Dataset.'Check 1: Status' -like 'ok*') -and
        ($Dataset.'Check 2: Kategorie' -like 'ok*') -and
        ($Dataset.'Check 3: Genre' -like 'ok*') -and
        ($Dataset.'Check 4: Preiscode' -like 'ok*') -and
        ($Dataset.'Check 5: Standard VP ab Lieferant' -like 'ok*') -and
        ($Dataset.'Check 6: Fixer VP' -like 'ok*') -and
        ($Dataset.'Check 7: Release-Datum' -like 'ok*') -and
        ($Dataset.'Check 8: Errorcode' -like 'ok*') -and
        ($Dataset.'Check 9: VP' -like 'ok*' -or $Dataset.'Check 9: VP' -like 'Warnung*') -and
        ($Dataset.'Check 10: PrioEP' -like 'ok*') -and
        ($Dataset.'Check 11: RgEP' -like 'ok*') -and
        ($Dataset.'Check 12: Tiefpreis' -like 'ok*')
    )
    
    if ($checks1to12AllOk) {
        # Pruefe Title-Warnung Bedingungen
        $pmsSaapnt = 0
        $pimFehlercode = 0
        $pmsParsed = [long]::TryParse($Dataset.PMS_SAAPNT, [ref]$pmsSaapnt)
        $pimParsed = [long]::TryParse($Dataset.PIM_Fehlercode, [ref]$pimFehlercode)
        
        if ($pmsParsed -and $pmsSaapnt -lt 900000 -and $pimParsed -and $pimFehlercode -eq 999914) {
            return "Warnung - Title fehlt im PIM (Title Tag wahrscheinlich leer)"
        }
    }
    
    return "nicht ok"
}
