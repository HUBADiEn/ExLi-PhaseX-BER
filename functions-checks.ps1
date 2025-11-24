<#
.SYNOPSIS
    Check-Funktionen (Check 1-14) für PMS/PIM Vergleich

.NOTES
    File:           functions-checks.ps1
    Version:        1.0
    Änderungshistorie:
        1.0 - Initiale Version (aus V1.103 extrahiert)
            - Check 12 (Tiefpreis): Warnung für anderen priorisierten Lieferanten
            - Bedingung: nicht ok UND Check 9 = ok UND REDVPL < SLLVPL UND SAASEL = 0
#>

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
    catch { return "nicht ok - Datumsformat ungültig im PMS ('$($pmsDateString)')" }

    try {
        $pimDate = [datetime]::ParseExact($pimDateString, 'yyyyMMdd', $null)
    } catch {
        try { $pimDate = [datetime]$pimDateString }
        catch { return "nicht ok - Datumsformat ungültig im PIM ('$($pimDateString)')" }
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
    
    # Wenn VPs identisch sind
    if ($pmsVP -eq $pimVP) { return "ok" }
    
    # Spezialfall Buch-Kategorien mit Standard VP
    if (-not [string]::IsNullOrEmpty($pmsVP) -and -not [string]::IsNullOrEmpty($pimVP)) {
        $kategorie = $Dataset.PIM_Kategorie
        $standardVP = $Dataset.'PIM_Standard VP'
        
        # Prüfe ob Buch-Kategorie (B, B-EN, B-FR) UND Standard VP vorhanden
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
    $tol = [decimal]0.02
    $pms = [decimal]0
    $pim = [decimal]0
    $ok1 = [decimal]::TryParse($Dataset.PMS_SLLEPL, [ref]$pms)
    $ok2 = [decimal]::TryParse($Dataset.PIM_PrioEP, [ref]$pim)
    if (-not ($ok1 -and $ok2)) { return "nicht ok" }
    $diff = [Math]::Abs($pms - $pim)
    if ($diff -le $tol) {
        $diffStr = $diff.ToString("0.00##", [System.Globalization.CultureInfo]::InvariantCulture)
        return "ok - Diff von $diffStr"
    }
    if ($pms -eq $pim) { "ok" } else { "nicht ok" }
}

function Invoke-Check11_RgEP {
    param([PSCustomObject]$Dataset)
    if ($Dataset.PMS_SLLPAS -eq "passive") { return "ok - Status = passive" }
    if ($Dataset.PMS_SLOEPF -eq $Dataset.PIM_RgEP) { "ok" } else { "nicht ok" }
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
    
    $pmsNull = [string]::IsNullOrEmpty($pmsTiefpreisEffective)
    $pimNull = [string]::IsNullOrEmpty($pimTiefpreis)
    
    if ($pmsNull -and $pimNull) { return "ok" }
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
    
    # Wenn Status = passive, dann Check überspringen
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
    
    return "nicht ok"
}

# Erweiterte Check 13 Logik (wird nach Check 1-12 aufgerufen)
function Invoke-Check13_Extended {
    param([PSCustomObject]$Dataset)
    
    # Nur ausführen wenn Check 13 "nicht ok" ist
    if ($Dataset.'Check 13: L-Prio Fehlercode' -ne 'nicht ok') {
        return $Dataset.'Check 13: L-Prio Fehlercode'
    }
    
    # Prüfe ob Check 1-12 alle ok waren
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
        # Prüfe Title-Warnung Bedingungen
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
