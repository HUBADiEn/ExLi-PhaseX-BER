<#
.SYNOPSIS
    Allgemeine Hilfsfunktionen für PMS/PIM Vergleich

.NOTES
    File:           functions-helpers_v1.0.ps1
    Version:        1.0
    Änderungshistorie:
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Helpers = "1.0"

function Invoke-CalculateTimeDifference {
    param([Parameter(Mandatory=$true)][PSCustomObject]$Dataset)
    
    $pmsDateString = $Dataset.PMS_CHGDAT
    $pimDateString = $Dataset.'PIM_letzte Änderung'
    
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
    return [Math]::Round($timeSpan.TotalHours, 2)
}
