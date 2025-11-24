<#
.SYNOPSIS
    Dialog- und Export-Funktionen für PMS/PIM Vergleich

.NOTES
    File:           functions-dialogs_v1.0.ps1
    Version:        1.0
    Änderungshistorie:
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Dialogs = "1.0"

function Get-FilePathDialog {
    param(
        [string]$WindowTitle,
        [string]$InitialDirectory
    )
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $WindowTitle
    $dlg.Filter = "CSV-Dateien (*.csv)|*.csv"
    $dlg.InitialDirectory = $InitialDirectory
    if ($dlg.ShowDialog() -eq 'OK') {
        return $dlg.FileName
    }
    return $null
}

function Export-CsvFast {
    param(
        [Parameter(Mandatory=$true)]$Data,
        [Parameter(Mandatory=$true)][string]$Path,
        [string]$Delimiter = ';'
    )
    # Optimierung: StreamWriter für große Datasets (>100k Zeilen)
    if ($Data.Count -lt 100000) {
        $Data | Export-Csv -Path $Path -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
        return
    }
    
    Write-Host "      (Nutze optimierten StreamWriter-Export für $($Data.Count) Zeilen)" -ForegroundColor Cyan
    $writer = $null
    try {
        $writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.Encoding]::UTF8)
        
        # Header
        if ($Data.Count -gt 0) {
            $props = $Data[0].PSObject.Properties.Name
            $writer.WriteLine(($props -join $Delimiter))
            
            # Daten
            foreach ($item in $Data) {
                $values = foreach ($prop in $props) {
                    $val = $item.$prop
                    if ($null -eq $val) { '' }
                    else { $val.ToString() -replace $Delimiter, '_' }
                }
                $writer.WriteLine(($values -join $Delimiter))
            }
        }
    } finally {
        if ($writer) {
            $writer.Close()
            $writer.Dispose()
        }
    }
}
