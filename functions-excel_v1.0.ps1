<#
.SYNOPSIS
    Excel-Hilfsfunktionen für PMS/PIM Vergleich

.NOTES
    File:           functions-excel_v1.0.ps1
    Version:        1.0
    Änderungshistorie:
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Excel = "1.0"

function Get-ExcelColumnName {
    param([Parameter(Mandatory=$true)][int]$ColumnNumber)
    $div = $ColumnNumber
    $name = ""
    while ($div -gt 0) {
        $mod = ($div - 1) % 26
        $name = [char](65 + $mod) + $name
        $div = [math]::Floor(($div - $mod) / 26)
    }
    return $name
}

function Optimize-ColumnWidthForHeader {
    param(
        [string]$Path,
        [string]$WorksheetName
    )
    try {
        if (-not (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) { return }
        $excel = Open-ExcelPackage -Path $Path
        try {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws -and $ws.Dimension) {
                # Header in Zeile 3
                $range = $ws.Cells[3, 1, 3, $ws.Dimension.End.Column]
                $null = $range.AutoFitColumns()
                $ws.Column(1).AutoFit()
            }
        } finally {
            Close-ExcelPackage $excel
        }
    } catch { }
}

function Apply-SummaryRow {
    param(
        [string]$Path,
        [string]$WorksheetName,
        [PSCustomObject]$HeaderSummary,
        [PSCustomObject]$WarningSummary,
        [string]$ScriptVersion,
        [string]$SupplierNumber
    )
    try {
        if (-not (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) { return }
        $excel = Open-ExcelPackage -Path $Path
        try {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws -and $ws.Dimension) {
                $lastCol = $ws.Dimension.End.Column
                $propsError = $HeaderSummary.PSObject.Properties
                $propsWarning = $WarningSummary.PSObject.Properties
                
                # A1 = Lieferantennummer (keine Färbung)
                $ws.Cells[1, 1].Value = $SupplierNumber
                
                # A2 = Script Version ohne Prefix (keine Färbung)
                $ws.Cells[2, 1].Value = $ScriptVersion
                
                # Zeile 1: Warnungen (Hell-Orange Hintergrund nur wenn > 0)
                for ($col = 2; $col -le $lastCol; $col++) {
                    $headerText = $ws.Cells[3, $col].Text
                    if ([string]::IsNullOrEmpty($headerText)) { continue }
                    $p = $propsWarning | Where-Object { $_.Name -eq $headerText }
                    if ($p) {
                        $cell = $ws.Cells[1, $col]
                        $cell.Value = $p.Value
                        $intVal = 0
                        if ([int]::TryParse($p.Value, [ref]$intVal) -and $intVal -gt 0) {
                            $cell.Style.Fill.PatternType = 'Solid'
                            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFD699"))
                        }
                    }
                }
                
                # Zeile 2: Fehler (Rot Hintergrund nur wenn > 0)
                for ($col = 2; $col -le $lastCol; $col++) {
                    $headerText = $ws.Cells[3, $col].Text
                    if ([string]::IsNullOrEmpty($headerText)) { continue }
                    $p = $propsError | Where-Object { $_.Name -eq $headerText }
                    if ($p) {
                        $cell = $ws.Cells[2, $col]
                        $cell.Value = $p.Value
                        $intVal = 0
                        if ([int]::TryParse($p.Value, [ref]$intVal) -and $intVal -gt 0) {
                            $cell.Style.Fill.PatternType = 'Solid'
                            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#F8D0D0"))
                        }
                    }
                }
            }
        } finally {
            Close-ExcelPackage $excel
        }
    } catch { }
}

function Color-HeaderBySource {
    param(
        [string]$Path,
        [string]$WorksheetName
    )
    try {
        if (-not (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) { return }
        $excel = Open-ExcelPackage -Path $Path
        try {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws -and $ws.Dimension) {
                $lastCol = $ws.Dimension.End.Column
                # Header in Zeile 3
                for ($col = 1; $col -le $lastCol; $col++) {
                    $headerText = $ws.Cells[3, $col].Text
                    $cell = $ws.Cells[3, $col]
                    if ($headerText -like 'PMS_*') {
                        $cell.Style.Fill.PatternType = 'Solid'
                        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#DAE9F8"))
                    } elseif ($headerText -like 'PIM_*') {
                        $cell.Style.Fill.PatternType = 'Solid'
                        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#F2CEEF"))
                    }
                }
                
                $eanCol = $ws.Cells["A:A"]
                $eanCol.Style.Numberformat.Format = "@"
            }
        } finally {
            Close-ExcelPackage $excel
        }
    } catch { }
}
