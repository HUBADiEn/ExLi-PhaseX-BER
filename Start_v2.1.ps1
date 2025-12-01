<#
.SYNOPSIS
    Start-Script für PMS/PIM Vergleich - Prüft Modulversionen und startet Hauptlogik
    PERFORMANCE-OPTIMIERT für 12+ Mio Zeilen

.NOTES
    File:           Start_v2.1.ps1
    Version:        2.1
    Änderungshistorie:
        2.1 - Erwartet main_v2.1 (Fortschrittsausgabe alle 100k Zeilen)
        2.0 - PERFORMANCE: Radikale Optimierung fuer 12+ Mio Zeilen
    
    Benötigte Module:
        - config_v1.3.ps1            (ModuleVersion_Config = 1.3)
        - functions-excel_v1.0.ps1   (ModuleVersion_Excel = 1.0)
        - functions-dialogs_v1.0.ps1 (ModuleVersion_Dialogs = 1.0)
        - functions-checks_v1.7.ps1  (ModuleVersion_Checks = 1.7)
        - main_v2.1.ps1              (ModuleVersion_Main = 2.1)
#>

# =====================================================================
# START-SCRIPT VERSION
# =====================================================================
$script:StartVersion = "2.1"
$global:ScriptVersion = "Berechnung_V$($script:StartVersion)"

# =====================================================================
# BENÖTIGTE MODUL-VERSIONEN
# =====================================================================
$script:RequiredModules = [ordered]@{
    'config'           = @{ File = 'config_v1.3.ps1';            Variable = 'ModuleVersion_Config';  Version = '1.3' }
    'functions-excel'  = @{ File = 'functions-excel_v1.0.ps1';   Variable = 'ModuleVersion_Excel';   Version = '1.0' }
    'functions-dialogs'= @{ File = 'functions-dialogs_v1.0.ps1'; Variable = 'ModuleVersion_Dialogs'; Version = '1.0' }
    'functions-checks' = @{ File = 'functions-checks_v1.7.ps1';  Variable = 'ModuleVersion_Checks';  Version = '1.7' }
    'main'             = @{ File = 'main_v2.1.ps1';              Variable = 'ModuleVersion_Main';    Version = '2.1' }
}

# =====================================================================
# BOOTSTRAP (identisch mit v1.14 das funktioniert)
# =====================================================================
if (-not $env:PS_KEEP_NOEXIT) {
    try {
        $env:PS_KEEP_NOEXIT = '1'
        $scriptPath = $MyInvocation.MyCommand.Definition
        if (-not (Test-Path -LiteralPath $scriptPath)) { throw "Scriptpfad ungueltig" }
        $quoted = '"' + $scriptPath.Replace('"', '""') + '"'
        # /c statt /k damit Fenster nach ENTER automatisch schliesst
        $arguments = '/c powershell.exe -NoLogo -ExecutionPolicy Bypass -File ' + $quoted + ' & pause'
        Start-Process -FilePath 'cmd.exe' -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $scriptPath)
    } catch {
        $global:__ForcePauseAtEnd = $true
    }
    return
}

$global:__ForcePauseAtEnd = $false
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms

# =====================================================================
# HAUPTTEIL
# =====================================================================
$scriptSuccessfullyCompleted = $false
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

try {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "PMS/PIM Vergleich - Start" -ForegroundColor Cyan
    Write-Host "Version: $($global:ScriptVersion)" -ForegroundColor Cyan
    Write-Host "PERFORMANCE-OPTIMIERT (12+ Mio Zeilen)" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    
    Write-Host ""
    Write-Host "Pruefe Modul-Versionen..." -ForegroundColor Cyan
    Write-Host ""
    
    $versionErrors = @()
    
    foreach ($moduleName in $script:RequiredModules.Keys) {
        $moduleInfo = $script:RequiredModules[$moduleName]
        $filePath = Join-Path $ScriptDir $moduleInfo.File
        $requiredVersion = $moduleInfo.Version
        $versionVariable = $moduleInfo.Variable
        
        if (-not (Test-Path $filePath)) {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' nicht gefunden!"
            Write-Host "  [X] $($moduleInfo.File) - NICHT GEFUNDEN (Pfad: $filePath)" -ForegroundColor Red
            continue
        }
        
        try {
            . $filePath
        } catch {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' konnte nicht geladen werden: $($_.Exception.Message)"
            Write-Host "  [X] $($moduleInfo.File) - LADEFEHLER: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
        
        $actualVersion = Get-Variable -Name $versionVariable -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        
        if (-not $actualVersion) {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' enthaelt keine Versionsvariable '$versionVariable'!"
            Write-Host "  [X] $($moduleInfo.File) - KEINE VERSION GEFUNDEN" -ForegroundColor Red
            continue
        }
        
        if ($actualVersion -ne $requiredVersion) {
            $versionErrors += "FEHLER: Modul '$($moduleInfo.File)' hat Version $actualVersion, benoetigt wird $requiredVersion!"
            Write-Host "  [X] $($moduleInfo.File) - Version $actualVersion (benoetigt: $requiredVersion)" -ForegroundColor Red
            continue
        }
        
        Write-Host "  [OK] $($moduleInfo.File) - Version $actualVersion" -ForegroundColor Green
    }
    
    Write-Host ""
    
    if ($versionErrors.Count -gt 0) {
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "VERSIONSPRUEFUNG FEHLGESCHLAGEN!" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host ""
        foreach ($err in $versionErrors) {
            Write-Host $err -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "Benoetigte Module:" -ForegroundColor Cyan
        foreach ($moduleName in $script:RequiredModules.Keys) {
            $moduleInfo = $script:RequiredModules[$moduleName]
            Write-Host "  - $($moduleInfo.File) (Version $($moduleInfo.Version))" -ForegroundColor White
        }
        Write-Host ""
        Write-Host "Script wird beendet." -ForegroundColor Red
    } else {
        Write-Host "Alle Module erfolgreich geladen und verifiziert." -ForegroundColor Green
        Write-Host ""
        Write-Host "Starte Hauptlogik..." -ForegroundColor Cyan
        Write-Host ""
        
        Invoke-MainLogic
        $scriptSuccessfullyCompleted = $true
    }
}
catch {
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    Write-Host "EIN KRITISCHER FEHLER IST AUFGETRETEN:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Kritischer Fehler", "OK", "Error") | Out-Null
}
finally {
    if (-not $scriptSuccessfullyCompleted) {
        Write-Host ""
        if ($global:__ForcePauseAtEnd) { 
            Write-Host "Hinweis: Relaunch mit eigenem Fenster war nicht moeglich." -ForegroundColor Yellow 
        }
        Write-Host "Druecke ENTER um das Fenster zu schliessen." -ForegroundColor White
        Read-Host
    }
}
