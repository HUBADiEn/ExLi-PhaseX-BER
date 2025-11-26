<#
.SYNOPSIS
    Konfiguration und Konstanten für PMS/PIM Vergleich

.NOTES
    File:           config_v1.1.ps1
    Version:        1.1
    Änderungshistorie:
        1.1 - FIX: ScriptVersion-Zeile entfernt (wird von Start.ps1 gesetzt)
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Config = "1.1"

# =====================================================================
# EINSTELLUNGEN
# =====================================================================
# HINWEIS: $global:ScriptVersion wird von Start.ps1 gesetzt
$script:SaveToSharePoint = $false  # Immer lokal im Quell-Verzeichnis

# =====================================================================
# LOOKUP-TABELLEN
# =====================================================================
$script:UserLookupTable = @{
    'M0733302' = 'WOB'
    'M0779325' = 'AZG'
    'M0555315' = 'CPA'
}

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

# =====================================================================
# ERWARTETE HEADER
# =====================================================================
$script:PMS_Header_Expected = @(
    "SLLLFN","SLLEAN","SLLPAS","SLLCAT","SLLGNR","SLLPCD","FLGSTP","FLGFXP","FLGVKF","RELDAT",
    "XML01","XML02","XML03","XML04","XML05","SLLVPL","SLLEPL","SLOEPF","SLOWAH","REDVPL",
    "SLLERR","SAAPNT","SLLIGN","IMPDAT","CHGDAT","SAASEL"
)

$script:PIM_Header_Expected = @(
    "Lieferant","EAN","Status","Kategorie","Genre","Preiscode","Standard VP","Fixer VP","Release Date",
    "Acquisition Price","Sales Price","Publisher ID","Linedisc","Bonusgroup","VP","PrioEP","RgEP",
    "Währung RgEP","Tiefpreis","Errorcode","Fehlercode","L-Prio-Punkte","Sperrcode","Verwendete Kalkulation",
    "letzter Import","letzte Änderung","letzter Status"
)
