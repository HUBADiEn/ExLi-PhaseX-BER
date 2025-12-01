<#
.SYNOPSIS
    Konfiguration und Konstanten für PMS/PIM Vergleich

.NOTES
    File:           config_v1.3.ps1
    Version:        1.3
    Änderungshistorie:
        1.3 - PERFORMANCE: Index-Konstanten fuer String-Array-Zugriff
            - Ermoeglicht 12+ Mio Zeilen ohne OutOfMemoryException
        1.2 - PMS-Header: SLOEPF umbenannt zu SLOERG
        1.1 - FIX: ScriptVersion-Zeile entfernt (wird von Start.ps1 gesetzt)
        1.0 - Initiale Version (aus V1.103 extrahiert)
#>

# =====================================================================
# MODUL-VERSION (wird von Start.ps1 geprüft)
# =====================================================================
$script:ModuleVersion_Config = "1.3"

# =====================================================================
# EINSTELLUNGEN
# =====================================================================
$script:SaveToSharePoint = $false

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
# ERWARTETE HEADER (fuer Validierung)
# =====================================================================
$script:PMS_Header_Expected = @(
    "SLLLFN","SLLEAN","SLLPAS","SLLCAT","SLLGNR","SLLPCD","FLGSTP","FLGFXP","FLGVKF","RELDAT",
    "XML01","XML02","XML03","XML04","XML05","SLLVPL","SLLEPL","SLOERG","SLOWAH","REDVPL",
    "SLLERR","SAAPNT","SLLIGN","IMPDAT","CHGDAT","SAASEL"
)

$script:PIM_Header_Expected = @(
    "Lieferant","EAN","Status","Kategorie","Genre","Preiscode","Standard VP","Fixer VP","Release Date",
    "Acquisition Price","Sales Price","Publisher ID","Linedisc","Bonusgroup","VP","PrioEP","RgEP",
    "Währung RgEP","Tiefpreis","Errorcode","Fehlercode","L-Prio-Punkte","Sperrcode","Verwendete Kalkulation",
    "letzter Import","letzte Änderung","letzter Status"
)

# =====================================================================
# INDEX-KONSTANTEN FÜR PMS-FELDER (String-Array-Zugriff)
# Reihenfolge muss mit PMS_Header_Expected übereinstimmen!
# =====================================================================
$script:PMS_IDX = @{
    SLLLFN  = 0
    SLLEAN  = 1
    SLLPAS  = 2
    SLLCAT  = 3
    SLLGNR  = 4
    SLLPCD  = 5
    FLGSTP  = 6
    FLGFXP  = 7
    FLGVKF  = 8
    RELDAT  = 9
    XML01   = 10
    XML02   = 11
    XML03   = 12
    XML04   = 13
    XML05   = 14
    SLLVPL  = 15
    SLLEPL  = 16
    SLOERG  = 17
    SLOWAH  = 18
    REDVPL  = 19
    SLLERR  = 20
    SAAPNT  = 21
    SLLIGN  = 22
    IMPDAT  = 23
    CHGDAT  = 24
    SAASEL  = 25
}

# =====================================================================
# INDEX-KONSTANTEN FÜR PIM-FELDER (String-Array-Zugriff)
# Reihenfolge muss mit PIM_Header_Expected übereinstimmen!
# =====================================================================
$script:PIM_IDX = @{
    Lieferant           = 0
    EAN                 = 1
    Status              = 2
    Kategorie           = 3
    Genre               = 4
    Preiscode           = 5
    StandardVP          = 6   # "Standard VP"
    FixerVP             = 7   # "Fixer VP"
    ReleaseDate         = 8   # "Release Date"
    AcquisitionPrice    = 9   # "Acquisition Price"
    SalesPrice          = 10  # "Sales Price"
    PublisherID         = 11  # "Publisher ID"
    Linedisc            = 12
    Bonusgroup          = 13
    VP                  = 14
    PrioEP              = 15
    RgEP                = 16
    WaehrungRgEP        = 17  # "Währung RgEP"
    Tiefpreis           = 18
    Errorcode           = 19
    Fehlercode          = 20
    LPrioPunkte         = 21  # "L-Prio-Punkte"
    Sperrcode           = 22
    VerwendeteKalk      = 23  # "Verwendete Kalkulation"
    LetzterImport       = 24  # "letzter Import"
    LetzteAenderung     = 25  # "letzte Änderung"
    LetzterStatus       = 26  # "letzter Status"
}

# =====================================================================
# INDEX-KONSTANTEN FÜR RESULT-ARRAY (kombinierte Daten + Check-Ergebnisse)
# =====================================================================
$script:RES_IDX = @{
    # Meta-Felder
    EAN                 = 0
    Gefunden            = 1   # "nur im PMS", "nur im PIM", "im PMS und im PIM", "mehrfach im PIM"
    CheckSummary        = 2
    
    # Check-Ergebnisse (0-14)
    Check0              = 3   # Vorhanden in beiden Quellen
    Check1              = 4   # Status
    Check2              = 5   # Kategorie
    Check3              = 6   # Genre
    Check4              = 7   # Preiscode
    Check5              = 8   # Standard VP ab Lieferant
    Check6              = 9   # Fixer VP
    Check7              = 10  # Release-Datum
    Check8              = 11  # Errorcode
    Check9              = 12  # VP
    VPDiff              = 13
    Check10             = 14  # PrioEP
    PrioEPDiff          = 15
    Check11             = 16  # RgEP
    RgEPDiff            = 17
    Check12             = 18  # Tiefpreis
    Check13             = 19  # L-Prio Fehlercode
    Check14             = 20  # L-Prio
    LPrioDiff           = 21
    ZeitDiff            = 22  # ZeitDiff letzte Änderung
    ZeitDiffBewertung   = 23
    
    # PMS-Felder (ohne SLLLFN) - Start bei Index 24
    PMS_SLLEAN          = 24
    PMS_SLLPAS          = 25
    PMS_SLLCAT          = 26
    PMS_SLLGNR          = 27
    PMS_SLLPCD          = 28
    PMS_FLGSTP          = 29
    PMS_FLGFXP          = 30
    PMS_FLGVKF          = 31
    PMS_RELDAT          = 32
    PMS_XML01           = 33
    PMS_XML02           = 34
    PMS_XML03           = 35
    PMS_XML04           = 36
    PMS_XML05           = 37
    PMS_SLLVPL          = 38
    PMS_SLLEPL          = 39
    PMS_SLOERG          = 40
    PMS_SLOWAH          = 41
    PMS_REDVPL          = 42
    PMS_SLLERR          = 43
    PMS_SAAPNT          = 44
    PMS_SLLIGN          = 45
    PMS_IMPDAT          = 46
    PMS_CHGDAT          = 47
    PMS_SAASEL          = 48
    
    # PIM-Felder - Start bei Index 49
    PIM_Lieferant       = 49
    PIM_EAN             = 50
    PIM_Status          = 51
    PIM_Kategorie       = 52
    PIM_Genre           = 53
    PIM_Preiscode       = 54
    PIM_StandardVP      = 55
    PIM_FixerVP         = 56
    PIM_ReleaseDate     = 57
    PIM_AcquisitionPrice = 58
    PIM_SalesPrice      = 59
    PIM_PublisherID     = 60
    PIM_Linedisc        = 61
    PIM_Bonusgroup      = 62
    PIM_VP              = 63
    PIM_PrioEP          = 64
    PIM_RgEP            = 65
    PIM_WaehrungRgEP    = 66
    PIM_Tiefpreis       = 67
    PIM_Errorcode       = 68
    PIM_Fehlercode      = 69
    PIM_LPrioPunkte     = 70
    PIM_Sperrcode       = 71
    PIM_VerwendeteKalk  = 72
    PIM_LetzterImport   = 73
    PIM_LetzteAenderung = 74
    PIM_LetzterStatus   = 75
}

# Gesamtgröße des Result-Arrays
$script:RES_ARRAY_SIZE = 76

# =====================================================================
# OUTPUT CSV HEADER (für Export)
# =====================================================================
$script:OUTPUT_HEADER = @(
    "EAN"
    "Check Summary"
    "Check 0: Vorhanden in beiden Quellen"
    "Check 1: Status"
    "Check 2: Kategorie"
    "Check 3: Genre"
    "Check 4: Preiscode"
    "Check 5: Standard VP ab Lieferant"
    "Check 6: Fixer VP"
    "Check 7: Release-Datum"
    "Check 8: Errorcode"
    "Check 9: VP"
    "VP Diff"
    "Check 10: PrioEP"
    "PrioEP Diff"
    "Check 11: RgEP"
    "RgEP Diff"
    "Check 12: Tiefpreis"
    "Check 13: L-Prio Fehlercode"
    "Check 14: L-Prio"
    "L-Prio Diff"
    "ZeitDiff letzte Änderung"
    "ZeitDiff Bewertung"
    # PMS-Felder
    "PMS_SLLEAN"
    "PMS_SLLPAS"
    "PMS_SLLCAT"
    "PMS_SLLGNR"
    "PMS_SLLPCD"
    "PMS_FLGSTP"
    "PMS_FLGFXP"
    "PMS_FLGVKF"
    "PMS_RELDAT"
    "PMS_XML01"
    "PMS_XML02"
    "PMS_XML03"
    "PMS_XML04"
    "PMS_XML05"
    "PMS_SLLVPL"
    "PMS_SLLEPL"
    "PMS_SLOERG"
    "PMS_SLOWAH"
    "PMS_REDVPL"
    "PMS_SLLERR"
    "PMS_SAAPNT"
    "PMS_SLLIGN"
    "PMS_IMPDAT"
    "PMS_CHGDAT"
    "PMS_SAASEL"
    # PIM-Felder
    "PIM_Lieferant"
    "PIM_EAN"
    "PIM_Status"
    "PIM_Kategorie"
    "PIM_Genre"
    "PIM_Preiscode"
    "PIM_Standard VP"
    "PIM_Fixer VP"
    "PIM_Release Date"
    "PIM_Acquisition Price"
    "PIM_Sales Price"
    "PIM_Publisher ID"
    "PIM_Linedisc"
    "PIM_Bonusgroup"
    "PIM_VP"
    "PIM_PrioEP"
    "PIM_RgEP"
    "PIM_Währung RgEP"
    "PIM_Tiefpreis"
    "PIM_Errorcode"
    "PIM_Fehlercode"
    "PIM_L-Prio-Punkte"
    "PIM_Sperrcode"
    "PIM_Verwendete Kalkulation"
    "PIM_letzter Import"
    "PIM_letzte Änderung"
    "PIM_letzter Status"
)

# Mapping von RES_IDX zu OUTPUT_HEADER Index (für Export)
# Die Reihenfolge in OUTPUT_HEADER entspricht der gewünschten Spaltenreihenfolge
$script:RES_TO_OUTPUT = @(
    0   # EAN -> 0
    2   # CheckSummary -> 1
    3   # Check0 -> 2
    4   # Check1 -> 3
    5   # Check2 -> 4
    6   # Check3 -> 5
    7   # Check4 -> 6
    8   # Check5 -> 7
    9   # Check6 -> 8
    10  # Check7 -> 9
    11  # Check8 -> 10
    12  # Check9 -> 11
    13  # VPDiff -> 12
    14  # Check10 -> 13
    15  # PrioEPDiff -> 14
    16  # Check11 -> 15
    17  # RgEPDiff -> 16
    18  # Check12 -> 17
    19  # Check13 -> 18
    20  # Check14 -> 19
    21  # LPrioDiff -> 20
    22  # ZeitDiff -> 21
    23  # ZeitDiffBewertung -> 22
    # PMS-Felder (24-48)
    24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48
    # PIM-Felder (49-75)
    49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75
)

# =====================================================================
# HELPER-FUNKTION: Sicherer Array-Zugriff
# =====================================================================
function Get-SafeField {
    param(
        [string[]]$Fields,
        [int]$Index
    )
    if ($null -eq $Fields -or $Fields.Count -le $Index -or $Index -lt 0) {
        return ''
    }
    return $Fields[$Index]
}
