VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------
' Class Module clsStats
' CompMan services statistics
' ----------------------------------
Private lNewSheets              As Long
Private lNewSheetNames          As Long
Private lNewSheetCodeNames      As Long
Private lNewNames               As Long
Private lNewNonDocMod           As Long
Private lNewShapes              As Long
Private lNewCols                As Long
Private lNewRefs                As Long
Private lObsoleteSheets         As Long
Private lObsoleteNames          As Long
Private lObsoleteNonDocMods     As Long
Private lObsoleteShapes         As Long
Private lObsoleteCols           As Long
Private lObsoleteRefs           As Long
Private lChangedVBComps         As Long

Public Function Counted(ByVal sic As siCounter) As Variant
    Select Case sic
        Case sic_changed_comps:         Counted = lChangedVBComps
        Case sic_new_cols:              Counted = lNewCols
        Case sic_new_names:             Counted = lNewNames
        Case sic_new_non_doc_mod:       Counted = lNewNonDocMod
        Case sic_new_refs:              Counted = lNewRefs
        Case sic_new_shapes:            Counted = lNewShapes
        Case sic_new_sheets:            Counted = lNewSheets
        Case sic_new_sheet_codenames:   Counted = lNewSheetCodeNames
        Case sic_new_sheet_names:       Counted = lNewSheetNames
        Case sic_obsolete_cols:         Counted = lObsoleteCols
        Case sic_obsolete_names:        Counted = lObsoleteNames
        Case sic_obsolete_non_doc_mod:  Counted = lObsoleteNonDocMods
        Case sic_obsolete_refs:         Counted = lObsoleteRefs
        Case sic_obsolete_shapes:       Counted = lObsoleteShapes
        Case sic_obsolete_sheets:       Counted = lObsoleteSheets
    End Select
    
End Function

Public Sub Count(ByVal sic As siCounter)
    Select Case sic
        Case sic_changed_comps:         lChangedVBComps = lChangedVBComps + 1
        Case sic_new_cols:              lNewCols = lNewCols + 1
        Case sic_new_names:             lNewNames = lNewNames + 1
        Case sic_new_non_doc_mod:       lNewNonDocMod = lNewNonDocMod + 1
        Case sic_new_refs:              lNewRefs = lNewRefs + 1
        Case sic_new_shapes:            lNewShapes = lNewShapes + 1
        Case sic_new_sheets:            lNewSheets = lNewSheets + 1
        Case sic_new_sheet_codenames:   lNewSheetCodeNames = lNewSheetCodeNames + 1
        Case sic_new_sheet_names:       lNewSheetNames = lNewSheetNames + 1
        Case sic_obsolete_cols:         lObsoleteCols = lObsoleteCols + 1
        Case sic_obsolete_names:        lObsoleteNames = lObsoleteNames + 1
        Case sic_obsolete_non_doc_mod:  lObsoleteNonDocMods = lObsoleteNonDocMods + 1
        Case sic_obsolete_refs:         lObsoleteRefs = lObsoleteRefs + 1
        Case sic_obsolete_shapes:       lObsoleteShapes = lObsoleteShapes + 1
        Case sic_obsolete_sheets:       lObsoleteSheets = lObsoleteSheets + 1
    End Select
    
End Sub


