Attribute VB_Name = "mSyncNames"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncNames." & s
End Function

Public Sub SyncNew()
' ----------------------------------------------------------------
' Synchronize the names in the target Worksheet (Sync.Target) with
' those new in the source Workbook (Sync.Source) considering that
' new Names which refer to a new sheet will automatically be
' synchronized when the new sheet is copied from the source to
' the target Workbook. All other new names refer to a range in
' an already existing sheet which might be new inserted columns
' or rows. Theses new names cannot be syncronized programmatically
' but require a manual intervention. This requirement will be
' communicated at the end of the syncronization.
' ----------------------------------------------------------------
    Const PROC = "SyncNew"
    
    On Error GoTo eh
    Dim nm              As Name
    Dim v               As Variant
    Dim SheetReferred   As String
    
    For Each v In Sync.SourceNames
        If Sync.TargetNames.Exists(v) Then GoTo next_v
        Stats.Count sic_names_new
        '~~ The source name not yet exists in the target Workbook and thus is regarde new
        '~~ However, new names potentially in concert require a design change of the concerned sheet
        Set nm = Sync.Source.Names.Item(v)
        SheetReferred = Replace(Split(nm.RefersTo, "!")(0), "=", vbNullString)
        If Sync.Mode = Confirm Then
            '~~ When the new name refers to a new sheet it is not syncronized
            If Not Sync.IsNewSheet(SheetReferred) Then
                '~~ New Names coming with new sheets are not displayed for confirmation
                Sync.ConfInfo(nm) = "New! Manual synchronization required! 3)"
                Sync.ManualSynchRequired = True
            End If
        Else
            If Not Sync.IsNewSheet(SheetReferred) Then Sync.ManualSynchRequired = True
        End If
next_v:
    Next v

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncObsolete()
' ----------------------------------------------------------------
' Synchronize the names in Worksheet (Sync.Target) with those in
' Workbook (Sync.Source) - when either a sheet's name (wb_sheet_name) or
' a sheet's CodeName (wb_sheet_codename) is provided only those
' names which refer to that sheet.
' - Note: Obsolete names are removed but missing names cannot be
'   added but are reported in the log file Missing names must be
'   added manually in concert with the corresponding sheet design
'   changes. As a consequence, design changes should preferrably
'   be made prior copy of the Workbook for a VB-Project
'   modification.
' - Precondition: The Worksheet's CodeName and Name are identical
'   both Workbooks. I.e. these need to be synced first.
' ---------------------------------------------------------------
    Const PROC = "SyncObsolete"
    
    On Error GoTo eh
    Dim nm  As Name
    Dim v   As Variant
    
    For Each v In Sync.TargetNames
        If Sync.SourceNames.Exists(v) Then GoTo next_v
        Stats.Count sic_names_obsolete
        Set nm = Sync.Target.Names.Item(v)
        '~~ The target name does not exist in the source and thus  has become obsolete
        If Sync.Mode = Confirm Then
            Log.ServicedItem = nm
            Sync.ConfInfo = "Obsolete! Will be removed."
        Else
            nm.Delete
            Log.Entry = "Obsolete (removed)"
        End If
next_v:
    Next v

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
