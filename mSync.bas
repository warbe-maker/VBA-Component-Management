Attribute VB_Name = "mSync"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSync
'          All services and means for the synchronization of a target Workbook/
'          VBProject with a source Workbook/VBProject.
'
' Public services:
'
' ----------------------------------------------------------------------------
Private Const SHEET_SHAPE       As String = ": "    ' Sheet-Shape concatenator

Private Sync                    As clsSync
Private lMode                   As SyncMode
Private dctChanged              As Dictionary       ' Confirm buttons and clsRaw items to display changed
Private dctNameChange           As Dictionary
Private bSyncDenied             As Boolean          ' when True the synchronization is not performed
Private bAmbigous               As Boolean          ' when True sync is only done when the below is confirmed True
Private RestrictRenameAsserted  As Boolean          ' when False a sheet's CodeName a n d its Name may be changed at once
Private cSource                 As clsRaw
Private cTarget                 As clsComp
Private lSheetsObsolete         As Long
Private lSheetsNew              As Long
Private ManualSynchRequired     As Boolean

Public Enum siCounter
    sic_cols_new
    sic_cols_obsolete
    sic_names_new
    sic_names_obsolete
    sic_names_total
    sic_refs_new
    sic_refs_obsolete
    sic_refs_total
    sic_rows_obsolete
    sic_rows_new
    sic_shapes_new
    sic_shapes_obsolete
    sic_shapes_total
    sic_sheets_code
    sic_sheets_codename
    sic_sheets_name
    sic_sheets_new
    sic_sheets_obsolete
    sic_sheets_total
    sic_vbcomps_total
    sic_non_doc_mods_code
    sic_non_doc_mod_new
    sic_non_doc_mod_obsolete
    sic_non_doc_mod_total
End Enum

Private Enum SyncMode
    Count = 1
    Confirm = 2
    Synchronize = 3
End Enum

Public Sub SyncTargetRestore( _
                       ByRef bkp_folder As String, _
                       ByVal sTarget As String)
' --------------------------------------------------
'
' --------------------------------------------------

    Dim fso As New FileSystemObject

    With fso
        If Not .FolderExists(bkp_folder) Then GoTo xt
        If Not .FileExists(sTarget) Then GoTo xt
        .CopyFile bkp_folder & "\" & .GetFileName(sTarget), sTarget
        .DeleteFolder bkp_folder
    End With
    
xt: Set fso = Nothing
   Exit Sub
   
End Sub

Public Sub SyncTargetBackup(ByRef bkp_folder As String, _
                            ByVal sTarget As String)
' -----------------------------------------------------
' Saves a copy of the synchronization target Workbook
' (Sync.Target) in a time-stamped folder under the
' Workbook folder returned in (bkp_folder).
' -----------------------------------------------------
    Const PROC = "SyncTargetBackup"
    
    On Error GoTo eh
    Dim BckpFolderName  As String
    Dim fo              As Folder
    
    BckpFolderName = "Bckup-" & Format$(Now(), "YYMMDD-hhmmss")
    With New FileSystemObject
        While .FolderExists(.GetParentFolderName(sTarget) & "\" & BckpFolderName)
            Application.Wait Now() + 0.000001
            BckpFolderName = "Bckup-" & Format$(Now(), "YYMMDD-hhmmss")
        Wend
        Set fo = .CreateFolder(.GetParentFolderName(sTarget) & "\" & BckpFolderName)
        .CopyFile sTarget, fo.Path & "\" & .GetFileName(sTarget)
    End With

xt: bkp_folder = fo.Path
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub ByCodeLines( _
                 ByVal sync_target_comp_name As String, _
                 ByVal wb_source_full_name As String, _
        Optional ByRef sync_source_codelines As Dictionary = Nothing)
' -------------------------------------------------------------------------
' Synchronizes
'  the component (sync_target_comp_name) in the target Workbook
'  (Sync.Target) with the code (sync_source_codelines) in the Export-File
'  of the corresponding source Workbook's (wb_source_full_name)
'  component
' line by line.
' When the source code lines () are not provided they are obtained from the
' source Workbook's () corresponding Export-File.
' ----------------------------------------------------------------
    Const PROC = "ByCodeLines"

    On Error GoTo eh
    Dim i       As Long: i = 1
    Dim v       As Variant
    Dim ws      As Worksheet
    Dim wbRaw   As Workbook
    
    If sync_source_codelines Is Nothing Then
        '~~ Obtain non provided code lines for the line by line syncronization
        Set wbRaw = WbkGetOpen(wb_source_full_name)
        Set cSource.Wrkbk = wbRaw
        cSource.CompName = sync_target_comp_name
        Set sync_source_codelines = cSource.CodeLines
    End If
    
    With Sync.Target.VBProject.VBComponents(sync_target_comp_name).CodeModule
        If .CountOfLines > 0 _
        Then .DeleteLines 1, .CountOfLines   ' Remove all lines from the cloned raw component
        
        For Each v In sync_source_codelines    ' Insert the raw component's code lines
            .InsertLines i, sync_source_codelines(v)
            i = i + 1
        Next v
    End With
                
xt: Set cSource = Nothing
    Set wbRaw = Nothing
    Set ws = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub DisconnectLinkedRanges()
' --------------------------------------------
' Provided all sheets had been synchronized
' any range still linked to a source Workbooks
' range must be disconnected.
' --------------------------------------------
    Dim nm As Name
    Dim sName As String
    
    For Each nm In Sync.Target.Names
        On Error Resume Next
        sName = Split(nm.RefersTo, "]")(1)
        If Err.Number = 0 Then
            nm.RefersTo = "=" & sName
            cLog.ServicedItem = nm
            cLog.Entry = "Link to source sheet removed"
        End If
    Next nm
    
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Private Function NameChange( _
                      ByVal sh_name As String, _
                      ByVal sh_code_name As String) As Boolean
' ------------------------------------------------------------
' Returns TRUE when either name is involved in a name change.
' ------------------------------------------------------------
    NameChange = dctNameChange.Exists(sh_name)
    If Not NameChange Then NameChange = dctNameChange.Exists(sh_code_name)
End Function

Private Function NameExists( _
                      ByRef ne_wb As Workbook, _
                      ByVal ne_nm As Name) As Boolean
    Dim nm As Name
    For Each nm In ne_wb.Names
        NameExists = nm.Name = ne_nm.Name
        If NameExists Then Exit For
    Next nm
End Function

Private Function RefExists( _
                     ByRef re_wb As Workbook, _
                     ByVal re_ref As Reference) As Boolean
' --------------------------------------------------------
'
' --------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In re_wb.VBProject.References
        RefExists = ref.Name = re_ref.Name
        If RefExists Then Exit Function
    Next ref

End Function

Private Sub RefRemove(ByVal rr_ref As Reference)
' -------------------------------------------------
' Removes Reference (rr_ref) from Workbook (rr_wb).
' -------------------------------------------------
    Dim ref As Reference
    
    With Sync.Target.VBProject
        For Each ref In .References
            If ref.Name = rr_ref.Name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
End Sub

Private Sub RemoveInvalidRangeNames()
' -----------------------------------------------------------
' Removes names which point to a range which not or no longer
' exists.
' -----------------------------------------------------------
    Dim nm As Name
    For Each nm In Sync.Target.Names
        Debug.Print nm.value
        If InStr(nm.value, "#") <> 0 Or InStr(nm.RefersTo, "#") <> 0 Then
            cLog.ServicedItem = nm
            nm.Delete
            cLog.Entry = "Deleted! (invalid)"
        End If
    Next nm
End Sub

Private Sub RenameSheet(ByRef rs_wb As Workbook, _
                        ByVal rs_old_name As String, _
                        ByVal rs_new_name As String)
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "RenameSheet"
    
    On Error GoTo eh
    Dim sh  As Worksheet
    For Each sh In rs_wb.Worksheets
        If sh.Name = rs_old_name Then
            sh.Name = rs_new_name
            cLog.Entry = "Sheet-Name changed to '" & rs_new_name & "'"
            Exit For
        End If
    Next sh

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub RenameWrkbkModule( _
                        ByRef rdm_wb As Workbook, _
                        ByVal rdm_new_name As String)
' ---------------------------------------------------
' Renames in Workbook (rdm_wb) the Workbook Module
' to (rdm_new_name).
' ---------------------------------------------------
    Const PROC = "RenameWrkbkModule"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    
    With rdm_wb.SyncTargetithSource
        For Each vbc In .VBComponents
            If vbc.Type = vbext_ct_Document Then
                If IsWrkbkComp(vbc) Then
                    cLog.ServicedItem = vbc
                    vbc.Name = rdm_new_name
                    cLog.Entry = "Renamed to '" & rdm_new_name & "'"
                    DoEvents
                    Exit For
                End If
            End If
        Next vbc
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub ShapeCopy( _
                ByRef sc_source As Worksheet, _
                ByRef sc_target As Worksheet, _
                ByVal sc_name As String)
' ---------------------------------------------
'
' ---------------------------------------------
    Dim SourceShape As Shape
    Dim TargetShape As Shape
    
    For Each SourceShape In sc_source.Shapes
        If SourceShape.Name <> sc_name Then GoTo next_shape
        SourceShape.Copy
        sc_target.Paste
        Set TargetShape = sc_target.Shapes(sc_target.Shapes.Count)
        With TargetShape
            .Name = sc_name
            .top = SourceShape.top
            .Left = SourceShape.Left
            .Width = SourceShape.Width
            .Height = SourceShape.Height
        End With
next_shape:
    Next SourceShape
End Sub

Private Function SheetCodeName( _
                         ByRef sync_wb As Workbook, _
                         ByVal sync_sheet_name As String) As String
' -----------------------------------------------------------------
'
' -----------------------------------------------------------------
    Const PROC = "SheetCodeName"

    On Error GoTo eh
    Dim ws  As Worksheet
    
    For Each ws In sync_wb.Worksheets
        If ws.Name = sync_sheet_name Then
            SheetCodeName = ws.CodeName
            GoTo xt
        End If
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function

Private Function SheetExists( _
                       ByRef wb As Workbook, _
              Optional ByRef sh1_name As String = vbNullString, _
              Optional ByRef sh1_code_name As String = vbNullString, _
              Optional ByRef sh2_name As String = vbNullString, _
              Optional ByRef sh2_code_name As String = vbNullString) As Boolean
' -----------------------------------------------------------------------------
' Returns TRUE when the sheet either under the Name (sh1_name) or under the
' CodeName (sh1_code_name) exists in Workbook (wb).
' Returns FALSE when the sheet not exists in the Workbook (wb) under either
' name. When it exists by Name or CodeName both are returned (sh2_name,
' sh2_code_name).
' -----------------------------------------------------------------------------
    Const PROC = "SheetExists"
                             
    On Error GoTo eh
    Dim ws As Worksheet
    
    If sh1_name = vbNullString And sh1_code_name = vbNullString _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Neither a Sheet's Name nor CodeName is provided!"
    
    For Each ws In wb.Worksheets
        If sh1_name <> vbNullString Then
            If ws.Name = sh1_name Then
                sh2_name = ws.Name
                sh2_code_name = ws.CodeName
                SheetExists = True
                Exit For
            End If
        End If
        If sh1_code_name <> vbNullString Then
            If ws.CodeName = sh1_code_name Then
                sh2_name = ws.Name
                sh2_code_name = ws.CodeName
                SheetExists = True
                Exit For
            End If
        End If
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function SheetShapeExists( _
                            ByRef sync_wb As Workbook, _
                            ByVal sync_sheet_name As String, _
                            ByVal sync_sheet_code_name As String, _
                            ByVal sync_shape_name As String) As Boolean
' ----------------------------------------------------------------------
' Returns TRUE when the shape (sync_shape_name) exists in the Workbook
' (sync_wb) in a sheet with either the given Name (sync_sheet_name) or
' the provided CodeName (sync_sheet_code_name).
' Explanation: When this function is used to get the required info for
'              being confirmed, the concerned sheet may be one of which
'              the Name or the CodeName is about to be renamed - which
'              by then will not have taken place.
' ----------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim ws  As Worksheet
    Dim shp As Shape
    
    For Each ws In sync_wb.Worksheets
        If ws.Name <> sync_sheet_name And ws.CodeName <> sync_sheet_code_name Then GoTo next_sheet
        For Each shp In ws.Shapes
            If shp.Name = sync_shape_name Then
                SheetShapeExists = True
                GoTo xt
            End If
        Next shp
next_sheet:
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function

Private Sub SheetsOrder()
' -------------------------------------------------------
'
' -------------------------------------------------------
    Const PROC = "SheetsOrder"
    
    On Error GoTo eh
    Dim i           As Long
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    
    For i = 1 To Sync.Source.Worksheets.Count
        Set wsSource = Sync.Source.Worksheets(i)
        Set wsTarget = Sync.Target.Worksheets(i)
        If wsSource.Name <> wsTarget.Name Then
            '~~ Sheet position has changed
            If lMode = Confirm Then
                Stop ' pending confirmation info
            Else
                Stop ' pending implementation
            End If
        End If
    Next i
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

 
Private Sub SourceSheetNameChange( _
                            ByVal sh1_name As String, _
                            ByVal sh1_code_name As String, _
                            ByVal sh2_name As String, _
                            ByVal sh2_code_name As String)
' ----------------------------------------------------------
' Registers all sheet names involved in name changes.
' ----------------------------------------------------------
    If Not dctNameChange.Exists(sh1_name) Then dctNameChange.Add sh1_name, sh1_name
    If Not dctNameChange.Exists(sh1_code_name) Then dctNameChange.Add sh1_code_name, sh1_code_name
    If Not dctNameChange.Exists(sh2_name) Then dctNameChange.Add sh2_name, sh2_name
    If Not dctNameChange.Exists(sh2_code_name) Then dctNameChange.Add sh2_code_name, sh2_code_name
End Sub

Private Sub SyncConfirmation()
' ------------------------------------------------------------
' Collect all confirmation information regarding sheet changes
' ------------------------------------------------------------
    lMode = Confirm
    SyncReferencesNew
    SyncReferencesObsolete
    
    SyncSheetsName
    SyncSheetsCodeName
    SyncSheetsNew
    SyncSheetsObsolete
    SyncSheetsCode
    
    SyncShapesNew
    SyncShapesObsolete
    
    SyncVBCompsNew
    SyncVBCompsObsolete
    SyncVBCompsCodeChanged
    
    SyncNamesNew
End Sub

Private Sub SyncNamesNew()
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
    Const PROC = "SyncNamesNew"
    
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
        If lMode = Confirm Then
            '~~ When the new name refers to a new sheet it is not syncronized
            If Not Sync.NewSheetExists(SheetReferred) Then
                '~~ New Names coming with new sheets are not displayed for confirmation
                Sync.ConfInfo(nm) = "New! Manual synchronization required! 3)"
                ManualSynchRequired = True
            End If
        Else
            If Not Sync.NewSheetExists(SheetReferred) Then ManualSynchRequired = True
        End If
next_v:
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub SyncNamesObsolete()
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
'   both Workbooks. I.e. these need to be synched first.
' ---------------------------------------------------------------
    Const PROC = "SyncNamesObsolete"
    
    On Error GoTo eh
    Dim nm  As Name
    Dim v   As Variant
    
    For Each v In Sync.TargetNames
        If Sync.SourceNames.Exists(v) Then GoTo next_v
        Stats.Count sic_names_obsolete
        Set nm = Sync.Target.Names.Item(v)
        '~~ The target name does not exist in the source and thus  has become obsolete
        If lMode = Confirm Then
            cLog.ServicedItem = nm
            Sync.ConfInfo = "Obsolete! Will be removed."
        Else
            nm.Delete
            cLog.Entry = "Obsolete (removed)"
        End If
next_v:
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncReferencesNew()
' --------------------------------------------
' When lMode=Confirm to be synchronized
' References are collected for being confirmed
' else References are synchronized.
' --------------------------------------------
    Const PROC = "SyncReferencesNew"
    
    On Error GoTo eh
    Dim ref As Reference
    
    For Each ref In Sync.Source.VBProject.References
        If Not RefExists(Sync.Target, ref) Then
            cLog.ServicedItem = ref
            Stats.Count sic_refs_new
            If lMode = Confirm Then
                Sync.ConfInfo = "New!"
            Else
                Sync.Target.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
                cLog.Entry = "Added"
            End If
        End If
    Next ref
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

Private Sub SyncReferencesObsolete()
' --------------------------------------------
' When lMode=Confirm to be synchronized
' References are collected for being confirmed
' else References are synchronized.
' --------------------------------------------
    Const PROC = "SyncReferencesObsolete"
    
    On Error GoTo eh
    Dim ref     As Reference
    Dim sRef    As String
    
    For Each ref In Sync.Target.VBProject.References
        If Not RefExists(Sync.Source, ref) Then
            cLog.ServicedItem = ref
            Stats.Count sic_refs_new
            sRef = ref.Name
            If lMode = Confirm Then
                Sync.ConfInfo = "Obsolete!"
            Else
                RefRemove ref
                cLog.Entry = "Removed!"
            End If
        End If
    Next ref

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

Private Sub SyncShapesNew()
' -----------------------------------------------------------
' Copy new shapes from the sourec Workbook (Sync.Source) to
' the target Workbook (Sync.Target) and ajust the properties
' -----------------------------------------------------------
    Const PROC = "SyncShapesNew"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    Dim sShape      As String
    Dim sSheet      As String
    Dim shp         As Shape
    
    With Sync
        For Each v In .SourceSheetShapes
            sSheet = mSync.KeySheetName(v)
            sShape = mSync.KeyShapeName(v)
            If SheetShapeExists(sync_wb:=.Target _
                              , sync_sheet_name:=sSheet _
                              , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                              , sync_shape_name:=sShape _
                               ) _
            Then GoTo next_shape
            '~~ The source shape not yet exists in the target Workbook's corresponding sheet
            '~~ (idetified either by its Name or CodeName) and thus is regarde new and needs
            '~~ to be copied and its Properties adjusted.
            Set wsSource = .Source.Worksheets(.SourceSheetShapes(v))
            Set shp = wsSource.Shapes(sShape)
            cLog.ServicedItem = shp
            If lMode = Confirm Then
                '~~ New shapes coming with new sheets are not displayed for confirmation
                If Not .NewSheetExists(sSheet) Then
                     Stats.Count sic_shapes_new
                    .ConfInfo = "New!"
                End If
            Else
                '~~ When new shapes are syncronized the Worksheet's Name/CodeName will have been syncronized before!
                If Not .NewSheetExists(sSheet) Then
                    Set wsTarget = .Target.Worksheets(.SourceSheetShapes(v))
                    ShapeCopy sc_source:=wsSource _
                            , sc_target:=wsTarget _
                            , sc_name:=sShape
                    cLog.Entry = "Copied from source sheet"
                End If
            End If
next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncShapesObsolete()
' -----------------------------------------------------------
' Remove obsolete shapes in the target Workbook (Sync.Target)
' -----------------------------------------------------------
    Const PROC = "SyncShapesObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsTarget    As Worksheet
    Dim sShape      As String
    Dim sSheet      As String
    Dim shp         As Shape
    
    For Each v In Sync.TargetSheetShapes
        sSheet = mSync.KeySheetName(v)
        sShape = mSync.KeyShapeName(v)
        If SheetShapeExists(sync_wb:=Sync.Source _
                          , sync_sheet_name:=KeySheetName(v) _
                          , sync_sheet_code_name:=SheetCodeName(Sync.Target, sSheet) _
                          , sync_shape_name:=sShape _
                           ) _
        Then GoTo next_shape
        Set wsTarget = Sync.Target.Worksheets(Sync.TargetSheetShapes.Item(v))
        Set shp = wsTarget.Shapes(sShape)
        cLog.ServicedItem = shp
        
        Stats.Count sic_shapes_obsolete
        '~~ The target name does not exist in the source and thus  has become obsolete
        If lMode = Confirm Then
            Sync.ConfInfo = "Obsolete!"
        Else
            wsTarget.Shapes(sShape).Delete
            cLog.Entry = "Deteted!"
        End If
next_shape:
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Function KeySheetName(ByVal s As String) As String
    KeySheetName = Split(s, SHEET_SHAPE)(0)
End Function

Public Function KeyShapeName(ByVal s As String) As String
    KeyShapeName = Split(s, SHEET_SHAPE)(1)
End Function

Public Function KeySheetShape(ByVal sheet_name As String, _
                              ByVal shape_name As String) As String
    KeySheetShape = sheet_name & SHEET_SHAPE & shape_name
End Function

Private Sub SyncShapesProperties()
' -----------------------------------------
' Syncronize all shape's properties between
' source and target Workbook.
' -----------------------------------------
    Const PROC = "SyncShapesProperties"
        
    On Error GoTo eh
    Dim v           As Variant
    Dim sShape      As String
    Dim sSheet      As String
    
    With Sync
        For Each v In .SourceSheetShapes
            sSheet = mSync.KeySheetName(v)
            sShape = mSync.KeyShapeName(v)
            If Not SheetShapeExists(sync_wb:=.Target _
                                  , sync_sheet_name:=sSheet _
                                  , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                                  , sync_shape_name:=sShape _
                                   ) _
            Then GoTo next_shape
            mShape.SyncProperties .Source.Worksheets(sSheet).Shapes.Item(sShape) _
                                , .Target.Worksheets(sSheet).Shapes.Item(sShape)

next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncSheetsCode()
' -----------------------------------------------
' When lMode=Confirm all sheets which had changed
' are collected and provided for confirmation
' else the changes are syncronized.
' -----------------------------------------------
    Const PROC = "SyncSheetsCode"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim vbc                 As VBComponent
    Dim sCaption            As String
    Dim sExpFile            As String
    
    For Each vbc In Sync.Source.VBProject.VBComponents
        If Not vbc.Type = vbext_ct_Document Then GoTo next_sheet
        If Not IsSheetComp(vbc) Then GoTo next_sheet
        
        Set cSource = New clsRaw
        Set cSource.Wrkbk = Sync.Source
        cSource.CompName = vbc.Name
        If Not cSource.Exists(Sync.Target) Then GoTo next_sheet
        
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = Sync.Target
        cTarget.CompName = vbc.Name
        cSource.CloneExpFileFullName = cTarget.ExpFileFullName
        If Not cSource.Changed Then GoTo next_sheet
        
        cLog.ServicedItem = vbc
        Stats.Count sic_non_doc_mods_code
        
        If lMode = Confirm Then
            Sync.ConfInfo = "Code changed!"
            sCaption = "Display changes" & vbLf & "of" & vbLf & vbLf & vbc.Name & vbLf
            If Not dctChanged.Exists(sCaption) _
            Then dctChanged.Add sCaption, cSource
        Else
            sExpFile = cSource.ExpFileFullName
            mSync.ByCodeLines sync_target_comp_name:=vbc.Name _
                            , wb_source_full_name:=cSource.Wrkbk.FullName _
                            , sync_source_codelines:=cSource.CodeLines
            cLog.Entry = "Code updated line-by-line with code from Export-File '" & sExpFile & "'"
        End If
        Set cSource = Nothing
        Set cTarget = Nothing
next_sheet:
    Next vbc

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SyncSheetsCodeName()
' ------------------------------------
'
' ------------------------------------
    Const PROC = "SyncSheetsCodeName"
    
    On Error GoTo eh
    Dim v                       As Variant
    Dim wsTarget                As Worksheet
    Dim vbc                     As VBComponent
    Dim sSourceSheetCodeName    As String
    Dim sSourceSheetName        As String
    Dim sTargetSheetCodeName    As String
    Dim sTargetSheetName        As String
    
    For Each v In Sync.SourceSheets
        sSourceSheetName = Sync.SourceSheets(v)
        sSourceSheetCodeName = SheetCodeName(Sync.Source, sSourceSheetName)
        If Not SheetExists(wb:=Sync.Target _
                         , sh1_name:=sSourceSheetName _
                         , sh1_code_name:=sSourceSheetCodeName _
                         , sh2_name:=sTargetSheetName _
                         , sh2_code_name:=sTargetSheetCodeName _
                          ) _
        Then GoTo next_sheet
        If sTargetSheetCodeName <> sSourceSheetCodeName And sTargetSheetName = sSourceSheetName Then
            '~~ The sheet's CodName has changed while the sheet's Name remained unchanged
            cLog.ServicedItem = wsTarget
            Stats.Count sic_sheets_codename
            
            If lMode = Confirm Then
                Sync.ConfInfo = "CodeName change to '" & sSourceSheetCodeName & "'"
            Else
                For Each vbc In Sync.Target.VBProject.VBComponents
                    If vbc.Name = sTargetSheetCodeName Then
                        vbc.Name = sSourceSheetCodeName
                        '~~ When the sheet's CodeName has changed the sheet's code is synchronized line by line
                        '~~ because it is very likely code refers to the CodeName rather than to the sheet's Name or position
'                        mSync.ByCodeLines sync_target_comp_name:=wsSource.CodeName _
                                        , wb_source_full_name:=SyncSource.FullName
                        cLog.Entry = "CodeName changed to '" & sSourceSheetCodeName & "'"
                        Exit For
                    End If
                Next vbc
            End If
        End If
next_sheet:
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SyncSheetsName()
' --------------------------------
'
' --------------------------------
    Const PROC = "SyncSheetsName"
    
    On Error GoTo eh
    Dim v                       As Variant
    Dim sSourceSheetCodeName    As String
    Dim sSourceSheetName        As String
    Dim sTargetSheetCodeName    As String
    Dim sTargetSheetName        As String
    
    For Each v In Sync.SourceSheets
        sSourceSheetName = Sync.SourceSheets(v)
        sSourceSheetCodeName = SheetCodeName(Sync.Source, sSourceSheetName)
        If Not SheetExists(wb:=Sync.Target _
                         , sh1_name:=sSourceSheetName _
                         , sh1_code_name:=sSourceSheetCodeName _
                         , sh2_name:=sTargetSheetName _
                         , sh2_code_name:=sTargetSheetCodeName _
                          ) _
        Then GoTo next_comp
        If sTargetSheetCodeName = sSourceSheetCodeName And sTargetSheetName <> sSourceSheetName Then
            cLog.ServicedItem = Sync.Source.Worksheets(sSourceSheetName)
            Stats.Count sic_sheets_name
            
            '~~ The sheet's Name has changed while the sheets CodeName remained unchanged
            If lMode = Confirm Then
                Sync.ConfInfo = "Name change to '" & sSourceSheetName & "'."
                SourceSheetNameChange sSourceSheetName, sSourceSheetCodeName, sTargetSheetName, sTargetSheetCodeName
            Else
                Sync.Target.Worksheets(sTargetSheetName).Name = sSourceSheetName
                cLog.Entry = "Name changed to '" & sSourceSheetName & "'."
            End If
        End If
next_comp:
    Next v

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SyncSheetsNew()
' ---------------------------------------------------------------
' Synchronize new sheets in the source Workbook (Sync.Source) with
' the target Workbook (Sync.Target).
' - When the optional new sheets counter (sync_new_count) is
'   provided, the new sheets are only counted
' - In lMode=Confirm only the syncronization infos are collect
'   for being confirmed.
' Note-1: A Worksheet is regarded new when it exists in the
'         target Workbbook under its CodeName nor its Name and
'         it is asserted that sheet's name change is restricted
'         to either Name or CodeName but never both at once.
' Note-2: This procedure is called three times
'         1. To count the sheets indicated new
'         2. To get the new sheets confirmed
'         3. To copy the new sheet from the source to the target
'            Workbook
' --------------------------------------------------------------
    Const PROC = "SyncSheetsNew"
    
    On Error GoTo eh
    Dim ws                      As Worksheet
    Dim sSourceSheetName        As String
    Dim sTargetSheetName        As String
    Dim sSourceSheetCodeName    As String
    Dim sTargetSheetCodeName    As String
    Dim v                       As Variant
    
    For Each v In Sync.SourceSheets
        sSourceSheetName = Sync.SourceSheets(v)
        sSourceSheetCodeName = SheetCodeName(Sync.Source, sSourceSheetName)
        If Not SheetExists(wb:=Sync.Target _
                         , sh1_name:=sSourceSheetName _
                         , sh1_code_name:=sSourceSheetCodeName _
                         , sh2_name:=sTargetSheetName _
                         , sh2_code_name:=sTargetSheetCodeName _
                         ) Then
            If NameChange(sSourceSheetName, sSourceSheetCodeName) Then GoTo next_v
    
            '~~ The sheet not exist in the target Workbook under the Name nor under the CodeName.
            Set ws = Sync.Source.Worksheets(sSourceSheetName)
            cLog.ServicedItem = ws
            Stats.Count sic_sheets_new
            
            If lMode = Count Then
                '~~ This is just the first call for counting the potentially new sheets
                lSheetsNew = lSheetsNew + 1
            ElseIf lMode = Confirm Then
                '~~ This is just the second call for the collection of the sync confirmation info
                If lSheetsNew > 0 Or lSheetsObsolete > 0 Then
                    If Not RestrictRenameAsserted Then
                        bAmbigous = True
                        Sync.ConfInfo = "New! Ambigous! 1)"
                        Sync.NewSheet(sSourceSheetCodeName) = sSourceSheetName
                    Else
                        bAmbigous = False
                        Sync.ConfInfo = "New! 2)"
                        Sync.NewSheet(sSourceSheetCodeName) = sSourceSheetName
                    End If
                Else
                    Sync.ConfInfo = "New!"
                    Sync.NewSheet(sSourceSheetCodeName) = sSourceSheetName
                End If
            Else
                '~~ This is the third call for getting the syncronizations done
                '~~ The new sheet is copied to the corresponding position in the target Workbook
                Sync.Source.Worksheets(sSourceSheetName).Copy _
                After:=Sync.Target.Sheets(Sync.Target.Worksheets.Count)
                cLog.Entry = "Copied from source Workbook."
            End If
        End If
next_v:
    Next v
       
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SyncSheetsObsolete()
' --------------------------------------------------------------------
' Remove sheets in the target (Sync.Target) which are regarded
' obsolete because they do not exist in the target Workbook
' (Sync.Target) neither under their Name nor theit CodeName.
' - When the optional obsolet sheets counter (sync_obsolete_count)
'   is provided, the obsolete sheets are only counted
' - In lMode=Confirm only the syncronization infos are collected
'   for being confirmed.
' A Worksheet is finally only regarded a obsolete when:
' A) it exists in the source Workbook neither under its CodeName nor
'    its Name a n d  it had been confirmed that name changes on sheets
'    are restricted to either or but never both at once.
' B) the number of new sheets is 0
' Note: This procedure is called three times
' 1. To count the sheets indicated obsolete
' 2. To get the removal of the obsolete sheets confirmed
' 3. To remove the obsolete sheets
' -----------------------------------------------------------------
    Const PROC = "SyncSheetsObsolete"
    
    On Error GoTo eh
    Dim ws                      As Worksheet
    Dim cSource                 As clsRaw
    Dim cTarget                 As clsComp
    Dim v                       As Variant
    Dim sTargetSheetName        As String
    Dim sTargetSheetCodeName    As String
    
    For Each v In Sync.TargetSheets
        sTargetSheetName = Sync.TargetSheets(v)
        sTargetSheetCodeName = SheetCodeName(Sync.Target, sTargetSheetName)
        If Not SheetExists(wb:=Sync.Source _
                         , sh1_name:=sTargetSheetName _
                         , sh1_code_name:=sTargetSheetCodeName _
                         ) Then
            If NameChange(sTargetSheetName, sTargetSheetCodeName) Then GoTo next_v
            
            '~~ Target sheet not or no longer exists in source Workbook
            '~~ neither under the Name nor under the CodeName
            Set ws = Sync.Target.Worksheets(sTargetSheetName)
            cLog.ServicedItem = ws
            Stats.Count sic_sheets_obsolete
            
            If lMode = Count Then
                '~~ This is just the first call for counting the potentially new sheets
                lSheetsObsolete = lSheetsObsolete + 1
            ElseIf lMode = Confirm Then
                '~~ This is just the second call for the collection of the sync confirmation info
                If lSheetsNew > 0 Or lSheetsObsolete > 0 Then
                    If Not RestrictRenameAsserted Then
                        bAmbigous = True
                        Sync.ConfInfo = "Obsolete! Ambigous 1)"
                    Else
                        bAmbigous = False
                        Sync.ConfInfo = "Obsolete! 2)"
                    End If
                Else
                    Sync.ConfInfo = "Obsolete!"
                End If
            Else
                If Not RestrictRenameAsserted Then GoTo xt
                '~~ This is a Worksheet with no corresponding component and no corresponding sheet in the source Workbook.
                '~~ Because it has been asserted that sheets are never renamed by Name and CodeName at once
                '~~ this Worksheet is regarded obsolete for sure and will thus now be removed
                For Each ws In Sync.Target.Worksheets
                    If ws.CodeName = sTargetSheetCodeName Then
                        '~~ This is the obsolete sheet to be removed
                        Application.DisplayAlerts = False
                        ws.Delete
                        Application.DisplayAlerts = True
                        cLog.Entry = "Obsolete (deleted)"
                        Exit For
                    End If
                Next ws
            End If
            Set cTarget = Nothing
            Set cSource = Nothing
        End If
next_v:
    Next v
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

             
Public Function SyncTargetWithSource( _
                               ByRef wb_target As Workbook, _
                               ByRef wb_source As Workbook, _
                      Optional ByVal restricted_sheet_rename_asserted As Boolean = False, _
                      Optional ByRef bkp_folder As String) As Boolean
' -----------------------------------------------------------------------------------------
' Synchronizes a target Workbook (Sync.Target) with a source Workbook (Sync.Source).
' Returns TRUE when finished without error.
' -----------------------------------------------------------------------------------------
    Const PROC = "SyncTargetWithSource"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim v               As Variant
    Dim sMsg            As tMsg
    Dim sBttnCnfrmd     As String
    Dim sBttnTrmnt      As String
    Dim sBttnRestricted As String
    Dim cllButtons      As Collection
    Dim sReply          As String
    
    Set Sync = New clsSync
    Set Stats = New clsStats
    
    If dctNameChange Is Nothing Then Set dctNameChange = New Dictionary Else dctNameChange.RemoveAll
    If dctChanged Is Nothing Then Set dctChanged = New Dictionary Else dctChanged.RemoveAll
    
    RestrictRenameAsserted = restricted_sheet_rename_asserted
    Set Sync.Source = wb_source
    Set Sync.Target = wb_target
    Sync.CollectAllSyncItems
    
    ManualSynchRequired = False
    
    '~~ Count new and obsolete sheets
    lSheetsNew = 0
    lSheetsObsolete = 0
    
    lMode = Count
    SyncSheetsNew
    SyncSheetsObsolete
    
    RestrictRenameAsserted = False
    bAmbigous = True
    bSyncDenied = True

    Do
        
        '~~ Collect all synchronization info and get them confirmed
        SyncConfirmation
        
        '~~ Get the collected info confirmed
        sMsg.Section(1).sText = Sync.ConfInfo
        sMsg.Section(1).bMonspaced = True
        sMsg.Section(2).sText = "The above syncronisations need to be confirmed - or " & _
                                "terminated in case of any concerns!"
        
        sBttnCnfrmd = "Synchronize" & vbLf & vbLf & fso.GetBaseName(Sync.Target.Name) & vbLf & " with " & vbLf & fso.GetBaseName(Sync.Source.Name)
        sBttnTrmnt = "Terminate!" & vbLf & vbLf & "Synchronization denied" & vbLf & "because of concerns"
        sBttnRestricted = "Confirmed" & vbLf & "that Sheet rename" & vbLf & "is restricted" & vbLf & "(either Name  o r  CodeName)"
        
        If bAmbigous And Not RestrictRenameAsserted Then
            '~~ When Sheet names are regarded ambigous synchronization can only take place when it is confirmed
            '~~ that only either the CodeName or the Name is changed but not both. This ensures that sheets which cannot
            '~~ be mapped between the source and the target Workbook are either obsolete or new. The mapping inability
            '~~ may indicate that both sheet names (Name and CodeName) had been changed which cannot be synchronized
            '~~ because of the missing mapping.
            Set cllButtons = mMsg.Buttons(sBttnRestricted, sBttnTrmnt, vbLf)
            sMsg.Section(3).sText = "1) Sheets of which neither the Name nor the CodeName refers to a counterpart are regarded New or Obsolete. " & _
                                    "However, this assumption is only true when never a sheet's Name  a n d  its CodeName is changed. Because " & _
                                    "this is very crucial for a programmatic syncronization it needs to be explicitely asserted - either by the " & _
                                    "corresponding argument of the syncronization service or - when this had not been provided - now in this dialog."
        Else
            Set cllButtons = mMsg.Buttons(sBttnCnfrmd, sBttnTrmnt, vbLf)
            sMsg.Section(3).sText = "2) New and Obsolete sheets had been made unambigous by the assertion that never a sheet's Name  a n d  its CodeName is changed."
        End If
        
        If ManualSynchRequired Then
            sMsg.Section(4).sText = "3) Because this synchronization service (yet) not uses a manifest for sheet design changes " & _
                                    "all these kind of syncronization issues remain a manual task. All these remaining tasks can " & _
                                    "be found in the services' log file in the target Workbook's folder."
        Else
            sMsg.Section(4).sText = vbNullString
        End If
        For Each v In dctChanged
            cllButtons.Add v
        Next v
        
        sReply = mMsg.Dsply(msg_title:="Confirm synchronization actions" _
                          , msg:=sMsg _
                          , msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sBttnTrmnt
                GoTo xt
            Case sBttnCnfrmd
                bSyncDenied = False
                Exit Do
            Case sBttnRestricted
                '~~ Collection of confirmation info is done again with this restriction now confirmed
                RestrictRenameAsserted = True
                Sync.ConfInfoClear
                Stats.Clear
                Sync.CollectAllSyncItems

            Case Else
                '~~ Display the requested changes
                Set cSource = dctChanged(sReply)
                cSource.DsplyAllChanges
        End Select
    Loop

    If Not bSyncDenied Then
        mSync.SyncTargetBackup bkp_folder, Sync.Target.FullName
        Stats.Clear
        lMode = Synchronize
        dctNameChange.RemoveAll
        If dctNameChange Is Nothing Then Set dctNameChange = New Dictionary Else dctNameChange.RemoveAll
        
        SyncReferencesNew
        SyncReferencesObsolete
        
        SyncSheetsName
        SyncSheetsCodeName
        Sync.CollectAllSyncItems ' re-collect with new names
        
        SyncSheetsNew
        '~~ When a new sheet is copied from the source to the targget Workbook any ranges referring to another
        '~~ sheet will be linked back to the source sheet. Because all sheets will be in synch here, these
        '~~ links will be dropped.
        DisconnectLinkedRanges
        Sync.CollectAllSyncItems ' re-collect with new sheets
        
        SyncSheetsObsolete
        '~~ Removing sheets will leave invalid range names behind which are now removed
        RemoveInvalidRangeNames
        
        Sync.CollectAllSyncItems ' to clear from removed sheets
        
        SyncSheetsCode
        SyncSheetsOrder
        
        SyncShapesNew
        SyncShapesObsolete
        SyncShapesProperties
        
        SyncVBCompsNew
        SyncVBCompsObsolete
        SyncVBCompsCodeChanged
        
        SyncNamesNew
        
        Set dctChanged = Nothing
        SyncTargetWithSource = True
    End If
    
xt: Set fso = Nothing
    Set dctNameChange = Nothing
    Set Sync = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Sub SyncSheetsOrder()
' ----------------------------------------------------------------------------
' Syncronize the order of the sheets in the synch target Workbook (Sync.Target)
' to appear in the same order as in the synch source Workbook (Sync.Source).
' ----------------------------------------------------------------------------
    Const PROC = "SyncSheetsOrder"
    
    On Error GoTo eh
    Dim ws          As Worksheet
    Dim SheetName   As String
    Dim i           As Long
    
    For i = 1 To Sync.Source.Worksheets.Count
        If Sync.TargetSheets.Exists(Sync.Source.Worksheets(i).Name) Then
            Set ws = Sync.Source.Worksheets(i)
            SheetName = ws.Name
            If Sync.Target.Worksheets(i).Name <> SheetName Then
                cLog.ServicedItem = ws
                If i = 1 Then
                    Sync.Target.Worksheets(SheetName).Move Before:=Sheets(i + 1)
                    cLog.Entry = "Order synched!"
                Else
                    Sync.Target.Worksheets(SheetName).Move After:=Sheets(i)
                    cLog.Entry = "Order synched!"
                End If
            End If
        End If
    Next i

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub
Private Sub SyncVBCompsCodeChanged()
' -----------------------------------------------------
' When lMode=Confirm all components which had changed
' are collected and provided for confirmation else the
' changes are syncronized.
' -----------------------------------------------------
    Const PROC = "SyncVBCompsCodeChanged"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim vbc                 As VBComponent
    Dim sCaption            As String
    
    For Each vbc In Sync.Source.VBProject.VBComponents
        If IsSheetComp(vbc) Then GoTo next_vbc
        Set cSource = New clsRaw
        Set cSource.Wrkbk = Sync.Source
        cSource.CompName = vbc.Name
        If Not cSource.Exists(Sync.Target) Then GoTo next_vbc
        
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = Sync.Target
        cTarget.CompName = vbc.Name
        cSource.CloneExpFileFullName = cTarget.ExpFileFullName
        If Not cSource.Changed Then GoTo next_vbc
        
        Stats.Count sic_non_doc_mods_code
        cLog.ServicedItem = vbc
        
        If lMode = Confirm Then
            Sync.ConfInfo = "Changed"
            sCaption = "Display changes" & vbLf & "of" & vbLf & vbLf & vbc.Name & vbLf
            If Not dctChanged.Exists(sCaption) _
            Then dctChanged.Add sCaption, cSource
        Else
            cLog.ServicedItem = vbc
            mRenew.ByImport rn_wb:=Sync.Target _
                          , rn_comp_name:=vbc.Name _
                          , rn_exp_file_full_name:=cSource.ExpFileFullName
            cLog.Entry = "Renewed/updated by import of '" & cSource.ExpFileFullName & "'"
        End If
        
        Set cTarget = Nothing
        Set cSource = Nothing
next_vbc:
    Next vbc

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function SyncVBCompsNew()
' ----------------------------------------------------
' Synchronize new components in the source Workbook
' (Sync.Source) with the target Workbook (Sync.Target).
' In lMode=Confirmation only the syncronization infos
' are collect for being confirmed.
' ----------------------------------------------------
    Const PROC = "SyncVBCompsNew"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim vbc     As VBComponent
    Dim cComp   As clsComp
    
    For Each vbc In Sync.Source.VBProject.VBComponents
        If vbc.Type = vbext_ct_Document Then GoTo next_vbc
        If vbc.Type = vbext_ct_ActiveXDesigner Then GoTo next_vbc
        
        Set cSource = New clsRaw
        Set cSource.Wrkbk = Sync.Source
        cSource.CompName = vbc.Name
        If CompExists(Sync.Target, vbc.Name) Then GoTo next_vbc
        
        '~~ No component exists under the source component's name
        cLog.ServicedItem = vbc
        Stats.Count sic_non_doc_mod_new
        
        If lMode = Confirm Then
            Sync.ConfInfo = "New! Corresponding source Workbook Export-File will by imported."
        Else
            Sync.Target.VBProject.VBComponents.Import cSource.ExpFileFullName
            cLog.Entry = "Component imported from Export-File '" & cSource.ExpFileFullName & "'"
        End If
        
        Set cSource = Nothing
next_vbc:
    Next vbc

xt: Set cComp = Nothing
    Set fso = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function SyncVBCompsObsolete()
' ---------------------------------------------------------
' Synchronize obsolete components in the source Workbook
' (Sync.Source) with the target Workbook (Sync.Target). In
' lMode=Confirm only the syncronization infos are collected
' for being confirmed.
' ---------------------------------------------------------
    Const PROC = "SyncVBCompsObsolete"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim sType       As String
    Dim cTarget     As clsComp
    
    '~~ Collect obsolete Standard Modules, Class modules, and UserForms
    For Each vbc In Sync.Target.VBProject.VBComponents
        If vbc.Type = vbext_ct_Document Then GoTo next_vbc
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = Sync.Target
        cTarget.CompName = vbc.Name
        If cTarget.Exists(Sync.Source) Then GoTo next_vbc
        
        cLog.ServicedItem = vbc
        Stats.Count sic_non_doc_mod_obsolete
        
        If lMode = Confirm Then
            Sync.ConfInfo = "Obsolete!"
        Else
            sType = cTarget.TypeString
            Sync.Target.VBProject.VBComponents.Remove vbc
            cLog.Entry = "Removed!"
        End If
        Set cTarget = Nothing
next_vbc:
    Next vbc

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function WbkGetOpen(ByVal go_wb_full_name As String) As Workbook
' ----------------------------------------------------------------------
' Returns an opened Workbook object named (go_wb_full_name) or Nothing
' when a file named (go_wb_full_name) not exists.
' ----------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If fso.FileExists(go_wb_full_name) Then
        If WbkIsOpen(wb_full_name:=go_wb_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wb_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wb_full_name)
    End If
    
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function WbkIsOpen( _
            Optional ByVal wb_base_name As String = vbNullString, _
            Optional ByVal wb_full_name As String = vbNullString) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE when a Workbook either identified by its BaseName (wb_base_name)
' or by its full name (wb_full_name) is open. When the BaseName is
' provided in the current Excel instance, else in any Excel instance.
' -------------------------------------------------------------------------
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim xlApp As Excel.Application
    
    If wb_base_name = vbNullString And wb_full_name = vbNullString Then GoTo xt
    
    With New FileSystemObject
        If wb_full_name <> vbNullString Then
            '~~ With the full name the open test spans all application instances
            If Not .FileExists(wb_full_name) Then GoTo xt
            If wb_base_name = vbNullString Then wb_base_name = .GetFileName(wb_full_name)
            On Error Resume Next
            Set xlApp = VBA.GetObject(wb_full_name).Application
            WbkIsOpen = Err.Number = 0
        Else
            On Error Resume Next
            wb_base_name = Application.Workbooks(wb_base_name).Name
            WbkIsOpen = Err.Number = 0
        End If
    End With

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

