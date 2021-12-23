Attribute VB_Name = "mSync"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSync
' Provides all services and means for the synchronization of a
' Target- with a Source-Workbook/VBProject.
'
' Public services:
'
' ----------------------------------------------------------------------------
Private Const SHEET_CONTROL_CONCAT = ": "        ' Sheet-Shape concatenator
Private Const BKP_FOLDER_PREFIX = "SyncBckp-"

Public Sync         As clsSync
Private bSyncDenied As Boolean      ' when True the synchronization is not performed
Private cSource     As clsRaw

Public Enum SyncMode
    Count = 1
    Confirm = 2
    Synchronize = 3
End Enum

Public Sub ByCodeLines( _
                 ByVal sync_target_comp_name As String, _
                 ByVal wb_source_full_name As String, _
                 ByRef sync_source_codelines As Dictionary)
' ---------------------------------------------------------
' Synchronizes the code of a Target-VBComponent
' (sync_target_comp_name) in a Target-Workbook/VB-Project
' (Sync.Target) with the code in the Export-File of the
' corresponding Source-Workbook/VB-Project's
' (wb_source_full_name) component line by line.
' ---------------------------------------------------------
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ClearLinksToSource()
' -----------------------------------------------------------------------------
' Provided all sheets, form controls and ActiveX controls had been synchronized
' there may still be references to the source Workbook for range names and
' form control OnAction specification which are to be eliminated.
' -----------------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim v           As Variant
    Dim ws          As Worksheet
    Dim shp         As Shape
    Dim nm          As name
    Dim sName       As String
    Dim sOnAction   As String
    Dim SheetName   As String
    Dim ControlName As String
    
    '~~ Clear back-referred range names
    For Each nm In Sync.Target.Names
        On Error Resume Next
        sName = Split(nm.RefersTo, "]")(1)
        If Err.Number = 0 Then
            nm.RefersTo = "=" & sName
            Log.ServicedItem = nm
            Log.Entry = "Link to source sheet removed"
        End If
    Next nm
    
    '~~ Clear OnAction configuration
    For Each v In Sync.TargetSheetControls
        SheetName = mSync.KeySheetName(v)
        ControlName = mSync.KeyControlName(v)
        Set shp = Sync.Target.Worksheets(SheetName).Shapes(ControlName)
        On Error Resume Next
        sOnAction = shp.OnAction
        sOnAction = Replace(sOnAction, Sync.Source.name, Sync.Target.name)
        shp.OnAction = sOnAction
    Next v
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Public Function KeyControlName(ByVal s As String) As String
    KeyControlName = Split(s, SHEET_CONTROL_CONCAT)(1)
End Function

Public Function KeySheetControl(ByVal sheet_name As String, _
                                ByVal control_name As String) As String
    KeySheetControl = sheet_name & SHEET_CONTROL_CONCAT & control_name
End Function

Public Function KeySheetName(ByVal s As String) As String
    KeySheetName = Split(s, SHEET_CONTROL_CONCAT)(0)
End Function

Private Function NameExists( _
                      ByRef ne_wb As Workbook, _
                      ByVal ne_nm As name) As Boolean
    Dim nm As name
    For Each nm In ne_wb.Names
        NameExists = nm.name = ne_nm.name
        If NameExists Then Exit For
    Next nm
End Function

Private Sub RemoveInvalidRangeNames()
' -----------------------------------------------------------
' Removes names which point to a range which not or no longer
' exists.
' -----------------------------------------------------------
    Const PROC = "RemoveInvalidRangeNames"
    
    On Error GoTo eh
    Dim nm As name
    For Each nm In Sync.Target.Names
        Debug.Print nm.Value
        If InStr(nm.Value, "#") <> 0 Or InStr(nm.RefersTo, "#") <> 0 Then
            Log.ServicedItem = nm
            On Error Resume Next
            nm.Delete
            Log.Entry = "Deleted! (invalid)"
        End If
    Next nm

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
        If sh.name = rs_old_name Then
            sh.name = rs_new_name
            Log.Entry = "Sheet-Name changed to '" & rs_new_name & "'"
            Exit For
        End If
    Next sh

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
                If mComp.IsWrkbkDocMod(vbc) Then
                    Log.ServicedItem = vbc
                    vbc.name = rdm_new_name
                    Log.Entry = "Renamed to '" & rdm_new_name & "'"
                    DoEvents
                    Exit For
                End If
            End If
        Next vbc
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
'
'Private Sub SheetsOrder()
'' -------------------------------------------------------
''
'' -------------------------------------------------------
'    Const PROC = "SheetsOrder"
'
'    On Error GoTo eh
'    Dim i           As Long
'    Dim wsSource    As Worksheet
'    Dim wsTarget    As Worksheet
'
'    For i = 1 To Sync.Source.Worksheets.Count
'        Set wsSource = Sync.Source.Worksheets(i)
'        Set wsTarget = Sync.Target.Worksheets(i)
'        If wsSource.Name <> wsTarget.Name Then
'            '~~ Sheet position has changed
'            If Sync.Mode = Confirm Then
'                Stop ' pending confirmation info
'            Else
'                Stop ' pending implementation
'            End If
'        End If
'    Next i
'
'xt: Exit Sub
'
'eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Private Sub CollectSyncIssuesForConfirmation()
' --------------------------------------------
' Collect all synchronization issues for being
' confirmed. Executes the synchronization
' services in 'Confirm' mode.
' ----------------------------------------------
    Const PROC = "CollectSyncIssuesForConfirmation"
    
    On Error GoTo eh
    Sync.Mode = Confirm
    
    mSyncRefs.SyncNew
    mSyncRefs.SyncObsolete
    
    mSyncSheets.SyncName
    mSyncSheets.SyncCodeName
    mSyncSheets.SyncNew
    mSyncSheets.SyncObsolete

    mSyncComps.SyncNew
    mSyncComps.SyncObsolete
    mSyncComps.SyncCodeChanges
    
    mSyncSheetCtrls.SyncControlsNew
    mSyncSheetCtrls.SyncControlsObsolete
    
    mSyncNames.SyncNew
    mSyncNames.SyncObsolete
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncRestore(ByVal sWrkbk As String)
' ------------------------------------------------------------
' Restores a synchronization target Workbook by deleting it
' and renaming the ....Backup file
' ------------------------------------------------------------
    Const PROC = "SyncRestore"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim fo          As Folder
    Dim fl          As File
    Dim sBckp       As String
    Dim lFiles      As Long
    Dim BckpFile    As File
    
    On Error Resume Next
    mCompMan.WbkGetOpen(sWrkbk).Close SaveChanges:=False
    With fso
        If .FileExists(sWrkbk & ".Backup") Then
            .CopyFile sWrkbk & ".Backup", sWrkbk
            .DeleteFile sWrkbk & ".Backup"
        End If
    End With
    
xt: Set fso = Nothing
    mCompMan.WbkGetOpen sWrkbk
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncBackup(ByVal sWrkbk As String)
' -----------------------------------------------------
' Saves a copy of the synchronization target Workbook
' (Sync.Target) under its name with a suffix .Backup
' -----------------------------------------------------
    Const PROC = "SyncBackup"
    
    On Error GoTo eh
    
    With New FileSystemObject
        .CopyFile sWrkbk, sWrkbk & ".Backup"
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
         
Public Function SyncTargetWithSource( _
          ByRef wb_target As Workbook, _
          ByRef wb_source As Workbook, _
 Optional ByVal restricted_sheet_rename_asserted As Boolean = False, _
 Optional ByVal design_rows_cols_added_or_deleted As Boolean = False) As Boolean
' --------------------------------------------------------------------
' Synchronizes a target Workbook (Sync.Target) with a source Workbook
' (Sync.Source). Returns TRUE when finished without error.
' --------------------------------------------------------------------
    Const PROC = "SyncTargetWithSource"
    
    On Error GoTo eh
    Set Sync = New clsSync
    Set Stats = New clsStats
        
    Log.Service = ErrSrc(PROC)
    mCompMan.DsplyProgress
    
    '~~ Make sure both Workbook's Export-Files are up-to-date
    wb_source.Save
    wb_target.Save
    
    Sync.ChangedClear
    Sync.RestrictRenameAsserted = restricted_sheet_rename_asserted
    Set Sync.Source = wb_source
    Set Sync.Target = wb_target
    Sync.CollectAllSyncItems
    
    Sync.ManualSynchRequired = False
    
    '~~ Count new and obsolete sheets
    Sync.Mode = Count
    mSyncSheets.SyncNew
    mSyncSheets.SyncObsolete
    
    Sync.Ambigous = True
    bSyncDenied = True

    If GetSyncIssuesConfirmed Then
        mSync.SyncBackup wb_target.FullName
        DoSynchronization design_rows_cols_added_or_deleted
        SyncTargetWithSource = True
    End If
    
xt: Set Sync = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function DoSynchronization( _
                    Optional ByVal design_rows_cols_added_or_deleted As Boolean = False) As Boolean
' -------------------------------------------------
' Perform all synch services in 'Synchronize' mode.
' -------------------------------------------------
    Const PROC = "DoSynchronization"
    
    On Error GoTo eh
    Sync.Mode = Synchronize
    Stats.Clear
    
    Sync.NameChange.RemoveAll
    
    mSyncRefs.SyncNew
    mSyncRefs.SyncObsolete
    
    mSyncSheets.SyncCodeName
    mSyncSheets.SyncName
    Sync.CollectAllSyncItems ' re-collect with new names
    
    mSyncSheets.SyncNew
    mSyncSheets.SyncObsolete
    '~~ When a new sheet is copied from the source to the targget Workbook any ranges referring to another
    '~~ sheet will be linked back to the source sheet. Because all sheets will be in synch here, these
    '~~ links will be dropped.
    Sync.CollectAllSyncItems ' re-collect with new sheets
    
    '~~ Removing sheets will leave invalid range names behind which are now removed
    RemoveInvalidRangeNames
    Sync.CollectAllSyncItems ' to clear from removed sheets
    
    mSyncSheets.SyncOrder
    Sync.CollectAllSyncItems ' to clear from removed items
    
    mSyncComps.SyncNew
    mSyncComps.SyncObsolete
    mSyncComps.SyncCodeChanges
    
    If Not design_rows_cols_added_or_deleted Then
        mSyncRanges.SyncNamedColumnsWidth
        mSyncRanges.SyncNamedRowsHeight
    End If
    
    mSyncSheetCtrls.SyncControlsNew
    mSyncSheetCtrls.SyncControlsObsolete
    Sync.CollectAllSyncItems ' to clear from removed items
    Sync.CollectAllSyncItems ' to clear from removed items
    
    mSyncSheetCtrls.SyncControlsProperties
    Sync.CollectAllSyncItems ' to clear from removed items
    
    ClearLinksToSource
    
    mSyncNames.SyncNew
    mSyncNames.SyncObsolete
    
    Sync.ChangedClear

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function GetSyncIssuesConfirmed() As Boolean
' --------------------------------------------------
' Loops CollectSyncIssuesForConfirmation and their
' display until either confirmed or terminated.
' Returns TRUE when confirmed and LASE otherwise.
' ---------------------------------------------------
    Const PROC = "GetSyncIssuesConfirmed"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim v               As Variant
    Dim sMsg            As TypeMsg
    Dim sBttnCnfrmd     As String
    Dim sBttnTrmnt      As String
    Dim sBttnRestricted As String
    Dim cllButtons      As Collection
    Dim sReply          As String
    
    Do
        '~~ Collect all synchronization issues by performing the services in 'Confirm' mode
        CollectSyncIssuesForConfirmation
        
        '~~ Display the collected synchronization issues for confirmation
        With sMsg.Section(1).Text
            .Text = Sync.ConfInfo
            .MonoSpaced = True
            .FontSize = 8.5
        End With
        sMsg.Section(2).Text.Text = "The above syncronisation issues need to be confirmed - or " & _
                                "terminated in case of any concerns!"
        
        sBttnCnfrmd = "Synchronize"
        sBttnTrmnt = "Terminate!"
        sBttnRestricted = "Confirmed" & vbLf & "that Sheet rename" & vbLf & "is restricted" & vbLf & "(either Name  o r  CodeName)"
        
        If Sync.Ambigous And Not Sync.RestrictRenameAsserted Then
            '~~ When Sheet names are regarded ambigous synchronization can only take place when it is confirmed
            '~~ that only either the CodeName or the Name is changed but not both. This ensures that sheets which cannot
            '~~ be mapped between the source and the target Workbook are either obsolete or new. The mapping inability
            '~~ may indicate that both sheet names (Name and CodeName) had been changed which cannot be synchronized
            '~~ because of the missing mapping.
            mMsg.Buttons cllButtons, sBttnRestricted, sBttnTrmnt, vbLf
            sMsg.Section(3).Text.Text = "1) Sheets in the source Workbbook of which neither the Name nor the CodeName refers to a counterpart in the target Workbook " & _
                                    "are regarded  " & mBasic.Spaced("new") & ". Sheets in the target Workbook in contrast are regarded  " & mBasic.Spaced("obsolete.") & _
                                    "  However, this assumption only holds true when  " & mBasic.Spaced("never") & "  a sheet's Name  a n d  its CodeName is changed. " & _
                                    "Because this is absolutely crucial for this syncronization it needs to be explicitely  " & mBasic.Spaced("asserted.")
        Else
            mMsg.Buttons cllButtons, sBttnCnfrmd, sBttnTrmnt, vbLf
            sMsg.Section(3).Text.Text = "2) New and Obsolete sheets had been made unambigous by the assertion that never a sheet's Name  a n d  its CodeName is changed."
        End If
        
        If Sync.ManualSynchRequired Then
            sMsg.Section(4).Text.Text = "3) Because this synchronization service (yet) not uses a manifest for sheet design changes " & _
                                    "all these kind of syncronization issues remain a manual task. All these remaining tasks can " & _
                                    "be found in the services' log file in the target Workbook's folder."
        Else
            sMsg.Section(4).Text.Text = vbNullString
        End If
        For Each v In Sync.Changed
            cllButtons.Add v
        Next v
        
        sReply = mMsg.Dsply(dsply_title:="Confirm the below synchronization issues" _
                          , dsply_msg:=sMsg _
                          , dsply_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sBttnTrmnt
                GoTo xt
            Case sBttnCnfrmd
                GetSyncIssuesConfirmed = True
                GoTo xt
            Case sBttnRestricted
                '~~ Collection of confirmation info is done again with this restriction now confirmed
                Sync.RestrictRenameAsserted = True
                Sync.ConfInfoClear
                Stats.Clear
                Sync.CollectAllSyncItems
            Case Else
                '~~ Display the requested changes
                Set cSource = Sync.Changed(sReply)
                cSource.DsplyAllChanges
        End Select
    Loop
    
xt: Set fso = Nothing
'    Set Sync = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
            wb_base_name = Application.Workbooks(wb_base_name).name
            WbkIsOpen = Err.Number = 0
        End If
    End With

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

