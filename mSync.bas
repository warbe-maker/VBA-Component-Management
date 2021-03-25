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
Private Const SHEET_CONTROL_CONCAT = ": "        ' Sheet-Shape concatenator
Private Const BKP_FOLDER_PREFIX = "SyncBckp-"

Public Sync                     As clsSync
Private bSyncDenied             As Boolean      ' when True the synchronization is not performed
Private cSource                 As clsRaw

Public Enum SyncMode
    Count = 1
    Confirm = 2
    Synchronize = 3
End Enum

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
                      ByVal ne_nm As Name) As Boolean
    Dim nm As Name
    For Each nm In ne_wb.Names
        NameExists = nm.Name = ne_nm.Name
        If NameExists Then Exit For
    Next nm
End Function

Private Sub RemoveInvalidRangeNames()
' -----------------------------------------------------------
' Removes names which point to a range which not or no longer
' exists.
' -----------------------------------------------------------
    Dim nm As Name
    For Each nm In Sync.Target.Names
        Debug.Print nm.Value
        If InStr(nm.Value, "#") <> 0 Or InStr(nm.RefersTo, "#") <> 0 Then
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

Public Function SheetCodeName( _
                        ByRef sync_wb As Workbook, _
                        ByVal sync_sheet_name As String) As String
' ----------------------------------------------------------------
'
' ----------------------------------------------------------------
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
            If Sync.Mode = Confirm Then
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

Private Sub SyncConfirmation()
' ------------------------------------------------------------
' Collect all confirmation information regarding sheet changes
' ------------------------------------------------------------
    Sync.Mode = Confirm
    mSyncRefs.SyncNew
    mSyncRefs.SyncObsolete
    
    mSyncSheets.SyncName
    mSyncSheets.SyncCodeName
    mSyncSheets.SyncNew
    mSyncSheets.SyncObsolete
    mSyncSheets.SyncCode
    
    mSyncSheetCtrls.SyncShapesNew
    mSyncSheetCtrls.SyncShapesObsolete
    
    mSyncComps.SyncNew
    mSyncComps.SyncObsolete
    mSyncComps.SyncCodeChanges
    
    mSyncNames.SyncNew
    
End Sub

Public Sub SyncRestore( _
        Optional ByVal backup_folder As String = vbNullString)
' ------------------------------------------------------------
' Restores a synchronization target Workbook with its backup.
' ! The backup folder has just file with the name of the     !
' ! file to restore in the backup folder's parent folder.    !
' The backup folder is selected when not provided. When the
' selected or provided folder is not a backup folder the
' service terminates without notice.
' ------------------------------------------------------------
    Const PROC = "SyncRestore"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim fo      As Folder
    Dim fl      As File
    Dim sBckp   As String
    Dim lFiles  As Long
        
    sBckp = backup_folder
    
    '~~ Select the desired backup folder when none is provided
    If sBckp = vbNullString Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Select the desired backup folder"
            .AllowMultiSelect = False
            .InitialFileName = mMe.ServicedRoot
            If .show = -1 Then ' if OK is pressed
                sBckp = .SelectedItems(1)
            End If
        End With
    End If
    
    '~~ When the provided or selected folder is not a backup folder terminate without notice
    If InStr(1, sBckp, BKP_FOLDER_PREFIX, vbTextCompare) = 0 Then
        Application.StatusBar = "Restore service denied! The selected or provided folder's name not begins with '" & BKP_FOLDER_PREFIX & "'"
        GoTo xt
    End If
    
    With fso
        For Each fl In .GetFolder(sBckp).Files
            lFiles = lFiles + 1
        Next fl
        
        '~~ When the backup folder has more then 1 file terminate without notice
        If lFiles > 1 Then
            Application.StatusBar = "Restore service denied! The selected backup folder contains more than one (the backup) file"
            GoTo xt
        End If
        .CopyFile fl.Path, .GetParentFolderName(sBckp)
        
        '~~ Remove all backup folders
        For Each fo In .GetFolder(.GetParentFolderName(sBckp)).SubFolders
            If InStr(1, fo.Path, "\" & BKP_FOLDER_PREFIX, vbTextCompare) = 0 Then .DeleteFolder fo.Path
        Next fo
    End With
    
xt: Set fso = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub SyncBackup(ByRef bkp_folder As String, _
                            ByVal sTarget As String)
' -----------------------------------------------------
' Saves a copy of the synchronization target Workbook
' (Sync.Target) in a time-stamped folder under the
' Workbook folder returned in (bkp_folder).
' -----------------------------------------------------
    Const PROC = "SyncBackup"
    
    On Error GoTo eh
    Dim BckpFolderName  As String
    Dim fo              As Folder
    
    BckpFolderName = BKP_FOLDER_PREFIX & Format$(Now(), "YYMMDD-hhmmss")
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

    Do
        
        '~~ Collect all synchronization info and get them confirmed
        SyncConfirmation
        
        '~~ Get the collected info confirmed
        sMsg.Section(1).sText = Sync.ConfInfo
        sMsg.Section(1).bMonspaced = True
        sMsg.Section(2).sText = "The above syncronisation issues need to be confirmed - or " & _
                                "terminated in case of any concerns!"
        
        sBttnCnfrmd = "Synchronize" & vbLf & vbLf & fso.GetBaseName(Sync.Target.Name) & vbLf & " with " & vbLf & fso.GetBaseName(Sync.Source.Name)
        sBttnTrmnt = "Terminate!" & vbLf & vbLf & "Synchronization denied" & vbLf & "because of concerns"
        sBttnRestricted = "Confirmed" & vbLf & "that Sheet rename" & vbLf & "is restricted" & vbLf & "(either Name  o r  CodeName)"
        
        If Sync.Ambigous And Not Sync.RestrictRenameAsserted Then
            '~~ When Sheet names are regarded ambigous synchronization can only take place when it is confirmed
            '~~ that only either the CodeName or the Name is changed but not both. This ensures that sheets which cannot
            '~~ be mapped between the source and the target Workbook are either obsolete or new. The mapping inability
            '~~ may indicate that both sheet names (Name and CodeName) had been changed which cannot be synchronized
            '~~ because of the missing mapping.
            Set cllButtons = mMsg.Buttons(sBttnRestricted, sBttnTrmnt, vbLf)
            sMsg.Section(3).sText = "1) Sheets in the source Workbbook of which neither the Name nor the CodeName refers to a counterpart in the target Workbook " & _
                                    "are regarded  " & mBasic.Spaced("new") & ". Sheets in the target Workbook in contrast are regarded  " & mBasic.Spaced("obsolete.") & _
                                    "  However, this assumption only holds true when  " & mBasic.Spaced("never") & "  a sheet's Name  a n d  its CodeName is changed. " & _
                                    "Because this is absolutely crucial for this syncronization it needs to be explicitely  " & mBasic.Spaced("asserted.")
        Else
            Set cllButtons = mMsg.Buttons(sBttnCnfrmd, sBttnTrmnt, vbLf)
            sMsg.Section(3).sText = "2) New and Obsolete sheets had been made unambigous by the assertion that never a sheet's Name  a n d  its CodeName is changed."
        End If
        
        If Sync.ManualSynchRequired Then
            sMsg.Section(4).sText = "3) Because this synchronization service (yet) not uses a manifest for sheet design changes " & _
                                    "all these kind of syncronization issues remain a manual task. All these remaining tasks can " & _
                                    "be found in the services' log file in the target Workbook's folder."
        Else
            sMsg.Section(4).sText = vbNullString
        End If
        For Each v In Sync.Changed
            cllButtons.Add v
        Next v
        
        sReply = mMsg.Dsply(msg_title:="Confirm the below synchronization issues" _
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

    If Not bSyncDenied Then
        mSync.SyncBackup bkp_folder, Sync.Target.FullName
        Stats.Clear
        Sync.Mode = Synchronize
        Sync.NameChange.RemoveAll
        
        mSyncRefs.SyncNew
        mSyncRefs.SyncObsolete
        
        mSyncSheets.SyncName
        mSyncSheets.SyncCodeName
        Sync.CollectAllSyncItems ' re-collect with new names
        
        mSyncSheets.SyncNew
        '~~ When a new sheet is copied from the source to the targget Workbook any ranges referring to another
        '~~ sheet will be linked back to the source sheet. Because all sheets will be in synch here, these
        '~~ links will be dropped.
        DisconnectLinkedRanges
        Sync.CollectAllSyncItems ' re-collect with new sheets
        
        mSyncSheets.SyncObsolete
        '~~ Removing sheets will leave invalid range names behind which are now removed
        RemoveInvalidRangeNames
        Sync.CollectAllSyncItems ' to clear from removed sheets
        
        mSyncSheets.SyncCode
        mSyncSheets.SyncOrder
        
        mSyncSheetCtrls.SyncShapesNew
        mSyncSheetCtrls.SyncShapesObsolete
        mSyncSheetCtrls.SyncShapesProperties
        Sync.CollectAllSyncItems ' to clear from removed items
        
        mSyncSheetCtrls.SyncOOBsNew
        mSyncSheetCtrls.SyncOOBsObsolete
        mSyncSheetCtrls.SyncOOBsProperties
        Sync.CollectAllSyncItems ' to clear from removed items
        
        mSyncComps.SyncNew
        mSyncComps.SyncObsolete
        mSyncComps.SyncCodeChanges
        
        mSyncNames.SyncNew
        
        Sync.ChangedClear
        SyncTargetWithSource = True
    End If
    
xt: Set fso = Nothing
    Set Sync = Nothing
    Exit Function

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

