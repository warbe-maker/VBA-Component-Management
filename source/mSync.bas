Attribute VB_Name = "mSync"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSync
'
' Provides all services and means for the synchronization of a Target- with
' its corresponding Source-Workbook.
'
' Note: All Run.... services are invoked via Application.Run
'
'
' W. Rauschenberger, Berlin June 2022
' ----------------------------------------------------------------------------
Private Const SHEET_CONTROL_CONCAT  As String = ": "                                ' Sheet-Shape concatenator
Private Const WB_SYNC_TARGET        As String = "_CompMan_Sync_Target"              ' name suffix for the sync target working copy

Public Const TITLE_SYNC_REFS        As String = "Synchronize VB-Project References" ' Identifies the synchronization dialog

Public Sync                         As clsSync
Private bSyncDenied                 As Boolean      ' when True the synchronization is not performed
Private cSource                     As clsComp

Public Enum SyncMode
    Count = 1
    Confirm = 2
    Synchronize = 3
End Enum

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function TheSyncTargetsSource(ByVal sts_target_wb As Workbook) As String
' ----------------------------------------------------------------------------
' Returns the full name of the Workbook which is considered the
' 'Syncronization-Target-Workbook's' corresponding source. A vbNullString is
' returne when it cannot be identified in CompMan's 'Serviced-Folder.
' ----------------------------------------------------------------------------
    Const PROC = "TheSyncTargetsSource"
    
    On Error GoTo eh
    Dim cll As Collection
    
    If mFile.Exists(ex_folder:=mConfig.FolderServiced, ex_file:=mSync.SyncTargetOriginName(sts_target_wb), ex_result_files:=cll) Then
        If cll.Count <> 1 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "Unable to identify the 'Synchronization-Target-Workbook's corresponding 'Synchronization-Source-Workbook' because the " & _
                                                "'Serviced-Folder' contains either none or more than one file named '" & mSync.SyncTargetOriginName(sts_target_wb) & "'"
        TheSyncTargetsSource = cll(1).Path
    End If

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsSyncTargetCopy(ByVal itc_wb As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Workbook (itc_wb) is a synchornization target Workbook
' working copy.
' ----------------------------------------------------------------------------
    IsSyncTargetCopy = InStr(itc_wb.Name, WB_SYNC_TARGET) <> 0
End Function

Public Function SyncTargetOriginName(ByVal ston_wb As Workbook) As String
' ----------------------------------------------------------------------------
' Returns serviced 'Synchronization-Target-Workbook's (ston_wb) origin name.
' ----------------------------------------------------------------------------
    SyncTargetOriginName = Replace(ston_wb.Name, WB_SYNC_TARGET, vbNullString)
    
End Function

Public Function SaveAsSyncTargetWorkingCopy(ByVal satc_wb As Workbook) As Workbook
' ----------------------------------------------------------------------------
' Saves the Workbook (satc_wb) as 'Synchronization-Target-Workbook' working
' copy.
' ----------------------------------------------------------------------------
    
    Dim fso     As New FileSystemObject
    Dim ext     As String
    Dim sCopyAs As String
    
    If Not IsSyncTargetCopy(satc_wb) Then
        ext = fso.GetExtensionName(satc_wb.FullName)
        sCopyAs = satc_wb.Path & Replace(satc_wb.Name, "." & ext, vbNullString) & WB_SYNC_TARGET & "." & ext
        mFile.Delete sCopyAs
        satc_wb.SaveAs FileName:=sCopyAs, AccessMode:=xlExclusive
        Set SaveAsSyncTargetWorkingCopy = mWbk.GetOpen(sCopyAs)
    End If
    Set fso = Nothing
    
End Function

Public Sub ByCodeLines( _
                 ByVal sync_target_comp_name As String, _
                 ByVal wb_source_full_name As String, _
                 ByRef sync_source_codelines As Dictionary)
' ----------------------------------------------------------------------------
' Synchronizes the code of a Target-VBComponent
' (sync_target_comp_name) in a Target-Workbook/VB-Project
' (Sync.TargetWb) with the code in the Export-File of the
' corresponding Source-Workbook/VB-Project's
' (wb_source_full_name) component line by line.
' ----------------------------------------------------------------------------
    Const PROC = "ByCodeLines"

    On Error GoTo eh
    Dim i       As Long: i = 1
    Dim v       As Variant
    Dim ws      As Worksheet
    Dim wbRaw   As Workbook
    Dim cSource As clsComp
    
    If sync_source_codelines Is Nothing Then
        '~~ Obtain non provided code lines for the line by line syncronization
        Set wbRaw = WbkGetOpen(wb_source_full_name)
        Set cSource = New clsComp
        With cSource
            Set .Wrkbk = wbRaw
            .CompName = sync_target_comp_name
            Set sync_source_codelines = .CodeLines
        End With
    End If
    
    With Sync.TargetWb.VBProject.VBComponents(sync_target_comp_name).CodeModule
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    Dim shp         As Shape
    Dim nm          As Name
    Dim sName       As String
    Dim sOnAction   As String
    Dim SheetName   As String
    Dim ControlName As String
    
    '~~ Clear back-referred range names
    For Each nm In Sync.TargetWb.Names
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
        Set shp = Sync.TargetWb.Worksheets(SheetName).Shapes(ControlName)
        On Error Resume Next
        sOnAction = shp.OnAction
        sOnAction = Replace(sOnAction, Sync.SourceWb.Name, Sync.TargetWb.Name)
        shp.OnAction = sOnAction
    Next v
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CollectSyncIssuesForApplicationRun()
' ------------------------------------------------------------------------------
' Collects all synchronization issues and displays them in a modeless dialog
' for being performed by a click on the corresponding command button. This
' procedure is repeated until there os no synchronization still outstanding.
' ------------------------------------------------------------------------------
    Const PROC = "CollectSyncIssuesForApplicationRun"
    
    On Error GoTo eh
    Sync.Mode = Confirm
        
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CollectSyncIssuesForConfirmation()
' --------------------------------------------
' Collect all synchronization issues for being
' confirmed. Executes the synchronization
' services in 'Confirm' mode.
' ----------------------------------------------
    Const PROC = "CollectSyncIssuesForConfirmation"
    
    On Error GoTo eh
    Sync.Mode = Confirm
        
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function DoSynchronization(Optional ByVal design_rows_cols_added_or_deleted As Boolean = False) As Boolean
' -------------------------------------------------
' Perform all synch services in 'Synchronize' mode.
' -------------------------------------------------
    Const PROC = "DoSynchronization"
    
    On Error GoTo eh
    Sync.Mode = Synchronize
    Stats.Clear
    
    Sync.NameChange.RemoveAll
       
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Private Function GetSyncIssuesConfirmed() As Boolean
' ------------------------------------------------------------------------------
' Loops CollectSyncIssuesForConfirmation and their display until either
' confirmed or terminated. Returns TRUE when confirmed and FALSE otherwise.
' ------------------------------------------------------------------------------
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
            Set cllButtons = mMsg.Buttons(sBttnRestricted, sBttnTrmnt, vbLf)
            sMsg.Section(3).Text.Text = "1) Sheets in the source Workbbook of which neither the Name nor the CodeName refers to a counterpart in the target Workbook " & _
                                    "are regarded  " & mBasic.Spaced("new") & ". Sheets in the target Workbook in contrast are regarded  " & mBasic.Spaced("obsolete.") & _
                                    "  However, this assumption only holds true when  " & mBasic.Spaced("never") & "  a sheet's Name  a n d  its CodeName is changed. " & _
                                    "Because this is absolutely crucial for this syncronization it needs to be explicitely  " & mBasic.Spaced("asserted.")
        Else
            Set cllButtons = mMsg.Buttons(sBttnCnfrmd, sBttnTrmnt, vbLf)
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
        
        If Not mMsg.IsValidMsgButtonsArg(cllButtons) Then Stop
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    Const PROC = "RemoveInvalidRangeNames"
    
    On Error GoTo eh
    Dim nm As Name
    For Each nm In Sync.TargetWb.Names
        Debug.Print nm.Value
        If InStr(nm.Value, "#") <> 0 Or InStr(nm.RefersTo, "#") <> 0 Then
            Log.ServicedItem = nm
            On Error Resume Next
            nm.Delete
            Log.Entry = "Deleted! (invalid)"
        End If
    Next nm

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
        If sh.Name = rs_old_name Then
            sh.Name = rs_new_name
            Log.Entry = "Sheet-Name changed to '" & rs_new_name & "'"
            Exit For
        End If
    Next sh

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
                    vbc.Name = rdm_new_name
                    Log.Entry = "Renamed to '" & rdm_new_name
                    Exit For
                End If
            End If
        Next vbc
    End With
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
'    For i = 1 To Sync.SourceWb.Worksheets.Count
'        Set wsSource = Sync.SourceWb.Worksheets(i)
'        Set wsTarget = Sync.TargetWb.Worksheets(i)
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
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Public Sub SyncBackup(ByVal sWbSync As String)
' ------------------------------------------------------------------------------
' Creates a backup for the 'Synchronization-Workbook' (sWbSync) by copying
' over an existing backup file or creating one.
' When the 'Synchronization-Workbook' (sWbSync) was initially open it is
' reopened after backup.
' ------------------------------------------------------------------------------
    Const PROC = "SyncBackup"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim wb      As Workbook
    Dim WasOpen As Boolean
    
    With fso
        If .FileExists(sWbSync) Then
            If mWbk.IsOpen(sWbSync, wb) Then
                WasOpen = True
                wb.Close SaveChanges:=False
            End If
        
            .CopyFile sWbSync, sWbSync & ".Backup"
            If WasOpen Then mWbk.GetOpen sWbSync
        End If
    End With

xt: Set fso = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncRestore(ByVal sWbSync As String)
' ------------------------------------------------------------------------------
' Restores the 'Synchronization-Workbook' (sWbSync) - provided a backup exists -
' by copying the backup file which overwrites an existing file or creates
' one if not existing. The backup is deleted. When the Workbook was initially
' open it is reopened after restore.
' ------------------------------------------------------------------------------
    Const PROC = "SyncRestore"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim wb      As Workbook
    Dim WasOpen As Boolean
    
    On Error Resume Next
    With fso
        If .FileExists(sWbSync) Then
            If mWbk.IsOpen(sWbSync, wb) Then
                WasOpen = True
                wb.Close SaveChanges:=False
            End If
        End If
    
        If .FileExists(sWbSync & ".Backup") Then
            .CopyFile sWbSync & ".Backup", sWbSync ' copy with overwrite
            .DeleteFile sWbSync & ".Backup"
        End If
    
        If WasOpen Then mWbk.GetOpen sWbSync
    End With
    
xt: Set fso = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
        
Public Function SyncTargetWithSource(ByRef wb_target As Workbook, _
                                     ByRef wb_source As Workbook, _
                            Optional ByVal restricted_sheet_rename_asserted As Boolean = False, _
                            Optional ByVal design_rows_cols_added_or_deleted As Boolean = False) As Boolean
' ------------------------------------------------------------------------------
' Synchronizes the 'Synchronization-Target-Workbook (wb_target) with the
' 'Synchronization_source-Workbook and returns TRUE when finished without error.
' ------------------------------------------------------------------------------
    Const PROC = "SyncTargetWithSource"
    
    On Error GoTo eh
    Set Sync = New clsSync
    Set Stats = New clsStats
        
    Log.Service = SRVC_SYNC_WORKBOOKS
    mService.Progress Log.Service
    
    With Sync
        .ChangedClear
        .RestrictRenameAsserted = restricted_sheet_rename_asserted
        Set .SourceWb = wb_source
        Set .TargetWb = wb_target
        .CollectAllSyncItems
        .ManualSynchRequired = False
        
        '~~ Count new and obsolete sheets
        .Mode = Count
        mSyncSheets.SyncNew
        mSyncSheets.SyncObsolete
        .Ambigous = True
    End With
    
    bSyncDenied = True

    If GetSyncIssuesConfirmed Then
        mSync.SyncBackup wb_target.FullName
        DoSynchronization design_rows_cols_added_or_deleted
        SyncTargetWithSource = True
    End If
    
xt: Set Sync = Nothing
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub RunSync(ByVal wb_target As Workbook, _
                   ByVal wb_source As Workbook)
' ------------------------------------------------------------------------------
' Synchronizes the 'Synchronization-Target-Workbook (wb_target) with its corres-
' ponding (same named) 'Synchronization-Source-Workbook. Items rynchronized are:
' - References
' - Sheet-Code-Lines
' - Sheet-Code-Names
' - Sheet-Names
' - Sheet-Objects
' - Sheet-Order
' - Modules-Code-Lines
' - Workbook-Code-Lines
' ------------------------------------------------------------------------------
    Const PROC = "RunSync"
    
    On Error GoTo eh
    Dim cllObsolete As New Collection
    Dim cllNew      As New Collection
    
    '~~ Collect to be synced References and display them in mode-less dialog
    '~~ Once initiated, the dialog loops until all References are synchronized or the dialog is teminated
    mSyncRefs.SyncRefs wb_target, wb_source
    
xt: Set cllNew = Nothing
    Set cllObsolete = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsOpen( _
            Optional ByVal wb_name As String = vbNullString, _
            Optional ByVal wb_full_name As String = vbNullString) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE when a Workbook either identified by its BaseName (wb_base_name)
' or by its full name (wb_full_name) is open. When the BaseName is
' provided in the current Excel instance, else in any Excel instance.
' -------------------------------------------------------------------------
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim xlApp As Excel.Application
    
    If wb_name = vbNullString And wb_full_name = vbNullString Then GoTo xt
    
    With New FileSystemObject
        If wb_full_name <> vbNullString Then
            '~~ With the full name the open test spans all application instances
            If Not .FileExists(wb_full_name) Then GoTo xt
            If wb_name = vbNullString Then wb_name = .GetFileName(wb_full_name)
            On Error Resume Next
            Set xlApp = VBA.GetObject(wb_full_name).Application
            WbkIsOpen = Err.Number = 0
        Else
            On Error Resume Next
            wb_name = Application.Workbooks(wb_name).Name
            WbkIsOpen = Err.Number = 0
        End If
    End With

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

