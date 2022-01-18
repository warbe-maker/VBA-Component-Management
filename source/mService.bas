Attribute VB_Name = "mService"
Option Explicit

Public Const SERVICES_LOG_FILE = "CompMan.Services.log"
Private wbServiced  As Workbook     ' The service Workbook throughout all services

Public Property Get Serviced() As Workbook
    Set Serviced = wbServiced
End Property

Public Property Set Serviced(ByRef wb As Workbook)
    Set wbServiced = wb
End Property

Public Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub RenewComp( _
      Optional ByVal rc_exp_file_full_name As String = vbNullString, _
      Optional ByVal rc_comp_name As String = vbNullString, _
      Optional ByRef rc_wb As Workbook = Nothing)
' --------------------------------------------------------------------
' This service renews a component, either specified via its name
' (rc_comp_name) or via the Export-File (rc_exp_file_full_name), in
' the Workbook (rc_wb) which defaults to the ActiveWorkbook.
' - When the provided Export-File (rc_exp_file_full_name) does exist
'   but a component name has been provided a file selection dialog is
'   displayed with the possible files already filtered.
' - When no Export-File is selected the service terminates without
'   notice.
' - When the Workbook (rc_wb) is omitted it defaults to the
'   ActiveWorkbook.
' - When the ActiveWorkbook or the provided Workbook is ThisWorkbook
'   the service terminates without notice
'
' Usage ways
' - When the Addin instance of this Workbook is setup/open from the
'   immediate window:
'   Application.Run "CompMan.xlam!mService.RenewComp"
'
'   This way allows to renew any component in any Workbook including
'   this development instance
' - When only the development instance is open, from the immediate
'   window:
'   Application.Run "CompMan.xlsb!mService.RenewComp"
'
' W. Rauschenberger Berlin, Jan 2021
' --------------------------------------------------------------------
    Const PROC = "RenewComp"

    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim Comp        As New clsComp
    Dim flFile      As File
    Dim wbTemp      As Workbook
    Dim wbActive    As Workbook
    Dim sBaseName   As String
    
    mBasic.BoP ErrSrc(PROC)
    If rc_wb Is Nothing Then Set rc_wb = ActiveWorkbook
    If Log Is Nothing Then Set Log = New clsLog
    
    Comp.Wrkbk = rc_wb
    If rc_exp_file_full_name <> vbNullString Then
        If Not fso.FileExists(rc_exp_file_full_name) Then
            rc_exp_file_full_name = vbNullString ' enforces selection when the component name is also not provided
        End If
    End If
    
    If rc_comp_name <> vbNullString Then
        Comp.CompName = rc_comp_name
        If Not Comp.Exists Then
            If rc_exp_file_full_name <> vbNullString Then
                rc_comp_name = fso.GetBaseName(rc_exp_file_full_name)
            End If
        End If
    End If
    
    If ThisWorkbook Is rc_wb Then
        Debug.Print "The service '" & ErrSrc(PROC) & "' cannot run when ThisWorkbook is identical with the ActiveWorkbook!"
        GoTo xt
    End If
    
    If rc_exp_file_full_name = vbNullString _
    And rc_comp_name = vbNullString Then
        '~~ ---------------------------------------------
        '~~ Select the Export-File for the re-new service
        '~~ of which the base name will be regared as the component to be renewed.
        '~~ --------------------------------------------------------
        If mFile.SelectFile(sel_init_path:=mExport.ExpFileFolderPath(rc_wb) _
                          , sel_filters:="*.bas,*.cls,*.frm" _
                          , sel_filter_name:="File" _
                          , sel_title:="Select the Export-File for the re-new service" _
                          , sel_result:=flFile) _
        Then rc_exp_file_full_name = flFile.Path
    End If
    
    If rc_comp_name <> vbNullString _
    And rc_exp_file_full_name = vbNullString Then
        Comp.CompName = rc_comp_name
        '~~ ------------------------------------------------
        '~~ Select the component's corresponding Export-File
        '~~ ------------------------------------------------
        sBaseName = fso.GetBaseName(rc_exp_file_full_name)
        '~~ Select the Export-File for the re-new service
        If mFile.SelectFile(sel_init_path:=mExport.ExpFileFolderPath(rc_wb) _
                          , sel_filters:="*" & Comp.ExpFileExt _
                          , sel_filter_name:="File" _
                          , sel_title:="Select the Export-File for the provided component '" & rc_comp_name & "'!" _
                          , sel_result:=flFile) _
        Then rc_exp_file_full_name = flFile.Path
    End If
    
    If rc_exp_file_full_name = vbNullString Then
        MsgBox Title:="Service '" & ErrSrc(PROC) & "' will be aborted!" _
             , Prompt:="Service '" & ErrSrc(PROC) & "' will be aborted because no " & _
                       "existing Export-File has been provided!" _
             , Buttons:=vbOKOnly
        GoTo xt ' no Export-File selected
    End If
    
    With Comp
        Log.ServicedItem = .VBComp
        If rc_comp_name <> vbNullString Then
            If fso.GetBaseName(rc_exp_file_full_name) <> rc_comp_name Then
                MsgBox Title:="Service '" & ErrSrc(PROC) & "' will be aborted!" _
                     , Prompt:="Service '" & ErrSrc(PROC) & "' will be aborted because the " & _
                               "Export-File '" & rc_exp_file_full_name & "' and the component name " & _
                               "'" & rc_comp_name & "' do not indicate the same component!" _
                     , Buttons:=vbOKOnly
                GoTo xt
            End If
            .CompName = rc_comp_name
        Else
            .CompName = fso.GetBaseName(rc_exp_file_full_name)
        End If
        
        If .Wrkbk Is ActiveWorkbook Then
            Set wbActive = ActiveWorkbook
            Set wbTemp = Workbooks.Add ' Activates a temporary Workbook
            Log.Entry = "Active Workbook de-activated by creating a temporary Workbook"
        End If
            
        mRenew.ByImport rn_wb:=.Wrkbk _
             , rn_comp_name:=.CompName _
             , rn_raw_exp_file_full_name:=rc_exp_file_full_name
        Log.Entry = "Component renewed/updated by (re-)import of '" & rc_exp_file_full_name & "'"
    End With
    
xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        Log.Entry = "Temporary created Workbook closed without save"
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            Log.Entry = "De-activated Workbook '" & wbActive.name & "' re-activated"
            Set wbActive = Nothing
        End If
    End If
    Set Comp = Nothing
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Continue()
' -------------------------------------------
' Continues the paused CompMan Addin Services
' -------------------------------------------
    If mMe.IsDevInstnc Then
        mConfig.CompManAddinIsPaused = False
        mMe.DisplayStatus
    End If
End Sub


Private Function CodeModuleIsEmpty(ByRef vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0
        If Not CodeModuleIsEmpty Then CodeModuleIsEmpty = .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Public Function CompMaxLen(ByRef ml_wb As Workbook) As Long
    Dim vbc As VBComponent
    For Each vbc In ml_wb.VBProject.VBComponents
        CompMaxLen = mBasic.Max(CompMaxLen, Len(vbc.name))
    Next vbc
End Function

Public Function Denied(ByVal den_service As String) As Boolean
' --------------------------------------------------------------------------
' Returns TRUE when all preconditions for a service execution are fulfilled.
' --------------------------------------------------------------------------
    Const PROC = "Denied"
    
    On Error GoTo eh
    Dim sStatus As String
    
    If Log Is Nothing Then Set Log = New clsLog
    Log.Service = den_service
    
    If mMe.IsAddinInstnc And mMe.CompManAddinIsPaused Then
        '~~ When the service is about to be provided by the Addin but the Addin is currently paused
        '~~ another try with the serviced provided by the open Development instance may do the job.
        sStatus = "The CompMan Addin is currently paused. Open the development instance and retry."
    ElseIf WbkIsRestoredBySystem Then
        sStatus = "Service denied! Workbook appears restored by the system!"
    ElseIf mConfig.CompManAddinIsPaused And mMe.IsAddinInstnc And den_service = "ExportChangedComponents" Then
        '~~ Note: Export of code changes can be made by the development instance itself
        sStatus = "Service denied! The CompMan Addin is currently paused!"
    ElseIf FolderNotVbProjectExclusive Then
        sStatus = "Service denied! The Workbook is not the only one in its parent folder!"
    ElseIf Not mCompMan.WinMergeIsInstalled Then
        sStatus = "Service denied! WinMerge is required but not installed!"
    End If
    
xt: If sStatus <> vbNullString Then
        Log.Entry = sStatus
        Application.StatusBar = Log.Service & sStatus
        Denied = True
    End If
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mService." & s
End Function

Public Sub ExportChangedComponents(ByVal hosted As String)
' -----------------------------------------------------------
' Exclusively called by mCompMan.ExportChangedComponents,
' triggered by the Before_Save event.
' - Any VBComponent the code has changed (temporary
'   Export-File differs from the current Export-File or no
'   Export-File exists) exported
' - Outdated Export-Files (components no longer existing) are
'   removed
' - Modified 'Clone-Components' require a confirmation by the
'   user.
'
' Note: For the case the Export-Folder may have changed any
'       Export-File found within the Workbook-Folder outside
'       the specified Export-Folder is removed.
' Attention: This procedure is called exclusively by
'            mCompMan.UpdateOutdatedCommonComponents! When called directly
'            by the user, e.g. via the 'Imediate Window' an
'            error will be raised because an 'mService.Serviced'
'            Workbook is not set.
' --------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
'    Dim fso         As New FileSystemObject
'    Dim Comp        As clsComp
'    Dim RawComp     As clsRaw
    
    mBasic.BoP ErrSrc(PROC)
    If mService.Serviced Is Nothing _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The procedure '" & ErrSrc(PROC) & "' has been called without a prior set of the 'Serviced' Workbook. " & _
                                                 "(it may have been called directly via the 'Immediate Window'"
    If mService.Denied(PROC) Then GoTo xt
    mCompMan.ManageRawCommonComponentsProperties hosted
    
    mExport.ChangedComponents
        
xt: Set dctHostedRaws = Nothing
'    Set Comp = Nothing
'    Set RawComp = Nothing
    Set Log = Nothing
'    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function FolderNotVbProjectExclusive() As Boolean

    Dim fso As New FileSystemObject
    Dim fl  As File
    
    For Each fl In fso.GetFolder(mService.Serviced.Path).Files
        If VBA.Left$(fso.GetFileName(fl.Path), 2) = "~$" Then GoTo next_fl
        If VBA.StrComp(fl.Path, mService.Serviced.FullName, vbTextCompare) <> 0 Then
            Select Case fso.GetExtensionName(fl.Path)
                Case "xlsm", "xlam", "xlsb" ' may contain macros, a VB-Project repsectively
                    FolderNotVbProjectExclusive = True
                    Exit For
            End Select
        End If
next_fl:
    Next fl

End Function

Public Function SyncVBProjects( _
                Optional ByRef wb_target As Workbook = Nothing, _
                Optional ByVal wb_source_name As String = vbNullString, _
                Optional ByVal restricted_sheet_rename_asserted As Boolean = False, _
                Optional ByVal design_rows_cols_added_or_deleted As Boolean = False) As Boolean
' --------------------------------------------------------------------------------------------
' Synchronizes the target Workbook (wb_target) with the source Workbook (wb_source).
' Returns TRUE when successfully finished. The service is denied when the following
' preconditions are not met:
' - CompMan basic configuration is complete and valid
' - the Target-Workbook is located within the configured 'Serviced-Root'
' - the Target-Workbook is the only Workbook in its parent folder
' - CompMan services are not 'Paused'
' - the open/ed Workbook is not a restored version
' - WinMerge is installed (used to display code changes)
' Note: This service is usually called by a developer via the 'Immediate Window'
'       without arguments (may have already been provided by a test procedure).
' --------------------------------------------------------------------------------------------
    Const PROC = "SynchTargetWithSource"
    
    On Error GoTo eh
    Dim sStatus As String
    Dim wbRaw   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Assure complete and correct provision of arguments or get correct ones selected via a dialog
    If Not SyncSourceAndTargetSelected(wb_target:=wb_target _
                                     , cr_raw_name:=wb_source_name _
                                     , wb_source:=wbRaw _
                                      ) Then GoTo xt
    
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    Set mService.Serviced = wb_target
    
    If mService.Denied(PROC) Then GoTo xt
    
    sStatus = Log.Service
        
    SyncVBProjects = mSync.SyncTargetWithSource(wb_target:=wb_target _
                                              , wb_source:=wbRaw _
                                              , restricted_sheet_rename_asserted:=restricted_sheet_rename_asserted _
                                              , design_rows_cols_added_or_deleted:=design_rows_cols_added_or_deleted)

xt: mBasic.EoP ErrSrc(PROC)
    Set Log = Nothing
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsRestoredBySystem() As Boolean
    WbkIsRestoredBySystem = InStr(ActiveWindow.Caption, "(") <> 0 _
                         Or InStr(mService.Serviced.FullName, "(") <> 0
End Function

Public Sub Install(Optional ByRef in_wb As Workbook = Nothing)
    Const PROC = "Install"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If in_wb Is Nothing Then Set in_wb = SelectServicedWrkbk(PROC)
    If in_wb Is Nothing Then GoTo xt
    mInstall.CommonComponents in_wb

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function SelectServicedWrkbk(ByVal gs_service As String) As Workbook
    Dim fl As File
    
    If mFile.SelectFile(sel_filters:="*.xl*" _
                      , sel_filter_name:="VB-Projects/Workbook" _
                      , sel_title:="Select the Workbook (may already be open, will be opened if not) to be served by the " & gs_service & " service" _
                      , sel_result:=fl _
                       ) _
    Then Set SelectServicedWrkbk = mCompMan.WbkGetOpen(fl.Path) _
    Else Set SelectServicedWrkbk = Nothing

End Function

Private Function SyncSourceAndTargetSelected( _
                     Optional ByRef wb_target As Workbook = Nothing, _
                     Optional ByVal cr_raw_name As String = vbNullString, _
                     Optional ByRef wb_source As Workbook = Nothing, _
                     Optional ByVal cr_sync_confirm_info As Boolean = False) As Boolean
' ---------------------------------------------------------------------------
' Returns True when the Sync-Target_VB-Project and the Sync-Source-VB-Project
' are valid. When bc_sync_confirm_info is True a confirmation dialog is
' displayed. The dialog is also displayed when function is called with
' invalid arguments.
' ---------------------------------------------------------------------------
    Const PROC                  As String = "SyncSourceAndTargetSelected"
    Const TARGET_PROJECT        As String = "Target-Workbook/VBProject"
    Const SOURCE_PROJECT        As String = "Source-Workbook/VBProject"
    
    On Error GoTo eh
    Dim sBttCloneRawConfirmed   As String: sBttCloneRawConfirmed = "Selected Source- and" & vbLf & _
                                                                   "Target-Workbook/VBProject" & vbLf & _
                                                                   "Confirmed"
    Dim sBttnTargetProject      As String: sBttnTargetProject = "Select/change the" & vbLf & vbLf & TARGET_PROJECT & vbLf & " "
    Dim sBttnSourceProject      As String: sBttnSourceProject = "Configure/change the" & vbLf & vbLf & SOURCE_PROJECT & vbLf & " "
    Dim sBttnTerminate          As String: sBttnTerminate = "Terminate providing a " & vbLf & _
                                                            "Source- and Target-Workbook/VBProject"
    
    Dim fso         As New FileSystemObject
    Dim sMsg        As TypeMsg
    Dim sReply      As String
    Dim bWbClone    As Boolean
    Dim bWbRaw      As Boolean
    Dim sWbClone    As String
    Dim sWbRaw      As String ' either a full name or a registered raw project's basename
    Dim cllButtons  As Collection
    Dim fl          As File
    
    If Not wb_target Is Nothing Then sWbClone = wb_target.FullName
    sWbRaw = cr_raw_name
    
    While (Not bWbClone Or Not bWbRaw) Or (cr_sync_confirm_info And sReply <> sBttCloneRawConfirmed)
        If sWbClone = vbNullString Then
            sWbClone = "n o t  p r o v i d e d !"
        ElseIf Not fso.FileExists(sWbClone) Then
            sWbRaw = sWbRaw & ": i n v a l i d ! (does not exist)"
        Else
            sWbClone = Split(sWbClone, ": ")(0)
            bWbClone = True
        End If
        
        If sWbRaw = vbNullString Then
            sWbRaw = "n o t  p r o v i d e d !"
        ElseIf Not fso.FileExists(sWbRaw) Then
            sWbRaw = sWbRaw & ": i n v a l i d ! (does not exist)"
        Else
            sWbRaw = Split(sWbRaw, ": ")(0)
            bWbRaw = True
        End If
    
        If bWbRaw And bWbClone And Not cr_sync_confirm_info Then GoTo xt
        
        With sMsg
            .Section(1).Label.Text = TARGET_PROJECT & ":"
            .Section(1).Text.Text = sWbClone
            .Section(1).Text.MonoSpaced = True
            .Section(2).Label.Text = SOURCE_PROJECT & ":"
            .Section(2).Text.Text = sWbRaw
            .Section(2).Text.MonoSpaced = True
            
            If cr_sync_confirm_info _
            Then .Section(3).Text.Text = "Please confirm the above current 'Basic CompMan Configuration'." _
            Else .Section(3).Text.Text = "Please provide/complete the 'Basic CompMan Configuration'."
            
            .Section(3).Text.Text = .Section(3).Text.Text & vbLf & vbLf & _
                                "Attention!" & vbLf & _
                                "1. The '" & TARGET_PROJECT & "' must not be identical with the '" & SOURCE_PROJECT & "' and the two Workbooks must not have the same name." & vbLf & _
                                "2. Both VB-Projects/Workbook must exclusively reside in their parent Workbook" & vbLf & _
                                "3. Both Workbook folders must be subfolders of the configured '" & FOLDER_SERVICED & "'."

        End With
        
        '~~ Buttons preparation
        If Not bWbClone Or Not bWbRaw _
        Then mMsg.Buttons cllButtons, sBttnSourceProject, sBttnTargetProject, vbLf, sBttnTerminate _
        Else mMsg.Buttons cllButtons, sBttCloneRawConfirmed, vbLf, sBttnSourceProject, sBttnTargetProject
        
        sReply = mMsg.Dsply(dsply_title:="Basic configuration of the Component Management (CompMan Addin)" _
                          , dsply_msg:=sMsg _
                          , dsply_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sBttnTargetProject
                Do
                    If mFile.SelectFile(sel_filters:="*.xl*" _
                                      , sel_filter_name:="Workbook/VB-Project" _
                                      , sel_title:="Select the '" & TARGET_PROJECT & " to be synchronized with the '" & SOURCE_PROJECT & "'" _
                                      , sel_result:=fl _
                                       ) Then
                        sWbClone = fl.Path
                        Exit Do
                    End If
                Loop
                cr_sync_confirm_info = True
                '~~ The change of the VB-Clone-Project may have made the VB-Raw-Project valid when formerly invalid
                sWbRaw = Split(sWbRaw, ": ")(0)
            Case sBttnSourceProject
                Do
                    If mFile.SelectFile(sel_filters:="*.xl*" _
                                      , sel_filter_name:="Workbook/VB-Project" _
                                      , sel_title:="Select the '" & SOURCE_PROJECT & " as the synchronization source for the '" & TARGET_PROJECT & "'" _
                                      , sel_result:=fl _
                                       ) Then
                        sWbRaw = fl.Path
                        Exit Do
                    End If
                Loop
                cr_sync_confirm_info = True
                '~~ The change of the VB-Raw-Project may have become valid when formerly invalid
                sWbClone = Split(sWbClone, ": ")(0)
            
            Case sBttCloneRawConfirmed: cr_sync_confirm_info = False
            Case sBttnTerminate: GoTo xt
                
        End Select
        
    Wend ' Loop until the confirmed or configured basic configuration is correct
    
xt: If bWbClone Then
       Set wb_target = mCompMan.WbkGetOpen(sWbClone)
    End If
    If bWbRaw Then
        Application.EnableEvents = False
        Set wb_source = mCompMan.WbkGetOpen(sWbRaw)
        Application.EnableEvents = True
        cr_raw_name = fso.GetBaseName(sWbRaw)
    End If
    SyncSourceAndTargetSelected = bWbClone And bWbRaw
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function UsedCommonComponents(ByRef cl_wb As Workbook) As Dictionary
' ---------------------------------------------------------------------------
' Returns a Dictionary of all Used Common Components with its VBComponent
' object as key and its name as item.
' ---------------------------------------------------------------------------
    Const PROC = "UsedCommonComponents"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim dct     As New Dictionary
    Dim fso     As New FileSystemObject
    Dim Comp    As clsComp
    Dim RawComp As clsRaw
    
    mBasic.BoP ErrSrc(PROC)
        
    For Each vbc In cl_wb.VBProject.VBComponents
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = cl_wb
            .CompName = vbc.name
            Log.ServicedItem = .VBComp
            If .KindOfComp = enCommCompUsed Then
                If .Changed Then
                    dct.Add vbc, vbc.name
                Else
                    Log.Entry = "Code un-changed."
                End If
            End If
        End With
        Set Comp = Nothing
        Set RawComp = Nothing
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)
    Set UsedCommonComponents = dct
    Set fso = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FilesDiffer(ByVal fd_exp_file_1 As File, _
                            ByVal fd_exp_file_2 As File) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when file 1 and file 2 are different whereby case differences
' and empty lines are ignored. This function guarantees a uniform comparison of
' export files throughout CompMan.
' ----------------------------------------------------------------------------
    FilesDiffer = mFile.Differs(fd_file1:=fd_exp_file_1 _
                              , fd_file2:=fd_exp_file_2 _
                              , fd_stop_after:=1 _
                              , fd_ignore_empty_records:=True _
                              , fd_compare:=vbTextCompare).Count <> 0
                            
End Function

Public Function FilesDifference(ByVal fd_exp_file_1 As File, _
                                ByVal fd_exp_file_2 As File) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with either 0 items when file 1 and file 2 are
' identical or with one item when the two files differ. Empty lines and case
' differences are ignored because the do not constitute a relevant code change
' ----------------------------------------------------------------------------
    Set FilesDifference = mFile.Differs(fd_file1:=fd_exp_file_1 _
                                      , fd_file2:=fd_exp_file_2 _
                                      , fd_stop_after:=1 _
                                      , fd_ignore_empty_records:=True _
                                      , fd_compare:=vbTextCompare)
                            
End Function

Public Sub RemoveTempRenamed()
' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------
    Const PROC = "RemoveTempRenamed"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim vbc As VBComponent
    
    With mService.Serviced.VBProject
        For Each v In .VBComponents
            Set vbc = v
            If mService.IsRenamedByCompMan(vbc.name) Then
                .VBComponents.Remove vbc
            End If
        Next v
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function IsRenamedByCompMan(ByVal comp_name As String) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the component's name indicates that it is one which had been
' renamed by CompMan for an update (rename/import) service.
' ------------------------------------------------------------------------------
    IsRenamedByCompMan = InStr(comp_name, RENAMED_BY_COMPMAN) <> 0
End Function

Public Function Progress(ByVal p_service As String, _
                Optional ByVal p_result As Long = 0, _
                Optional ByVal p_of As Long = 0, _
                Optional ByVal p_op As String = vbNullString, _
                Optional ByVal p_comps As String = vbNullString, _
                Optional ByVal p_dots As Long = 0) As String
' --------------------------------------------------------------------------
' Displays a services progress in the Application.StatusBar in the form:
' service for serviced: n of m op [(component [, component] ..]
' --------------------------------------------------------------------------
    Const MSG_SCHEEME = "<s><n> of <m> <op> <comps> <dots>"
    
    Dim sMsg        As String
    
    sMsg = Replace(MSG_SCHEEME, "<s>", p_service)
    sMsg = Replace(sMsg, "<n>", p_result)
    sMsg = Replace(sMsg, "<m>", p_of)
    sMsg = Replace(sMsg, "<op>", p_op)
    If p_comps <> vbNullString Then
        If Right(p_comps, 2) = ", " Then p_comps = Left(p_comps, Len(p_comps) - 2)
        If Right(p_comps, 1) = "," Then p_comps = Left(p_comps, Len(p_comps) - 1)
        sMsg = Replace(sMsg, "<comps>", "(" & p_comps & ")")
    Else
        sMsg = Replace(sMsg, "<comps>", vbNullString)
    End If
    If p_dots > 0 Then
        sMsg = Replace(sMsg, "<dots>", VBA.String(p_dots, "."))
    Else
        sMsg = Replace(sMsg, "<dots>", vbNullString)
    End If
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 250) & " ..."
        
    Progress = sMsg
    
End Function

Public Function ExpFilesDiffDisplay( _
          ByVal fd_exp_file_left_full_name As String, _
          ByVal fd_exp_file_right_full_name As String, _
          ByVal fd_exp_file_left_title As String, _
          ByVal fd_exp_file_right_title As String) As Boolean
' ----------------------------------------------------------------------------
' Displays the differences between export file 1 and 2 by means of WinMerge!
' Note: CompMan ignores any differences caused by empty code lines or case.
'       When a difference is displayed it is thus not because of this kind of
'       differneces but of others. Unfortunately it depends on the installed
'       WinMerge's set option wether or not these kind of differences are
'       displayed.
' ----------------------------------------------------------------------------
    Const PROC                          As String = "ExpFilesDiffDisplay"
    Const WINMERGE_SETTINGS_BASE_KEY    As String = "HKEY_CURRENT_USER\SOFTWARE\Thingamahoochie\WinMerge\Settings\"
    Const WINMERGE_BLANK_LINES          As String = "IgnoreBlankLines"
    Const WINMERGE_IGNORE_CASE          As String = "IgnoreCase"
    
    On Error GoTo eh
    Dim waitOnReturn        As Boolean: waitOnReturn = True
    Dim windowStyle         As Integer: windowStyle = 1
    Dim sCommand            As String
    Dim fso                 As New FileSystemObject
    Dim wshShell            As Object
    Dim sIniFile            As String
    Dim sIgnoreBlankLines   As String ' 1 = True, 0 = False
    Dim sIgnoreCase         As String ' 1 = True, 0 = False
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , source:=ErrSrc(PROC) _
                 , Description:="WinMerge is obligatory for the Compare service of this module but not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(fd_exp_file_left_full_name) _
    Then Err.Raise Number:=AppErr(2) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & fd_exp_file_left_full_name & """ does not exist!"
    
    If Not fso.FileExists(fd_exp_file_right_full_name) _
    Then Err.Raise Number:=AppErr(3) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & fd_exp_file_right_full_name & """ does not exist!"
    
    '~~ Save WinMerge configuration items and set them for CompMan
    sIgnoreBlankLines = mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_BLANK_LINES)
    sIgnoreCase = mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_IGNORE_CASE)
    mReg.Value(WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_BLANK_LINES) = "1"
    mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_IGNORE_CASE) = "1"
    
    '~~ Prepare command line
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & fd_exp_file_left_title & DQUOTE & _
               " /dr " & DQUOTE & fd_exp_file_right_title & DQUOTE & " " & _
               """" & fd_exp_file_left_full_name & """" & " " & _
               """" & fd_exp_file_right_full_name & """" ' & sIniFile doesn't work

    
    '~~ Execute command line
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        ExpFilesDiffDisplay = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
    '~~ Restore WinMerge configuration items
    mReg.Value(WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_BLANK_LINES) = sIgnoreBlankLines
    mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_IGNORE_CASE) = sIgnoreCase

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function


