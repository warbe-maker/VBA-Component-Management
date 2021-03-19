Attribute VB_Name = "mService"
Option Explicit

Public Const SERVICES_LOG_FILE = "CompMan.Services.log"
Public cLog         As New clsLog
Public cStats       As New clsStats
Private wbServiced  As Workbook
Public dctServiced  As Dictionary

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
    Dim cComp       As New clsComp
    Dim flFile      As File
    Dim wbTemp      As Workbook
    Dim wbActive    As Workbook
    Dim sBaseName   As String
    
    If rc_wb Is Nothing Then Set rc_wb = ActiveWorkbook
    cComp.Wrkbk = rc_wb
    If rc_exp_file_full_name <> vbNullString Then
        If Not fso.FileExists(rc_exp_file_full_name) Then
            rc_exp_file_full_name = vbNullString ' enforces selection when the component name is also not provided
        End If
    End If
    
    If rc_comp_name <> vbNullString Then
        cComp.CompName = rc_comp_name
        If Not cComp.Exists Then
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
        If mFile.SelectFile(sel_init_path:=cComp.ExpFilePath _
                          , sel_filters:="*.bas,*.cls,*.frm" _
                          , sel_filter_name:="File" _
                          , sel_title:="Select the Export-File for the re-new service" _
                          , sel_result:=flFile) _
        Then rc_exp_file_full_name = flFile.Path
    End If
    
    If rc_comp_name <> vbNullString _
    And rc_exp_file_full_name = vbNullString Then
        cComp.CompName = rc_comp_name
        '~~ ------------------------------------------------
        '~~ Select the component's corresponding Export-File
        '~~ ------------------------------------------------
        sBaseName = fso.GetBaseName(rc_exp_file_full_name)
        '~~ Select the Export-File for the re-new service
        If mFile.SelectFile(sel_init_path:=cComp.ExpFilePath _
                          , sel_filters:="*" & cComp.ExpFileExtension _
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
    
    With cComp
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
            cLog.Entry = "Active Workbook de-activated by creating a temporary Workbook"
        End If
            
        mRenew.ByImport rn_wb:=.Wrkbk _
             , rn_comp_name:=.CompName _
             , rn_exp_file_full_name:=rc_exp_file_full_name
        cLog.ServicedItem(.TypeString) = .CompName
        cLog.Entry = "Component renewed/updated by (re-)import of '" & rc_exp_file_full_name & "'"
    End With
    
xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        cLog.Entry = "Temporary created Workbook closed without save"
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            cLog.Entry = "De-activated Workbook '" & wbActive.Name & "' re-activated"
            Set wbActive = Nothing
        End If
    End If
    Set cComp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Continue()
' -------------------------------------------
' Continues the paused CompMan Addin Services
' -------------------------------------------
    mMe.AddInPaused = False
    mMe.DisplayStatus
End Sub

Public Sub Pause()
' ----------------------------------
' Pauses the CompMan Addin Services
' ---------------------------------
    mMe.AddInPaused = True
    mMe.DisplayStatus
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
        CompMaxLen = mBasic.Max(CompMaxLen, Len(vbc.Name))
    Next vbc
End Function

Public Function Denied(ByRef den_serviced_wb As Workbook, _
                       ByVal den_service As String, _
              Optional ByVal den_new_log As Boolean = False) As Boolean
' --------------------------------------------------------------------------
' Returns True when all preconditions for a service execution are fulfilled.
' --------------------------------------------------------------------------
    Dim sStatus As String
    
    Set cLog.ServicedWrkbk(sw_new_log:=den_new_log) = den_serviced_wb
    cLog.Service = den_service

    If Not mMe.BasicConfig Then
        sStatus = "Service denied! The Basic CompMan Configuration is invalid!"
        cLog.Entry = sStatus
        cLog.Entry = "The assertion of a valid basic configuration has been terminated though invalid!"
        Denied = True
    ElseIf WbkIsRestoredBySystem(den_serviced_wb) Then
        sStatus = "Service denied! Workbook appears restored by the system!"
        cLog.Entry = sStatus
        Denied = True
    ElseIf Not WbkInServicedRoot(den_serviced_wb) Then
        sStatus = "Service denied! Workbook is not within the configured 'serviced root': " & mMe.ServicedRoot & "!"
        cLog.Entry = sStatus
        Denied = True
    ElseIf mMe.AddInPaused Then
        sStatus = "Service denied! The CompMan Addin is currently paused!"
        cLog.Entry = sStatus
        Denied = True
    ElseIf FolderNotVbProjectExclusive(den_serviced_wb) Then
        sStatus = "Service denied! The Workbook is not the only one in its parent folder!"
        cLog.Entry = sStatus
        Denied = True
    End If
    If Denied _
    Then Application.StatusBar = cLog.Service & sStatus

End Function

Private Function ErrSrc(ByVal S As String) As String
    ErrSrc = "mService." & S
End Function

Public Sub ExportChangedComponents( _
                             ByRef ec_wb As Workbook, _
                    Optional ByVal ec_hosted As String = vbNullString)
' --------------------------------------------------------------------
' Exclusively performed/trigered by the Before_Save event:
' - Any code change (detected by the comparison of a temporary export
'   file with the current export file) is backed-up/exported
' - Outdated Export Files (components no longer existing) are removed
' - Clone code modifications update the raw code when confirmed by the
'   user
' --------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    Dim vbc                 As VBComponent
    Dim lComponents         As Long
    Dim lCompsRemaining     As Long
    Dim lExported           As Long
    Dim sExported           As String
    Dim bUpdated            As Boolean
    Dim lUpdated            As Long
    Dim sUpdated            As String
    Dim sMsg                As String
    Dim fso                 As New FileSystemObject
    Dim sProgressDots       As String
    Dim sStatus             As String
    Dim v                   As Variant
    
    mErH.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    Set wbServiced = ec_wb
    If mService.Denied(den_serviced_wb:=ec_wb, den_service:=PROC) Then GoTo xt
    
    sStatus = cLog.Service
    Set Stats = New clsStats
    CollectServicedComponents
    
    mCompMan.DeleteObsoleteExpFiles do_wb:=ec_wb
    If Not mCompMan.ManageVbProjectProperties(mh_hosted:=ec_hosted _
                                            , mh_wb:=ec_wb) _
    Then
        Application.StatusBar = cLog.Service & " failed. See log-file in Workbookfolder"
        GoTo xt
    End If
    
    lComponents = ec_wb.VBProject.VBComponents.Count
    lCompsRemaining = lComponents
    sProgressDots = String$(lCompsRemaining, ".")

    For Each v In dctServiced
        Set vbc = dctServiced(v)
        Set cComp = New clsComp
        sProgressDots = Left(sProgressDots, Len(sProgressDots) - 1)
        Application.StatusBar = sStatus & sProgressDots & sExported & " " & vbc.Name
        mTrc.BoC ErrSrc(PROC) & " " & vbc.Name
        Set cComp = New clsComp
        With cComp
            Set .Wrkbk = ec_wb
            .CompName = vbc.Name
            cLog.ServicedItem(.TypeString) = .CompName
            If CodeModuleIsEmpty(.VBComp) Then
                '~~ Empty Code Modules are exported only when the Workbook is a VB-Raw-Project
                If mRawHosts.Exists(.WrkbkBaseName) Then
                    If Not mRawHosts.IsRawVbProject(.WrkbkBaseName) Then GoTo next_vbc
                End If
            End If
            If Not fso.FileExists(.ExpFileFullName) Then
                .VBComp.Export .ExpFileFullName
                cLog.Entry = "Initially exported to '" & .ExpFileFullName & "'."
            End If
            If Not .Changed Then
                cLog.Entry = "Code un-changed! (temporary Export-File identical with last change Export-File)"
                GoTo next_vbc
            End If
        End With
                
        Select Case cComp.KindOfComp
            Case enRawClone
                '~~ Establish a component class object which represents the cloned raw's remote instance
                '~~ which is hosted in another Workbook
                Set cRaw = New clsRaw
                With cRaw
                    '~~ Provide all available information rearding the remote raw component
                    '~~ Attention must be paid to the fact that the sequence of property assignments matters
                    .HostFullName = mHostedRaws.HostFullName(comp_name:=cComp.CompName)
                    .CompName = cComp.CompName
                    .ExpFileExtension = cComp.ExpFileExtension  ' required to build the export file's full name
                    Set .ExpFile = fso.GetFile(.ExpFileFullName)
                    .CloneExpFileFullName = cComp.ExpFileFullName
                    .TypeString = cComp.TypeString
                    If Not .Changed And Not cComp.Changed Then GoTo next_vbc
                End With
                
                With cComp
                    If .Changed And Not cRaw.Changed Then
                        cLog.Entry = "The Clone's code changed! (a temporary Export-File differs from the last regular Export-File)"
                        '~~ --------------------------------------------------------------------------
                        '~~ The code change in the clone component is now in question whether it is to
                        '~~ be ignored, i.e. the change is reverted with the Workbook's next open or
                        '~~ the raw code should be updated accordingly to make the change permanent
                        '~~ for all users of the component.
                        '~~ --------------------------------------------------------------------------
                        .VBComp.Export .ExpFileFullName
                        '~~ In case the raw had been imported manually the new check for a change will indicate no change
                        If cRaw.Changed(check_again:=True) Then GoTo next_vbc
                        .ReplaceRawWithCloneWhenConfirmed rwu_updated:=bUpdated ' when confirmed in user dialog
                        If bUpdated Then
                            lUpdated = lUpdated + 1
                            sUpdated = vbc.Name & ", " & sUpdated
                            cLog.Entry = """Remote Raw"" has been updated with code of ""Raw Clone"""
                        End If
                        
                    ElseIf Not .Changed And cRaw.Changed Then
                        '~~ -----------------------------------------------------------------------
                        '~~ The raw had changed since the Workbook's open. This case is not handled
                        '~~ along with the Workbook's Save event but with the Workbook's Open event
                        '~~ -----------------------------------------------------------------------
                        cLog.Entry = "The Raw's code changed! (not considered with the export service)"
                        cLog.Entry = "The Clone will be updated with the next Workbook open"
                    End If
                End With
            
            Case enKindOfComp.enUnknown
                '~~ This should never be the case in is thus ignored
            
            Case Else ' enInternal, enHostedRaw
                With cComp
                    If .Changed Then
                        cLog.Entry = "Code changed! (temporary Export-File differs from last changes Export-File)"
                        Application.StatusBar = sStatus & vbc.Name & " Export to '" & .ExpFileFullName & "'"
                        vbc.Export .ExpFileFullName
                        sStatus = sStatus & vbc.Name & ", "
                        cLog.Entry = "Exported to '" & .ExpFileFullName & "'"
                        lExported = lExported + 1
                        If lExported = 1 _
                        Then sExported = vbc.Name _
                        Else sExported = sExported & ", " & vbc.Name
                        GoTo next_vbc
                    End If
                
                    If .KindOfComp = enHostedRaw Then
                        If mHostedRaws.ExpFileFullName(comp_name:=.CompName) <> .ExpFileFullName Then
                            mHostedRaws.ExpFileFullName(comp_name:=.CompName) = .ExpFileFullName
                            cLog.Entry = "Component's Export-File Full Name registered"
                        End If
                    End If
                End With
        End Select
                                
next_vbc:
        mTrc.EoC ErrSrc(PROC) & " " & vbc.Name
        lCompsRemaining = lCompsRemaining - 1
        Set cComp = Nothing
        Set cRaw = Nothing
    Next v
    
    sMsg = cLog.Service
    Select Case lExported
        Case 0:     sMsg = sMsg & "None of the " & lComponents & " components' code has changed."
        Case Else:  sMsg = sMsg & lExported & " of " & lComponents & " changed components exported: " & sExported
    End Select
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: Set dctHostedRaws = Nothing
    Set cComp = Nothing
    Set cRaw = Nothing
    Set cLog = Nothing
    Set fso = Nothing
    mErH.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function FolderNotVbProjectExclusive(ByRef wb As Workbook) As Boolean

    Dim fso As New FileSystemObject
    Dim fl  As File
    
    For Each fl In fso.GetFolder(wb.Path).Files
        If VBA.Left$(fso.GetFileName(fl.Path), 2) = "~$" Then GoTo next_fl
        If VBA.StrComp(fl.Path, wb.FullName, vbTextCompare) <> 0 Then
            Select Case fso.GetExtensionName(fl.Path)
                Case "xlsm", "xlam", "xlsb" ' may contain macros, a VB-Project repsectively
                    FolderNotVbProjectExclusive = True
                    Exit For
            End Select
        End If
next_fl:
    Next fl

End Function

Public Sub ExportAll(Optional ByRef ea_wb As Workbook = Nothing)
' --------------------------------------------------------------
'
' --------------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim sStatus As String
    
    mErH.BoP ErrSrc(PROC)
    If ea_wb Is Nothing Then Set ea_wb = ServicedWrkbk("Se")
    If ea_wb Is Nothing Then GoTo xt
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied(den_serviced_wb:=ea_wb, den_service:=PROC) Then GoTo xt
    
    sStatus = cLog.Service

    mCompMan.DeleteObsoleteExpFiles ea_wb
    
    If mMe.IsAddinInstnc _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
    
    For Each vbc In ea_wb.VBProject.VBComponents
        Set cComp = New clsComp
        With cComp
            Set .Wrkbk = ea_wb
            .CompName = vbc.Name ' this assignment provides the name for the export file
            vbc.Export .ExpFileFullName
        End With
        Set cComp = Nothing
    Next vbc

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Function SyncTargetWithSource( _
                      Optional ByRef wb_target As Workbook = Nothing, _
                      Optional ByVal wb_source_name As String = vbNullString, _
                      Optional ByRef bkp_folder As String) As Boolean
' -----------------------------------------------------------------------------
' Synchronizes the target Workbook (wb_target) with the source Workbook
' (wb_source). Returns TRUE when successfully finished.
' The service is performed provided:
' - the CompMan basic configuration is complete and valid
' - the Workbook is located within the configured "Serviced Root"
' - the Workbook is the only one within its parent folder
' - the CompMan Addin is not paused
' - the open/ed Workbook is not a restored version
' -----------------------------------------------------------------------------
    Const PROC = "SynchTargetWithSource"
    
    On Error GoTo eh
    Dim sStatus As String
    Dim wbRaw   As Workbook
    
    mErH.BoP ErrSrc(PROC)
    '~~ Assure complete and correct provision of arguments or get correct ones selected via a dialog
    If Not CloneAndRawProject(wb_target:=wb_target _
                            , cr_raw_name:=wb_source_name _
                            , wb_source:=wbRaw _
                             ) Then GoTo xt
    
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    If mService.Denied(den_serviced_wb:=wb_target, den_service:=PROC) Then GoTo xt

    Set cLog.ServicedWrkbk = wb_target
    sStatus = cLog.Service
        
    SyncTargetWithSource = mSync.SyncTargetWithSource(wb_target:=wb_target _
                                                    , wb_source:=wbRaw _
                                                    , bkp_folder:=bkp_folder)

xt: mErH.EoP ErrSrc(PROC)
    Set cLog = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Sub UpdateRawClones( _
                     ByRef uc_wb As Workbook, _
            Optional ByVal uc_hosted As String = vbNullString)
' ------------------------------------------------------------
' Updates all clone components with the Export-File of the
' raw component provided the raw's code has changed.
' ------------------------------------------------------------
    Const PROC = "UpdateRawClones"
    
    On Error GoTo eh
    Dim wbActive    As Workbook
    Dim wbTemp      As Workbook
    Dim sStatus     As String
    Dim lMaxLenComp As Long
    
    mErH.BoP ErrSrc(PROC)
    If mService.Denied(den_serviced_wb:=uc_wb, den_service:=PROC, den_new_log:=True) Then GoTo xt

    Application.StatusBar = sStatus & "Maintain hosted raws"
    If Not mCompMan.ManageVbProjectProperties(mh_hosted:=uc_hosted _
                                            , mh_wb:=uc_wb) _
    Then
        Application.StatusBar = cLog.Service & " failed. See log-file in Workbookfolder"
        GoTo xt
    End If
        
    Application.StatusBar = sStatus & "De-activate '" & uc_wb.Name & "'"
    If uc_wb Is ActiveWorkbook Then
        '~~ De-activate the ActiveWorkbook by creating a temporary Workbook
        Set wbActive = uc_wb
        Set wbTemp = Workbooks.Add
    End If
    
    mUpdate.RawClones urc_wb:=uc_wb _
                    , urc_comp_max_len:=lMaxLenComp _
                    , urc_clones:=Clones(uc_wb)

xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            Set wbActive = Nothing
        End If
    End If
    Set dctHostedRaws = Nothing
    Set cLog = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function WbkInServicedRoot(ByRef idr_wb As Workbook) As Boolean
    WbkInServicedRoot = InStr(idr_wb.Path, mMe.ServicedRoot) <> 0
End Function

Private Function WbkIsRestoredBySystem(ByRef rbs_wb As Workbook) As Boolean
    WbkIsRestoredBySystem = InStr(ActiveWindow.Caption, "(") <> 0 _
                         Or InStr(rbs_wb.FullName, "(") <> 0
End Function

Public Sub Install(Optional ByRef in_wb As Workbook = Nothing)
    Const PROC = "Install"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    If in_wb Is Nothing Then Set in_wb = ServicedWrkbk(PROC)
    If in_wb Is Nothing Then GoTo xt
    mInstall.CloneRaws in_wb

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

Private Function ServicedWrkbk(ByVal gs_service As String) As Workbook
    Dim fl As File
    
    If mFile.SelectFile(sel_filters:="*.xl*" _
                      , sel_filter_name:="VB-Projects/Workbook" _
                      , sel_title:="Select the Workbook (may already be open, will be opened if not) to be served by the " & gs_service & " service" _
                      , sel_result:=fl _
                       ) _
    Then Set ServicedWrkbk = mCompMan.WbkGetOpen(fl.Path) _
    Else Set ServicedWrkbk = Nothing

End Function

Private Function CloneAndRawProject( _
                     Optional ByRef wb_target As Workbook = Nothing, _
                     Optional ByVal cr_raw_name As String = vbNullString, _
                     Optional ByRef wb_source As Workbook = Nothing, _
                     Optional ByVal cr_sync_confirm_info As Boolean = False) As Boolean
' ---------------------------------------------------------------------------
' Returns True when the Sync-Target_VB-Project and the Sync-Source-VB-Project are valid.
' When bc_sync_confirm_info is True a confirmation dialog is displayed. The dialog is
' also displayed when function is called with invalid arguments.
' ---------------------------------------------------------------------------
    Const PROC                  As String = "CloneAndRawProject"
    Const CLONE_PROJECT         As String = "VB-Clone-Project"
    Const RAW_PROJECT           As String = "VB-Raw-Project"
    
    On Error GoTo eh
    Dim sBttCloneRawConfirmed   As String: sBttCloneRawConfirmed = "VB-Clone- and VB-Raw-Project" & vbLf & "Confirmed"
    Dim sBttnCloneProject       As String: sBttnCloneProject = "Select/change the" & vbLf & vbLf & CLONE_PROJECT & vbLf & " "
    Dim sBttnRawProject         As String: sBttnRawProject = "Configure/change the" & vbLf & vbLf & RAW_PROJECT & vbLf & " "
    Dim sBttnTerminate          As String: sBttnTerminate = "Terminate providing a " & vbLf & _
                                                            "VB-Clone- and a VB-Raw-Project" & vbLf & _
                                                            "for being synchronized" & vbLf & _
                                                            "(sync service will be denied)"
    
    Dim fso         As New FileSystemObject
    Dim sMsg        As tMsg
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
            .Section(1).sLabel = CLONE_PROJECT & ":"
            .Section(1).sText = sWbClone
            .Section(1).bMonspaced = True
            .Section(2).sLabel = RAW_PROJECT & ":"
            .Section(2).sText = sWbRaw
            .Section(2).bMonspaced = True
            
            If cr_sync_confirm_info _
            Then .Section(3).sText = "Please sync_confirm_info the provided VB-Clone- and VB-Raw-Project." _
            Else .Section(3).sText = "Please provide/complete the VB-Clone- (sync target) and the VB-Raw-Project (sync source)."
            
            .Section(3).sText = .Section(3).sText & vbLf & vbLf & _
                                "Attention!" & vbLf & _
                                "1. The '" & CLONE_PROJECT & "' must not be identical with the '" & RAW_PROJECT & "' and the two Workbooks must not have the same name." & vbLf & _
                                "2. Both VB-Projects/Workbook must exclusively reside in their parent Workbook" & vbLf & _
                                "3. Both Workbook folders must be subfolders of the configured '" & FOLDER_SERVICED & "'."

        End With
        
        '~~ Buttons preparation
        If Not bWbClone Or Not bWbRaw _
        Then Set cllButtons = mMsg.Buttons(sBttnRawProject, sBttnCloneProject, vbLf, sBttnTerminate) _
        Else Set cllButtons = mMsg.Buttons(sBttCloneRawConfirmed, vbLf, sBttnRawProject, sBttnCloneProject)
        
        sReply = mMsg.Dsply(msg_title:="Basic configuration of the Component Management (CompMan Addin)" _
                          , msg:=sMsg _
                          , msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sBttnCloneProject
                Do
                    If mFile.SelectFile(sel_filters:="*.xl*" _
                                      , sel_filter_name:="Workbook/VB-Project" _
                                      , sel_title:="Select the '" & CLONE_PROJECT & " to be synchronized with the '" & RAW_PROJECT & "'" _
                                      , sel_result:=fl _
                                       ) Then
                        sWbClone = fl.Path
                        Exit Do
                    End If
                Loop
                cr_sync_confirm_info = True
                '~~ The change of the VB-Clone-Project may have made the VB-Raw-Project valid when formerly invalid
                sWbRaw = Split(sWbRaw, ": ")(0)
            Case sBttnRawProject
                Do
                    If mFile.SelectFile(sel_filters:="*.xl*" _
                                      , sel_filter_name:="Workbook/VB-Project" _
                                      , sel_title:="Select the '" & RAW_PROJECT & " as the synchronization source for the '" & CLONE_PROJECT & "'" _
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
    CloneAndRawProject = bWbClone And bWbRaw
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Function Clones( _
                  ByRef cl_wb As Workbook) As Dictionary
' ------------------------------------------------------
' Returns a Dictionary with clone component's object as
' the key and their kind of code change as item.
' ------------------------------------------------------
    Const PROC = "Clones"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim dct         As New Dictionary
    Dim fso         As New FileSystemObject
    
    mErH.BoP ErrSrc(PROC)
        
    For Each vbc In cl_wb.VBProject.VBComponents
        Set cComp = New clsComp
        With cComp
            Set .Wrkbk = cl_wb
            .CompName = vbc.Name
            cLog.ServicedItem(.TypeString) = .CompName
            If .KindOfComp = enRawClone Then
                Set cRaw = New clsRaw
                cRaw.HostFullName = mHostedRaws.HostFullName(comp_name:=.CompName)
                cRaw.CompName = .CompName
                cRaw.ExpFileExtension = .ExpFileExtension
                cRaw.CloneExpFileFullName = .ExpFileFullName
                cRaw.TypeString = .TypeString
                If .Changed Then
                    dct.Add vbc, vbc.Name
                Else
                    cLog.Entry = "Code un-changed."
                End If
                If cRaw.Changed Then
                    If Not dct.Exists(vbc) Then dct.Add vbc, vbc.Name
                Else
                    cLog.Entry = "Corresponding Raw's code un-changed."
                End If
            End If
        End With
        Set cComp = Nothing
        Set cRaw = Nothing
    Next vbc

xt: mErH.EoP ErrSrc(PROC)
    Set Clones = dct
    Set fso = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub CollectServicedComponents()
    Const PROC = "CollectServicedComponents"
    
    On Error GoTo eh
    Dim ws1     As Worksheet
    Dim ws2     As Worksheet
    Dim vbc     As VBComponent
    Dim shp     As Shape
    Dim nm      As Name
    Dim nmName  As Name
    Dim sKey    As String
    
    If dctServiced Is Nothing Then Set dctServiced = New Dictionary Else dctServiced.RemoveAll
    
    Stats.Count sic_vbcomps_total, wbServiced.VBProject.VBComponents.Count
    For Each vbc In wbServiced.VBProject.VBComponents
        With vbc
            mDct.DctAdd dctServiced, vbc.Type & ":" & vbc.Name, vbc, order_bykey, seq_ascending, , , True
        End With
    Next vbc
     
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

