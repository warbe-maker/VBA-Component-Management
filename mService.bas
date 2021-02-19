Attribute VB_Name = "mService"
Option Explicit

Public Const SERVICES_LOG_FILE = "CompMan.Services.log"
Public cLog             As New clsLog

Public Sub Continue()
' -------------------------------------------
' Continues the paused CompMan Addin Services
' -------------------------------------------
    mMe.AddInPaused = False
    wsAddIn.CompManAddInStatus = _
        "The AddIn is currently  a c t i v e ! The Workbook_Open service 'UpdateRawClones' " & _
        "and the Workbook_BeforeSave service 'ExportChangedComponents' will be executed for " & _
        "Workbooks calling them under the following met preconditions: " & vbLf & _
        "1. The CompMan's basic configuration is complete and valid" & vbLf & _
        "2. The Workbook is located in the configured 'Serviced Development Root' which currently is:" & vbLf & _
        "   '" & mMe.ServicedRoot & "'" & vbLf & _
        "3. The Workbook is the only one within its parent folder" & vbLf & _
        "4. The Workbook is not a restored version" & vbLf & _
        "5. The serviced VB-Project has a Conditional Compile Argument 'CompMan = 1'"
End Sub

Public Sub Pause()
' ----------------------------------
' Pauses the CompMan Addin Services
' ---------------------------------
    mMe.AddInPaused = True
    wsAddIn.CompManAddInStatus = _
        "The AddIn is currently  p a u s e d ! The Workbook_Open service 'UpdateRawClones' " & _
        "and the Workbook_BeforeSave service 'ExportChangedComponents' will be bypassed " & _
        "until the Addin is 'continued' again!"
End Sub

Private Function CodeModuleIsEmpty(ByVal vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0
        If Not CodeModuleIsEmpty Then CodeModuleIsEmpty = .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Public Function CompMaxLen(ByVal ml_wb As Workbook) As Long
    Dim vbc As VBComponent
    For Each vbc In ml_wb.VbProject.VBComponents
        CompMaxLen = mBasic.Max(CompMaxLen, Len(vbc.name))
    Next vbc
End Function

Public Function Denied(ByVal den_serviced_wb As Workbook, _
                       ByVal den_service As String, _
              Optional ByVal den_new_log As Boolean = False) As Boolean
' --------------------------------------------------------------------------
' Returns True when all preconditions for a service execution are fulfilled.
' --------------------------------------------------------------------------
    Dim sStatus As String
    
    cLog.ServicedWrkbk(sw_new_log:=den_new_log) = den_serviced_wb
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
        sStatus = "Service denied! The Workbook is not an exclusive in its parent folder!"
        cLog.Entry = sStatus
        Denied = True
    End If
    If Denied _
    Then Application.StatusBar = cLog.Service & sStatus

End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mService." & s
End Function

Public Sub ExportChangedComponents( _
                             ByVal ec_wb As Workbook, _
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
    
    mErH.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    If mService.Denied(den_serviced_wb:=ec_wb, den_service:=ErrSrc(PROC)) Then GoTo xt
    
    mCompMan.Service = PROC & " for '" & ec_wb.name & "': "
    sStatus = mCompMan.Service

    mCompMan.DeleteObsoleteExpFiles do_wb:=ec_wb
    If Not mCompMan.ManageVbProjectProperties(mh_hosted:=ec_hosted _
                                            , mh_wb:=ec_wb) _
    Then
        Application.StatusBar = cLog.Service & " failed. See log-file in Workbookfolder"
        GoTo xt
    End If
    
    lComponents = ec_wb.VbProject.VBComponents.Count
    lCompsRemaining = lComponents
    sProgressDots = String$(lCompsRemaining, ".")

    For Each vbc In ec_wb.VbProject.VBComponents
        Set cComp = New clsComp
        sProgressDots = Left(sProgressDots, Len(sProgressDots) - 1)
        Application.StatusBar = sStatus & sProgressDots & sExported & " " & vbc.name
        mTrc.BoC ErrSrc(PROC) & " " & vbc.name
        Set cComp = New clsComp
        With cComp
            .Wrkbk = ec_wb
            .CompName = vbc.name
            cLog.ServicedItem = .CompName
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
            If Not .Changed Then GoTo next_vbc
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
                    .ExpFile = fso.GetFile(.ExpFileFullName)
                    .CloneExpFileFullName = cComp.ExpFileFullName
                    If Not .Changed And Not cComp.Changed Then GoTo next_vbc
                End With
                
                With cComp
                    If .Changed And Not cRaw.Changed Then
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
                            sUpdated = vbc.name & ", " & sUpdated
                            cLog.Entry = """Remote Raw"" has been updated with code of ""Raw Clone"""
                        End If
                        
                    ElseIf Not .Changed And cRaw.Changed Then
                        '~~ -----------------------------------------------------------------------
                        '~~ The raw had changed since the Workbook's open. This case is not handled
                        '~~ along with the Workbook's Save event but with the Workbook's Open event
                        '~~ -----------------------------------------------------------------------
                    End If
                End With
            
            Case enKindOfComp.enUnknown
                '~~ This should never be the case in is thus ignored
            
            Case Else ' enInternal, enHostedRaw
                With cComp
                    If .Changed Then
                        Application.StatusBar = sStatus & vbc.name & " Export to '" & .ExpFileFullName & "'"
                        vbc.Export .ExpFileFullName
                        sStatus = sStatus & vbc.name & ", "
                        cLog.Entry = "Changes exported to '" & .ExpFileFullName & "'"
                        lExported = lExported + 1
                        If lExported = 1 _
                        Then sExported = vbc.name _
                        Else sExported = sExported & ", " & vbc.name
                        GoTo next_vbc
                    End If
                
                    If .KindOfComp = enHostedRaw Then
                        If mHostedRaws.ExpFileFullName(comp_name:=.CompName) <> .ExpFileFullName Then
                            mHostedRaws.ExpFileFullName(comp_name:=.CompName) = .ExpFileFullName
                            cLog.Entry = "Component's Export File Full Name registered"
                        End If
                    End If
                End With
        End Select
                                
next_vbc:
        mTrc.EoC ErrSrc(PROC) & " " & vbc.name
        lCompsRemaining = lCompsRemaining - 1
        Set cComp = Nothing
        Set cRaw = Nothing
    Next vbc
    
    sMsg = mCompMan.Service
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

Private Function FolderNotVbProjectExclusive(ByVal wb As Workbook) As Boolean

    Dim fso As New FileSystemObject
    Dim fl  As File
    
    For Each fl In fso.GetFolder(wb.Path).Files
        If fl.Path <> wb.FullName And VBA.Left$(fso.GetFileName(fl.Path), 2) <> "~$" Then
            Select Case fso.GetExtensionName(fl.Path)
                Case "xlsm", "xlam", "xlsb" ' may contain macros, a VB-Project repsectively
                    FolderNotVbProjectExclusive = True
                    Exit For
            End Select
        End If
    Next fl

End Function

Public Sub ExportAll(Optional ByVal ea_wb As Workbook = Nothing)
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
    If mService.Denied(den_serviced_wb:=ea_wb, den_service:=ErrSrc(PROC)) Then GoTo xt
    
    mCompMan.Service = PROC & " for '" & ea_wb.name & "': "
    sStatus = mCompMan.Service

    mCompMan.DeleteObsoleteExpFiles ea_wb
    
    If mMe.IsAddinInstnc _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
    
    For Each vbc In ea_wb.VbProject.VBComponents
        Set cComp = New clsComp
        With cComp
            .Wrkbk = ea_wb
            .CompName = vbc.name ' this assignment provides the name for the export file
            vbc.Export .ExpFileFullName
        End With
        Set cComp = Nothing
    Next vbc

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub SynchVbProject( _
           Optional ByVal sp_clone_project As Workbook = Nothing, _
           Optional ByVal sp_raw_project As String = vbNullString)
' -----------------------------------------------------------------
' Synchronizes the code of the Workbook (clone_project) with the
' code of the VB-Raw-Project (raw_project).
' Note: The sp_raw_project is identified by its 'BaseName' when
'       the service is called from the VB-Clone-Project. When
'       directly called via the immediate window both arguments
'       are provided via a file selection dialog.
' The service is performed provided:
' - the CompMan basic configuration is complete and valid
' - the Workbook is located within the configured "Serviced Root"
' - the Workbook is the only one within its parent folder
' - the CompMan Addin is not paused
' - the open/ed Workbook is not a restored version
' -------------------------------------------------------------
    Const PROC = "SynchVbProject"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim sStatus             As String
    Dim wbRaw               As Workbook
    
    mErH.BoP ErrSrc(PROC)
    If sp_clone_project Is Nothing Then Set sp_clone_project = ServicedWrkbk(PROC)
    If sp_clone_project Is Nothing Then GoTo xt
    If sp_raw_project = vbNullString Then
        Set wbRaw = ServicedWrkbk(PROC)
        If Not wbRaw Is Nothing Then sp_raw_project = fso.GetBaseName(wbRaw.FullName)
    End If
    If sp_raw_project = vbNullString Then GoTo xt
    
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    If mService.Denied(den_serviced_wb:=sp_clone_project, den_service:=ErrSrc(PROC)) Then GoTo xt

    mCompMan.Service = PROC & " for '" & sp_clone_project.name & "': "
    sStatus = mCompMan.Service
    
    If Not mRawHosts.Exists(sp_raw_project) Then
        cLog.Entry = "The Clone-VB-Project claims an invalid Raw_VB-Project! ('" & sp_raw_project & "' is not a registered Raw-VB-Project)"
        VBA.MsgBox Title:="The VB-Clone-Project refers to an invalid VB-Raw-Project!" _
                         , Prompt:="A VB-Raw-Project '" & sp_raw_project & "' is unknown/not registered."
        GoTo xt
    End If
    
    mSync.VbProject clone_project:=sp_clone_project _
                  , raw_project:=mRawHosts.FullName(sp_raw_project)

xt: mErH.EoP ErrSrc(PROC)
    Set cLog = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub UpdateRawClones( _
                     ByVal uc_wb As Workbook, _
            Optional ByVal uc_hosted As String = vbNullString)
' ------------------------------------------------------------
' Updates all clone components with the Export File of the
' raw component provided the raw's code has changed.
' ------------------------------------------------------------
    Const PROC = "UpdateRawClones"
    
    On Error GoTo eh
    Dim wbActive    As Workbook
    Dim wbTemp      As Workbook
    Dim sStatus     As String
    Dim lCompMaxLen As Long
    
    mErH.BoP ErrSrc(PROC)
    If mService.Denied(den_serviced_wb:=uc_wb, den_service:=ErrSrc(PROC), den_new_log:=True) Then GoTo xt

    mCompMan.Service = PROC & " for '" & uc_wb.name & "': "
    Application.StatusBar = sStatus & "Maintain hosted raws"
    If Not mCompMan.ManageVbProjectProperties(mh_hosted:=uc_hosted _
                                            , mh_wb:=uc_wb) _
    Then
        Application.StatusBar = cLog.Service & " failed. See log-file in Workbookfolder"
        GoTo xt
    End If
        
    Application.StatusBar = sStatus & "De-activate '" & uc_wb.name & "'"
    If uc_wb Is ActiveWorkbook Then
        '~~ De-activate the ActiveWorkbook by creating a temporary Workbook
        Set wbActive = uc_wb
        Set wbTemp = Workbooks.Add
    End If
    
    mUpdate.RawClones urc_wb:=uc_wb _
                    , urc_comp_max_len:=lCompMaxLen _
                    , urc_clones:=mCompMan.Clones(uc_wb)

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

Private Function WbkInServicedRoot(ByVal idr_wb As Workbook) As Boolean
    WbkInServicedRoot = InStr(idr_wb.Path, mMe.ServicedRoot) <> 0
End Function

Private Function WbkIsRestoredBySystem(ByVal rbs_wb As Workbook) As Boolean
    WbkIsRestoredBySystem = InStr(ActiveWindow.Caption, "(") <> 0 _
                         Or InStr(rbs_wb.FullName, "(") <> 0
End Function

Public Sub Install(Optional ByVal in_wb As Workbook = Nothing)
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
