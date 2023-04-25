VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ----------------------------------------------------------------------------
' Data Module wsConfig: Keeps a record of all configuration data used
' ===================== throughout all services.
'
' ----------------------------------------------------------------------------
Private sRenewResult                        As String

Private Property Let AddInStatus(ByVal s As String):                 MaintainMinRowHeight Me.Range("AddInStatus"), s:                        End Property

Private Property Let AutoOpenStatus(ByVal s As String):              MaintainMinRowHeight Me.Range("AutoOpenStatus"), s:                     End Property

Public Property Let ButtonCaption(Optional ByVal b_cmb As Button, _
                                           ByVal b_caption As String)
    Dim b   As Boolean
    
    With b_cmb.Parent
        If .ProtectContents Then
            b = True
            .Unprotect
        End If
        b_cmb.Text = b_caption
        If b Then .Protect
    End With

End Property

Private Property Let ButtonEnabled(Optional ByVal b_cmb As Button, _
                                            ByVal b_enabled As Boolean)
' ----------------------------------------------------------------------------
' Because the Enabled property does not what is expected the button is hidden
' or displayed instead.
' ----------------------------------------------------------------------------
    Dim bProtected  As Boolean
    Dim wsh         As Worksheet
    
    Set wsh = b_cmb.Parent
    If wsh.ProtectContents Then
        bProtected = True
        wsh.Unprotect
    End If
                                           
    With b_cmb
        If b_enabled Then
            .Visible = True
            .Font.Color = rgbBlack
        Else
            .Visible = False
            .Font.Color = rgbGrey
        End If
    End With
    
    If bProtected Then wsh.Protect

End Property

Private Property Get ButtonAddInProvideGiveUp() As Button:          Set ButtonAddInProvideGiveUp = Me.Buttons("cmbAddInSetUpGiveUp"):       End Property

Private Property Let ButtonHelpUrlAddinSetup(ByVal s As String):    UrlButton(Me.Range("ButtonHelpUrlAddinSetup")) = s:                     End Property

Private Property Let ButtonHelpUrlPauseContinue(ByVal s As String): UrlButton(Me.Range("ButtonHelpUrlPauseContinue")) = s:                  End Property

Private Property Get ButtonPauseAddIn() As Button:                  Set ButtonPauseAddIn = Me.Buttons("cmbAddInPause"):                     End Property

Private Property Get ButtonAddInRenew() As Button:                  Set ButtonAddInRenew = Me.Buttons("cmbAddInRenew"):                     End Property

Private Property Get ButtonSetupRenewAddIn() As Button:             Set ButtonSetupRenewAddIn = Me.Buttons("cmbAddInRenew"):                End Property

Private Property Get ButtonWrkbkAutoOpen() As Button:               Set ButtonWrkbkAutoOpen = Me.Buttons("cmbWrkbkAutoOpen"):               End Property

Private Property Let CompManAddInStatusUrl(ByVal s As String):      UrlLabel(Me.Range("CompManAddInStatusUrl")) = s:                        End Property

Private Property Let CompManWrkbkStatusUrl(ByVal s As String):      UrlLabel(Me.Range("CompManWrkbkStatusUrl")) = s:                        End Property

Private Property Let CommonComponentsFolderUrl(ByVal s As String):  UrlLabel(Me.Range("CommonComponentsFolderUrl")) = s:                    End Property

Public Property Let ConfigInfo(ByVal s As String):                  mWsh.Value(Me, "ConfigInfo") = s:                                       End Property

Private Property Get FolderAddin() As String:                       FolderAddin = mWsh.Value(Me, "FolderAddin"):                            End Property

Public Property Let FolderAddin(ByVal s As String):                 mWsh.Value(Me, "FolderAddin") = s:                                      End Property

Private Property Let FolderAddInInfo(ByVal s As String):            MaintainMinRowHeight Me.Range("FolderAddInInfo"), s:                    End Property

Private Property Let FolderAddInUrl(ByVal s As String):             UrlLabel(Me.Range("FolderAddInUrl")) = s:                               End Property

Public Property Get FolderCommonComponentsName() As String
    Dim fso As New FileSystemObject
    FolderCommonComponentsName = fso.GetFolder(FolderCommonComponentsPath).Name
    Set fso = Nothing
End Property

Public Property Get FolderCommonComponentsPath() As String:         FolderCommonComponentsPath = mWsh.Value(Me, "FolderCommonComponents"):  End Property

Public Property Let FolderCommonComponentsPath(ByVal s As String):  mWsh.Value(Me, "FolderCommonComponents") = s:                           End Property

Public Property Get FolderCompManRoot() As String:                  FolderCompManRoot = mWsh.Value(Me, "FolderCompManRoot"):                End Property

Public Property Let FolderCompManRoot(ByVal s As String):           mWsh.Value(Me, "FolderCompManRoot") = s:                                End Property

Private Property Let FolderCompManRootInfo(ByVal s As String):      MaintainMinRowHeight Me.Range("FolderCompManRootInfo"), s:              End Property

Private Property Let FolderCompManRootUrl(ByVal s As String):       UrlLabel(Me.Range("FolderCompManRootUrl")) = s:                         End Property

Public Property Get FolderExport() As String:                       FolderExport = mWsh.Value(Me, "FolderExport"):                          End Property

Private Property Let FolderExport(ByVal s As String):               mWsh.Value(Me, "FolderExport") = s:                                     End Property

Private Property Get FolderExportInfo() As String:                  FolderExportInfo = mWsh.Value(Me, "FolderExportInfo"):                  End Property

Private Property Let FolderExportInfo(ByVal s As String):           MaintainMinRowHeight Me.Range("FolderExportInfo"), s:                   End Property

Private Property Let FolderExportUrl(ByVal s As String):            UrlLabel(Me.Range("FolderExportUrl")) = s:                              End Property

Public Property Get FolderSyncArchive() As String:                  FolderSyncArchive = mWsh.Value(Me, "FolderSyncArchive"):                End Property

Public Property Let FolderSyncArchive(ByVal s As String):           mWsh.Value(Me, "FolderSyncArchive") = s:                                End Property

Private Property Let FolderSyncArchiveInfo(ByVal s As String):      MaintainMinRowHeight Me.Range("FolderSyncArchiveInfo"), s:              End Property

Private Property Let FolderSyncArchiveUrl(ByVal s As String):       UrlLabel(Me.Range("FolderSyncArchiveUrl")) = s:                         End Property

Public Property Get FolderSyncTarget() As String:                   FolderSyncTarget = mWsh.Value(Me, "FolderSyncTarget"):                  End Property

Public Property Let FolderSyncTarget(ByVal s As String):            mWsh.Value(Me, "FolderSyncTarget") = s:                                 End Property

Private Property Let FolderSyncTargetInfo(ByVal s As String):       MaintainMinRowHeight Me.Range("FolderSyncTargetInfo"), s:               End Property

Private Property Let FolderSyncTargetUrl(ByVal s As String):        UrlLabel(Me.Range("FolderSyncTargetUrl")) = s:                          End Property

Private Property Get MonitorAnchor() As Range:                      Set MonitorAnchor = Me.Range("ConfigInfo"):                             End Property

Private Property Let UrlButton(ByVal url_rng As Range, _
                               ByVal url_chapter As String)
                        
    mWsh.Url(Me, url_rng, xlUnderlineStyleNone, 11) = Replace(Replace(mCompMan.README_URL & url_chapter, "##", "#"), "?", vbNullString)
    
End Property

Private Property Let UrlLabel(ByVal url_rng As Range, _
                              ByVal url_chapter As String)
                        
    mWsh.Url(Me, url_rng, xlUnderlineStyleNone, 9) = Replace(Replace(mCompMan.README_URL & url_chapter, "##", "#"), "?", vbNullString)
    
End Property

Public Function AutoOpenAddinIsSetup() As Boolean
    With New FileSystemObject
        AutoOpenAddinIsSetup = .FileExists(mAddin.AutoOpenShortCut)
    End With
End Function

Public Function AutoOpenAddinRemove() As Boolean
    With New FileSystemObject
        If .FileExists(mAddin.AutoOpenShortCut) Then .DeleteFile mAddin.AutoOpenShortCut
    End With
End Function

Public Sub AutoOpenAddinSetup()
' ----------------------------------------------------------------------------
' Setup/replace the CompMan Add-in auto-open as a shortcut to the Add-in-
' Workbook in the users \AppData\Microsoft\Excel\XLSTART folder.
' ----------------------------------------------------------------------------
    Const PROC = "AutoOpenAddinSetup"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim oShell      As IWshShell3
    Dim oShortcut   As IWshShortcut
    Dim sShortcut   As String
    
    If mMe.IsAddinInstnc Then GoTo xt
    sShortcut = mAddin.AutoOpenShortCut
    If AutoOpenAddinIsSetup Then fso.DeleteFile sShortcut
    Set oShell = CreateObject("Wscript.shell")
    Set oShortcut = oShell.CreateShortcut(sShortcut)
    With oShortcut
        .TargetPath = mAddin.WbkFullName
        .Description = "AutoOpen CompMan Add-in"
        .Save
    End With
    wsConfig.CurrentStatus
    
xt: Set oShell = Nothing
    Set oShortcut = Nothing
    Exit Sub
 
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CompManAddinProvideGiveUpFlipFlop()
' ----------------------------------------------------------------------------
' Called by CommandButton cmbAddInSetUpGiveUp:
' Flipflops between "Provide Add-in" and "Give-Up add-in".
' Note: The Renew Add-in button is only visible when the Add-in is provided.
' ----------------------------------------------------------------------------
    Const PROC = "AutoOpenCompManFlipFlop"
    
    On Error GoTo eh
    If mMe.IsAddinInstnc Then GoTo xt
    Application.ScreenUpdating = False
    If AutoOpenAddinIsSetup Then
        mAddin.GiveUp
        ButtonCaption(ButtonAddInProvideGiveUp) = "Provide Add-in"
        ButtonAddInRenew.Visible = False
        ButtonPauseAddIn.Visible = False
    Else
        mAddin.SetupRenew
        ButtonCaption(ButtonAddInProvideGiveUp) = "Give up Add-in"
        ButtonAddInRenew.Visible = True
        ButtonPauseAddIn.Visible = True
    End If
    wsConfig.CurrentStatus
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AutoOpenCompManFlipFlop()
' ----------------------------------------------------------------------------
' Called by CommandButton cmbWrkbkAutoOpen:
' Flips the Auto-Open status of the CompMan Workbook. When auto-open is setup
' the auto-open is removed, else it is setup.
' ----------------------------------------------------------------------------
    Const PROC = "AutoOpenCompManFlipFlop"
    
    On Error GoTo eh
    
    Application.ScreenUpdating = False
    If mMe.IsAddinInstnc Then GoTo xt
    If AutoOpenCompManIsSetup Then
        AutoOpenCompManRemove
        ButtonCaption(ButtonWrkbkAutoOpen) = "Setup Auto-open"
    Else
        AutoOpenCompManSetup
        ButtonCaption(ButtonWrkbkAutoOpen) = "Remove Auto-open"
    End If
    Me.Verified
    
xt: Exit Sub
 
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function AutoOpenCompManIsSetup() As Boolean
    With New FileSystemObject
        AutoOpenCompManIsSetup = .FileExists(AutoOpenShortCutCompManWbk)
    End With
End Function

Public Sub AutoOpenCompManRemove()
    With New FileSystemObject
        If .FileExists(AutoOpenShortCutCompManWbk) Then .DeleteFile AutoOpenShortCutCompManWbk
    End With
End Sub

Public Sub AutoOpenCompManSetup()
    Dim oShell      As IWshShell3
    Dim oShortcut   As IWshShortcut

    Set oShell = CreateObject("Wscript.shell")
    Set oShortcut = oShell.CreateShortcut(AutoOpenShortCutCompManWbk)
    With oShortcut
        .TargetPath = ThisWorkbook.FullName
        .Description = "AutoOpen CompMan.xlsb"
        .Save
    End With
    
End Sub

Public Sub CurrentStatus()
' ----------------------------------------------------------------------------
' Analysis the status of the CompMan Add-in and the CompMan Workbook.
' ----------------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sStatus     As String
    Dim cll         As New Collection
    Dim i           As Long
    Dim sShortcut   As String
    Dim sAddin      As String
    Dim sComma      As String
    
    If mMe.IsAddinInstnc Then GoTo xt
    
    '~~ CompMan Add-in setup status
    If mAddin.Folder <> vbNullString Then
        If fso.FileExists(mAddin.WbkFullName) Then
            cll.Add "provided"
            
            '~~ CompMan Addin Paused status
            If mAddin.IsOpen Then
                ButtonPauseAddIn.Visible = True
                If mAddin.Paused Then
                    cll.Add "open but currently  p a u s e d !"
                    ButtonCaption(ButtonPauseAddIn) = "Continue Add-in"
                Else
                    ButtonCaption(ButtonPauseAddIn) = "Pause Add-in"
                End If
                cll.Add "open"
            Else
                cll.Add "not open"
                ButtonPauseAddIn.Visible = False
            End If

            '~~ CompMan Addin AutoOpen status
            sShortcut = mAddin.AutoOpenShortCut
            If fso.FileExists(sShortcut) Then
                '~~ Update the shortcut in case the location of the Add-in Workbook has changed
                If ShortCutTargetPath(sShortcut) <> mAddin.WbkFullName _
                Then ShortCutTargetPath(sShortcut) = mAddin.WbkFullName
                cll.Add "opened when Excel starts"
            Else
                cll.Add "auto-open at Excel start not setup"
            End If
        Else
            cll.Add "not provided"
        End If
    Else
        '~~ may have been de-configured
        If mAddin.FolderOld <> mAddin.Folder Then
            sAddin = mAddin.FolderOld & "\" & mAddin.WbkName
            If fso.FileExists(sAddin) Then
                cll.Add "still set up in former folder"
            Else
                cll.Add "not set up"
            End If
        End If
    End If
        
    sComma = vbNullString
    For i = 1 To cll.Count
        sStatus = sStatus & sComma & cll(i)
        sComma = ", "
    Next i
    AddInStatus = sStatus
    
    '~~ CompMan Workbook AutoOpen status
    sStatus = vbNullString
    If fso.FileExists(AutoOpenShortCutCompManWbk) Then
        '~~ Update the shortcut in case the location of the CompMan Workbook has changed
        If ShortCutTargetPath(AutoOpenShortCutCompManWbk) <> ThisWorkbook.FullName _
        Then ShortCutTargetPath(AutoOpenShortCutCompManWbk) = ThisWorkbook.FullName
        AutoOpenStatus = "auto-open for " & ThisWorkbook.Name & "  s e t u p."
    Else
        AutoOpenStatus = "auto-open for " & ThisWorkbook.Name & "  n o t  s e t u p !"
    End If
    ButtonHelpUrlAddinSetup = "#configuration-changes-compmans-config-worksheet"
    ButtonHelpUrlPauseContinue = "#configuration-changes-compmans-config-worksheet"

    FolderSyncTargetIsValid
    FolderSyncArchiveIsValid
        
xt: Set fso = Nothing
    Set cll = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Me.CodeName & "." & sProc
End Function

Public Function FolderAddInIsValid() As Boolean
' ----------------------------------------------------------------------------
' - Returns TRUE when the AddIn folder is configured and valid (i.e. exists)
' - Sets status and folder info in accordance with the current status.
' - Considers that the configured folder does not or no longer exist, in case
'   the Workbook is opened on a different computer for example.
' ----------------------------------------------------------------------------
    Const PROC = "FolderAddInIsValid"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim bValid  As Boolean
    
    FolderAddInUrl = mCompMan.README_CONFIG_CHANGES
    ButtonHelpUrlAddinSetup = "#setup-auto-open-for-the-compman-workbook"

    Select Case True
        Case mAddin.Folder = vbNullString
            FolderAddInInfo = "Not configured! A CompMan Add-in instance cannot be setup " & _
                               "(an available Add-in will have been removed)."
            bValid = False
        Case mAddin.Folder <> vbNullString
            If fso.FolderExists(mAddin.Folder) Then
                FolderAddInInfo = "Configured folder for the CompMan Add-in, possibly setup/renewed." & vbLf & _
                                   "Attention: In case a different folder is selected/established the previous folder becomes obsolete but will not be removed!"
                bValid = True
                bValid = True
            Else
                mAddin.Folder = vbNullString
                FolderAddInInfo = "Not configured! The former configured folder does not exist and thus had been reset."
                
                If mAddin.FolderOld <> vbNullString Then
                    If Not fso.FolderExists(mAddin.FolderOld) Then mAddin.FolderOld = vbNullString
                End If
                
                bValid = False
            End If
        Case mAddin.Folder = vbNullString
            FolderAddInInfo = "Not configured! " & vbLf & _
                               "Folder for the CompMan Add-in instance."
            bValid = False
    End Select
    wsConfig.CurrentStatus

    If bValid _
    Then ButtonSetupRenewAddIn.Visible = True _
    Else ButtonSetupRenewAddIn.Visible = False
    
xt: FolderAddInIsValid = bValid
    Set fso = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FolderCompManRootIsValid() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current configured Synchronization-Archive-Folder is
' valid. Considers that the configured folder does not or no longer exist, in
' case the Workbook is opened on a different computer for example.
' ----------------------------------------------------------------------------
    Const PROC = "FolderCompManRootIsValid"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim bValid  As Boolean
    
    FolderCompManRootUrl = mCompMan.README_CONFIG_CHANGES
    CompManWrkbkStatusUrl = mCompMan.README_CONFIG_CHANGES
    CommonComponentsFolderUrl = mCompMan.README_DEFAULT_FILES_AND_FOLDERS
    
    If FolderCompManRoot <> vbNullString Then
        If fso.FolderExists(FolderCompManRoot) Then
            FolderCompManRootInfo = "Folder serviced by CompMan. I.e. CompMan provides its services for any Workbook " & _
                                    "located (opened, saved) in this folder - provided the Workbook has one or more " & _
                                    "services enabled."
            If FolderExportIsValid Then ConfigInfo = vbNullString
            bValid = True
        Else
            FolderCompManRootInfo = "Not configured! The former configured folder does not exist and thus has been reset!." & vbLf & _
                                    "CompMan 's services are  n o t   a v a i l a b l e !" & vbLf & _
                                    "CompMan only serves Workbooks which are opened from within a configured 'Serviced Development and Test' folder."
            FolderCompManRoot = vbNullString
        End If
    Else
        FolderCompManRootInfo = "Missing! CompMan's services are  n o t   a v a i l a b l e !" & vbLf & _
                                "CompMan only serves Workbooks which are opened from within a configured 'Serviced Development and Test' folder."
    End If
    ButtonWrkbkAutoOpen.Visible = bValid
    
xt: FolderCompManRootIsValid = bValid
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub FolderExportConfigure()
' ----------------------------------------------------------------------------
' Called by Commandbutton cmbFolderExportSpecify:
' ----------------------------------------------------------------------------
    Dim sNew    As String
    Dim sOld    As String
    
    sOld = wsConfig.FolderExport
    sNew = VBA.InputBox(Prompt:="Enter the name for the folder to which changed components are exported " & _
                             "by CompMan's 'Export Changed Components' service (for Workbooks which a " & _
                             "correspondingly configured and reside within the configured serrviced 'Development and Test' folder." & vbLf & vbLf & _
                             "Attention! When the name of the Export folder is changed, all old named folders are renamed correspondingly." _
                    , Title:="Export folder name" _
                    , Default:=sOld)
    If sNew <> vbNullString Then
        FolderExport = sNew
        If sNew <> sOld And FolderCompManRootIsValid Then
            '~~ Provided a 'Serviced Development And Test' folder is configured,
            '~~ all export folders are renamed in accordance with the name change
            RenameSubFolders wsConfig.FolderCompManRoot, sOld, sNew
        End If
    Else
        FolderExport = sOld
    End If
    
End Sub

Public Function FolderExportIsValid(Optional ByRef iv_info As String) As Boolean
' ----------------------------------------------------------------------------
' Verification the name of the Export-Folder
' ----------------------------------------------------------------------------
    With Me
        FolderExportUrl = mCompMan.README_CONFIG_CHANGES
        If .FolderExport <> vbNullString Then
            FolderExportInfo = "Current configured name of the folder within a serviced Workbook's dedicated folder " & _
                                "into which CompMan exports new/modified components. The name may be changed however."
            FolderExportIsValid = True
        Else
            FolderExport = "source"
            FolderExportInfo = "Current configured (default) name of the folder within a serviced Workbook's dedicated " & _
                                "folder into which CompMan exports new/modified components. The name may be changed however."
            FolderExportIsValid = False
        End If
        iv_info = FolderExportInfo
    End With
End Function

Public Sub FolderSyncArchiveConfigure()
    FolderSyncArchive = mBasic.SelectFolder("Select the 'Synchronization Archive' folder for CompMan's 'Synchronize VB-Project' service")
    FolderSyncArchiveIsValid
    wsConfig.CurrentStatus
End Sub

Private Function FolderSyncArchiveIsValid() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current configured Synchronization-Archive-Folder is
' valid.
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    
    With Me
        FolderSyncArchiveUrl = mCompMan.README_CONFIG_CHANGES
        If FolderSyncArchive <> vbNullString Then
            If fso.FolderExists(FolderSyncTarget) Then
                FolderSyncArchiveIsValid = True
            Else
                FolderSyncArchive = vbNullString
                FolderSyncTargetInfo = "Not configured! The former configured folder does not exist and thus had been reset." & vbLf & _
                                        "CompMan's 'VB-Project-Synchronization' service will not be provided even when a Workbook " & _
                                        "has it enabled." & vbLf & _
                                        "Important! The folder must not be a sub-folder of the 'Serviced Development and Test' folder!"
            End If
        Else ' .FolderSyncArchive = vbNullString
            If FolderSyncTarget <> vbNullString Then
                FolderSyncArchiveInfo = "Not configured! When a Sync-Target-Folder is configured the configuration of a Sync-Archive-Folder is obligatory!"
            Else
                FolderSyncArchiveInfo = "Configuration only becomes obligatory when a Sync-Target-folder is configured!"
                FolderSyncArchiveIsValid = True
            End If
        End If
        If FolderSyncArchiveIsValid Then
            FolderSyncArchiveInfo = "The Sync-Target-Workbook will be copied into this folder before it gets synchronized with its corresponding Sync-Source-Workbook."
        End If
    End With
    
    Set fso = Nothing

End Function

Public Sub FolderSyncTargetConfigure()
    FolderSyncTarget = mBasic.SelectFolder("Select the folder for the CompMan's 'Synchronize VB-Project' service")
End Sub

Public Function FolderSyncTargetIsValid() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Sync-Target-Folder is either not configured ore the
' configured ones is valid.
' Considers that the configured folder does not or no longer exist, in case
' the Workbook is opened on a different computer for example.
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    
    With wsConfig
        FolderSyncTargetIsValid = FolderSyncTarget = vbNullString
        FolderSyncTargetUrl = mCompMan.README_CONFIG_CHANGES
        If FolderSyncTarget <> vbNullString Then
            If fso.FolderExists(FolderSyncTarget) Then
                FolderSyncTargetIsValid = True
                FolderSyncTargetInfo = "CompMan's folder for the 'Synchronization VB-Project' service. The service is " & _
                                       "provided for any Workbook prepared for the service and opened from within this folder."
            Else
                FolderSyncTarget = vbNullString
                FolderSyncTargetInfo = "Not configured! The former configured folder does not exist and thus had been reset." & vbLf & _
                                       "CompMan's 'Synchronize VB-Project' service will not be provided even when a Workbook " & _
                                       "has it enabled." & vbLf & _
                                       "Important! The folder must not be a sub-folder of the 'Serviced Development and Test' folder!"
            End If
        Else
            FolderSyncTargetInfo = "No 'Synchronization-Target' folder configured. CompMan's 'VB-Project Synchronization' service" & _
                                   "will not be provided even when a Workbook has it enabled." & vbLf & _
                                   "Important! The folder must not be a sub-folder of the 'Serviced Development and Test' folder!"
            FolderSyncTargetIsValid = True
        End If
    End With
    Set fso = Nothing
    
End Function

Private Sub MaintainMinRowHeight(ByVal mrh_rng As Range, _
                                 ByVal mrh_str As String)
    Dim bProtected  As Boolean
    Dim wsh         As Worksheet
    
    Set wsh = mrh_rng.Parent
    If wsh.ProtectContents Then
        bProtected = True
        wsh.Unprotect
    End If
    
    mWsh.Value(Me, mrh_rng) = mrh_str
    With mrh_rng
        .EntireRow.AutoFit
        If .RowHeight < 24 Then .RowHeight = 24
    End With
    
    If bProtected Then wsh.Protect

End Sub

Public Sub MonitorRenewStep(Optional ByVal rn_result As String = vbNullString, _
                            Optional ByVal rn_last_step As Boolean = False, _
                            Optional ByVal rn_renew_failed As Boolean = False)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    Dim Title           As String:      Title = "Establish or renew CompMan as Add-in"
    Dim Step            As TypeMsgText
    Dim Footer          As TypeMsgText
    Dim sProgressMsg    As String
    
    If rn_result <> vbNullString Then sRenewResult = rn_result
    If sRenewResult <> vbNullString Then sProgressMsg = ProgressMsg(RenewStep, RenewAction, sRenewResult)
    Step.FontSize = 8
    Footer.FontSize = 8
    Step.MonoSpaced = True
    
    If mMe.RenewStep > 0 And sRenewResult <> vbNullString Then
        Step.Text = sProgressMsg
        Step.MonoSpaced = True
        If Not rn_last_step Then
            mMsg.Monitor mon_title:=Title _
                       , mon_text:=Step _
                       , mon_steps_displayed:=12 _
                       , mon_width_min:=70 _
                       , mon_width_max:=80 _
                       , mon_pos:=MonitorAnchor
        Else
            mMsg.Monitor Title, Step
            If mMe.RenewTerminatedByUser Then
                Footer.Text = "Renew process terminated by user! Close window."
                mMsg.MonitorFooter Title, Footer
            Else
                If Not rn_renew_failed Then
                    With Footer
                        .Text = "Renew process  " & mBasic.Spaced("successfully") & "  finished! Close window."
                        .FontBold = True
                        .FontColor = rgbDarkGreen
                    End With
                    mMsg.MonitorFooter Title, Footer
                Else
                    With Footer
                        .Text = "Renew process  " & mBasic.Spaced("failed") & "  Close window and " & vbLf & _
                                "the only way to solve this problem is:" & vbLf & _
                                "1. Close all application and restart the computer" & vbLf & _
                                "2. Remove the CompMan.xlma from its folder" & vbLf & _
                                "3. Reopen " & mCompManClient.COMPMAN_DEVLP & " and repeat this Renew (then a setup in fact)."
                        .FontColor = rgbRed
                    End With
                    mMsg.MonitorFooter Title, Footer
                End If
            End If
        End If
        sProgressMsg = vbNullString
    End If
    sRenewResult = vbNullString
    
End Sub

Private Function ProgressMsg(ByVal pm_step As Long, _
                             ByVal pm_action As String, _
                             ByVal pm_result As String) As String
    If ProgressMsg = vbNullString Then
        ProgressMsg = mBasic.Align(pm_step, 2, AlignRight) & " " & pm_result & " " & Split(pm_action, vbLf)(0)
        If UBound(Split(pm_action, vbLf)) > 0 Then
            ProgressMsg = ProgressMsg & vbLf & "          " & Split(pm_action, vbLf)(1)
        End If
    Else
        ProgressMsg = ProgressMsg & vbLf & mBasic.Align(pm_step, 2, AlignRight) & " " & pm_result & " " & Split(pm_action, vbLf)(0)
        If UBound(Split(RenewAction, vbLf)) > 0 Then
            ProgressMsg = ProgressMsg & vbLf & "   " & Split(pm_action, vbLf)(1)
        End If
    End If
End Function

Public Function Verified() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when all configured folders are valid.
' ----------------------------------------------------------------------------
    Verified = FolderCompManRootIsValid _
           And FolderAddInIsValid _
           And FolderSyncTargetIsValid _
           And FolderSyncArchiveIsValid
    CompManAddInStatusUrl = mCompMan.README_CONFIG_CHANGES
    CompManWrkbkStatusUrl = mCompMan.README_CONFIG_CHANGES
    CommonComponentsFolderUrl = mCompMan.README_DEFAULT_FILES_AND_FOLDERS
    
End Function

Private Sub Worksheet_Activate()
    
    If mMe.IsAddinInstnc Then GoTo xt
    Me.Verified
    With Application
        .ScreenUpdating = False
        .DisplayFormulaBar = False
    End With
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayHeadings = False
    End With
    Me.Range("celParking").Select ' selects a non visible cell
    
xt: Exit Sub

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    With Me
        Select Case Target.Address
            Case .Range("FolderAddInInfo").Address, _
                 .Range("FolderCompManRootInfo").Address, _
                 .Range("FolderExportInfo").Address, _
                 .Range("FolderSyncArchiveInfo").Address, _
                 .Range("FolderSyncTargetInfo").Address, _
                 .Range("ConfigInfo").Address
                If InStr(Target.Value, "nvalid!") <> 0 _
                Or InStr(Target.Value, "issing!") <> 0 _
                Or InStr(Target.Value, "ot configured!") <> 0 Then
                    Target.Font.Color = 255
                Else
                    Target.Font.ColorIndex = xlAutomatic
                End If
            Case .Range("FolderCompManRoot").Address:   FolderCompManRootIsValid
            Case .Range("FolderAddIn").Address:         FolderAddInIsValid
            Case .Range("FolderSyncTarget").Address:    FolderSyncTargetIsValid
            Case .Range("FolderSyncArchive").Address:   FolderSyncArchiveIsValid
        End Select
    
        Select Case Target.Address
            Case .Range("FolderCompManRoot").Address, _
                 .Range("FolderExport").Address, _
                 .Range("FolderAddin").Address, _
                 .Range("FolderSyncTarget").Address, _
                 .Range("FolderSyncArchive").Address
        End Select
    End With


End Sub

