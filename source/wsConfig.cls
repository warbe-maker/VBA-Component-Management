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

Private bFolderServicedIsValid              As Boolean
Private bFolderAddinIsValid                 As Boolean
Private bFolderExportIsValid                As Boolean
Private bFolderSyncedIsValid                As Boolean
Private bFolderSyncArchiveIsValid           As Boolean
Private bAddinConfigObligatory              As Boolean
Private bSyncConfigObligatory               As Boolean
Private bCanceled                           As Boolean
Private sRenewResult                        As String

Public Property Let AddInStatus(ByVal s As String):                 MaintainMinRowHeight Me.Range("AddInStatus"), s:                End Property

Public Property Let AutoOpenStatus(ByVal s As String):              MaintainMinRowHeight Me.Range("AutoOpenStatus"), s:             End Property

Public Property Let ButtonHelpUrlAddinSetup(ByVal s As String):     UrlButton(Me.Range("ButtonHelpUrlAddinSetup")) = s:             End Property

Public Property Let ButtonHelpUrlPauseContinue(ByVal s As String):  UrlButton(Me.Range("ButtonHelpUrlPauseContinue")) = s:          End Property

Public Property Get ButtonSetupRenewAddIn() As Button:              Set ButtonSetupRenewAddIn = Me.Buttons("cmbAddInSetup"):        End Property

Public Property Get ButtonWrkbkAutoOpen() As Button:                Set ButtonWrkbkAutoOpen = Me.Buttons("cmbWrkbkAutoOpen"):       End Property

Public Property Let CompManAddInStatusUrl(ByVal s As String):       UrlLabel(Me.Range("CompManAddInStatusUrl")) = s:                End Property

Public Property Let CompManWrkbkStatusUrl(ByVal s As String):       UrlLabel(Me.Range("CompManWrkbkStatusUrl")) = s:                End Property

Public Property Let ConfigInfo(ByVal s As String):                  mWsh.Value(Me, "ConfigInfo") = s:                               End Property

Public Property Let FolderAddInInfo(ByVal s As String):             MaintainMinRowHeight Me.Range("FolderAddInInfo"), s:            End Property

Public Property Let FolderAddInUrl(ByVal s As String):              UrlLabel(Me.Range("FolderAddInUrl")) = s:                       End Property

Public Property Get FolderDevAndTest() As String:                   FolderDevAndTest = mWsh.Value(Me, "FolderDevAndTest"):          End Property

Public Property Let FolderDevAndTest(ByVal s As String):            mWsh.Value(Me, "FolderDevAndTest") = s:                         End Property

Public Property Let FolderDevAndTestInfo(ByVal s As String):        MaintainMinRowHeight Me.Range("FolderDevAndTestInfo"), s:       End Property

Public Property Let FolderDevAndTestUrl(ByVal s As String):         UrlLabel(Me.Range("FolderDevAndTestUrl")) = s:                  End Property

Public Property Get FolderExport() As String
    If mWsh.Value(Me, "FolderExport") = vbNullString Then
        mWsh.Value(Me, "FolderExport") = "source"
    End If
    FolderExport = mWsh.Value(Me, "FolderExport")
End Property

Public Property Let FolderExport(ByVal s As String):                mWsh.Value(Me, "FolderExport") = s:                             End Property

Public Property Get FolderExportInfo() As String:                   FolderExportInfo = mWsh.Value(Me, "FolderExportInfo"):          End Property

Public Property Let FolderExportInfo(ByVal s As String):            MaintainMinRowHeight Me.Range("FolderExportInfo"), s:           End Property

Public Property Let FolderExportUrl(ByVal s As String):             UrlLabel(Me.Range("FolderExportUrl")) = s:                      End Property

Public Property Get FolderSyncArchive() As String:                  FolderSyncArchive = mWsh.Value(Me, "FolderSyncArchive"):        End Property

Public Property Let FolderSyncArchive(ByVal s As String):           mWsh.Value(Me, "FolderSyncArchive") = s:                        End Property

Public Property Let FolderSyncArchiveInfo(ByVal s As String):       MaintainMinRowHeight Me.Range("FolderSyncArchiveInfo"), s:      End Property

Public Property Let FolderSyncArchiveUrl(ByVal s As String):        UrlLabel(Me.Range("FolderSyncArchiveUrl")) = s:                 End Property

Public Property Get FolderSyncTarget() As String:                   FolderSyncTarget = mWsh.Value(Me, "FolderSyncTarget"):          End Property

Public Property Let FolderSyncTarget(ByVal s As String):            mWsh.Value(Me, "FolderSyncTarget") = s:                         End Property

Public Property Let FolderSyncTargetInfo(ByVal s As String):        MaintainMinRowHeight Me.Range("FolderSyncTargetInfo"), s:       End Property

Public Property Let FolderSyncTargetUrl(ByVal s As String):         UrlLabel(Me.Range("FolderSyncTargetUrl")) = s:                  End Property

Private Property Get MonitorAnchor() As Range:                      Set MonitorAnchor = Me.Range("ConfigInfo"):                     End Property

Public Property Get ButtonPauseAddIn() As Button:                    Set ButtonPauseAddIn = Me.Buttons("cmbAddInPause"):              End Property

Public Property Let UrlButton(ByVal url_rng As Range, _
                              ByVal url_chapter As String)
                        
    mWsh.Url(Me, url_rng, xlUnderlineStyleNone, 11) = Replace(Replace(mCompMan.README_URL & url_chapter, "##", "#"), "?", vbNullString)
    
End Property

Public Property Let UrlLabel(ByVal url_rng As Range, _
                             ByVal url_chapter As String)
                        
    mWsh.Url(Me, url_rng, xlUnderlineStyleNone, 9) = Replace(Replace(mCompMan.README_URL & url_chapter, "##", "#"), "?", vbNullString)
    
End Property

Public Sub AutoOpenCompManAddinSetup()
' ----------------------------------------------------------------------------
' Setup/replace the CompMan Add-in auto-open as a shortcut to the Add-in-
' Workbook in the users \AppData\Microsoft\Excel\XLSTART folder.
' ----------------------------------------------------------------------------
    Const PROC = "AutoOpenCompManAddinSetup"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim oShell      As IWshShell3
    Dim oShortcut   As IWshShortcut
    Dim sShortcut   As String
    
    If mMe.IsAddinInstnc Then GoTo xt
    sShortcut = mAddin.AutoOpenShortCut
    If fso.FileExists(sShortcut) Then fso.DeleteFile sShortcut
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

Public Sub AutoOpenCompManWrkbkSetup()
' ----------------------------------------------------------------------------
' Flips the Auto-Open status of the CompMan Workbook. When auto-open is setup
' the auto-open is removed, else it is setup.
' ----------------------------------------------------------------------------
    Const PROC = "AutoOpenCompManWrkbkSetup"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim oShell      As IWshShell3
    Dim oShortcut   As IWshShortcut
    
    Application.ScreenUpdating = False
    If mMe.IsAddinInstnc Then GoTo xt
    If fso.FileExists(AutoOpenShortCutCompManWbk) Then
        fso.DeleteFile AutoOpenShortCutCompManWbk
        Me.ButtonCaption(Me.ButtonWrkbkAutoOpen) = "Setup Auto-open"
    Else
        Set oShell = CreateObject("Wscript.shell")
        Set oShortcut = oShell.CreateShortcut(AutoOpenShortCutCompManWbk)
        With oShortcut
            .TargetPath = ThisWorkbook.FullName
            .Description = "AutoOpen CompMan.xlsb"
            .Save
        End With
        Me.ButtonCaption(Me.ButtonWrkbkAutoOpen) = "Remove Auto-open"
    End If
    wsConfig.CurrentStatus
    
xt: Set oShell = Nothing
    Set oShortcut = Nothing
    Exit Sub
 
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Property Let ButtonEnabled(Optional ByVal bc_cmb As Button, _
                                           ByVal bc_enabled As Boolean)
' ----------------------------------------------------------------------------
' Because the Enabled property does not what is expected the button is hidden
' or displayed instead.
' ----------------------------------------------------------------------------
    Dim bProtected  As Boolean
    Dim wsh         As Worksheet
    
    Set wsh = bc_cmb.Parent
    If wsh.ProtectContents Then
        bProtected = True
        wsh.Unprotect
    End If
                                           
    With bc_cmb
        If bc_enabled Then
            .Visible = True
            .Font.Color = rgbBlack
        Else
            .Visible = False
            .Font.Color = rgbGrey
        End If
    End With
    
    If bProtected Then wsh.Protect

End Property

Public Property Let ButtonCaption(Optional ByVal bc_cmb As Button, _
                                           ByVal bc_caption As String)
    Dim bProtected  As Boolean
    Dim wsh         As Worksheet
    
    Set wsh = bc_cmb.Parent
    If wsh.ProtectContents Then
        bProtected = True
        wsh.Unprotect
    End If
    
    bc_cmb.Text = bc_caption
    If bProtected Then wsh.Protect

End Property

Public Sub CurrentStatus(Optional ByVal cs_valid = True)
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
    
    If mMe.IsAddinInstnc Then GoTo xt
    
    With Me
        '~~ CompMan Add-in AutoOpen status
        sShortcut = mAddin.AutoOpenShortCut
        If fso.FileExists(sShortcut) Then
            '~~ Update the shortcut in case the location of the Add-in Workbook has changed
            If ShortCutTargetPath(sShortcut) <> mAddin.WbkFullName _
            Then ShortCutTargetPath(sShortcut) = mAddin.WbkFullName
            cll.Add "auto-open setup"
        Else
            cll.Add "auto-open not setup"
        End If
        
        '~~ CompMan AddIn Paused status
        If mAddin.IsOpen Then
            .ButtonPauseAddIn.Visible = True
            If mAddin.Paused Then
                cll.Add "curretly  p a u s e d !"
                Me.ButtonCaption(.ButtonPauseAddIn) = "Continue Add-in"
            Else
                Me.ButtonCaption(.ButtonPauseAddIn) = "Pause Add-in"
            End If
            cll.Add "open"
        Else
            cll.Add "not open"
            .ButtonPauseAddIn.Visible = False
        End If
        
        '~~ CompMan Add-in setup status
        If mAddin.Folder <> vbNullString Then
            If fso.FileExists(mAddin.WbkFullName) Then
                cll.Add "set up"
            Else
                cll.Add "not set up"
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

        '~~ CompMan Add-in folder configured status
        If mAddin.Folder <> vbNullString Then
            cll.Add "configured"
            .ButtonEnabled(Me.ButtonSetupRenewAddIn) = True
        Else
            cll.Add "not configured"
            .ButtonEnabled(Me.ButtonSetupRenewAddIn) = False
        End If
        
        sStatus = cll(cll.Count)
        For i = cll.Count - 1 To 1 Step -1
            sStatus = sStatus & ", " & cll(i)
        Next i
        .AddInStatus = sStatus
        
        '~~ CompMan Workbook AutoOpen status
        sStatus = vbNullString
        If fso.FileExists(AutoOpenShortCutCompManWbk) Then
            '~~ Update the shortcut in case the location of the CompMan Workbook has changed
            If ShortCutTargetPath(AutoOpenShortCutCompManWbk) <> ThisWorkbook.FullName _
            Then ShortCutTargetPath(AutoOpenShortCutCompManWbk) = ThisWorkbook.FullName
            .AutoOpenStatus = "auto-open for " & ThisWorkbook.name & "  s e t u p."
        Else
            .AutoOpenStatus = "auto-open for " & ThisWorkbook.name & "  n o t  s e t u p !"
        End If
        .ButtonHelpUrlAddinSetup = "#setuprenew-add-in"
        .ButtonHelpUrlPauseContinue = "#pausecontinue-add-in"
    End With

xt: Set fso = Nothing
    Set cll = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Enabled(ByVal en_cmb As CommandButton, _
          Optional ByVal en_enabled As Boolean = True)
    With en_cmb
        .Enabled = en_enabled
    End With
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wsConfig." & sProc
End Function

Public Sub FolderAddInClear()
' ----------------------------------------------------------------------------
' Clears all about the Add-in:
' - Removes an existing CompMan Add-in from the Add-in folder
' - Sets the configured Add-in folder to vbNullString
'
' Note: The Add-in folder itself is not removed and thus may be re-configured
'       at any point in time
' ----------------------------------------------------------------------------
    mAddin.Folder = vbNullString
    Me.FolderAddInIsValid
End Sub

Public Sub FolderAddinConfigure()
    
    With Me
        '~~ Save the current configured folder in case it is re- or de-configured
        If mAddin.Folder <> vbNullString Then mAddin.FolderOld = mAddin.Folder
        '~~ Re- or de-configure
        mAddin.Folder = mBasic.SelectFolder("Select the folder for the 'CompMan Add-in-Instance'")
        If mAddin.Folder = vbNullString Or Not .FolderAddInIsValid Then
            If mAddin.FolderOld <> mAddin.Folder Then
                '~~ The Add-in folder had been re- or de-configured
                mAddin.Clear mAddin.FolderOld
            End If
        End If
    End With
    If mAddin.Folder = vbNullString Then
        CurrentStatus mAddin.FolderOld
    Else
        CurrentStatus mAddin.Folder
    End If
    
End Sub

Public Function FolderAddInIsValid() As Boolean
' ----------------------------------------------------------------------------
' - Returns TRUE when the AddIn folder is configured and valid (i.e. exists)
' - Sets status and folder info in accordance with the current status.
' ----------------------------------------------------------------------------
    Const PROC = "FolderAddInIsValid"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim bValid  As Boolean
    Dim sStatus As String
    
    With Me
        .FolderAddInUrl = "#the-add-in-folder"
        .ButtonHelpUrlAddinSetup = "#setup-auto-open-for-the-compman-workbook"

        Select Case True
            Case mAddin.Folder = vbNullString
                .FolderAddInInfo = "Not configured! A CompMan Add-in instance cannot be setup " & _
                                   "(an available Add-in will have been removed)."
                bValid = False
            Case mAddin.Folder <> vbNullString
                If fso.FolderExists(mAddin.Folder) Then
                    .FolderAddInInfo = "Configured folder for the CompMan Add-in, possibly setup/renewed." & vbLf & _
                                       "Attention: In case a different folder is selected/established the previous folder becomes obsolete but will not be removed!"
                    bValid = True
                    bValid = True
                Else
                    .FolderAddInInfo = "Invalid! (the folder does not exist)"
                    bValid = False
                End If
            Case mAddin.Folder = vbNullString
                .FolderAddInInfo = "Not configured! " & vbLf & _
                                   "Folder for the CompMan Add-in instance."
                bValid = False
        End Select
        CurrentStatus bValid
    
        If bValid _
        Then .ButtonSetupRenewAddIn.Visible = True _
        Else .ButtonSetupRenewAddIn.Visible = False
    End With
    
xt: FolderAddInIsValid = bValid
    Set fso = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub FolderDevAndTestClear()
    Debug.Print "FolderDevAndTestClear clicked!"
    Me.FolderDevAndTest = vbNullString
    Me.FolderDevAndTestIsValid
End Sub

Public Sub FolderDevAndTestConfigure()
    wsConfig.FolderDevAndTest = mBasic.SelectFolder("Select the 'Development and Test' folder serviced by CompMan")
End Sub

Public Function FolderDevAndTestIsValid() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current configured Synchronization-Archive-Folder is
' valid.
' ----------------------------------------------------------------------------
    Const PROC = "FolderDevAndTestIsValid"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim bValid  As Boolean
    
    With Me
        .FolderDevAndTestUrl = "#the-serviced-development-and-test-folder"
        .CompManWrkbkStatusUrl = "#setup-auto-open-for-the-compman-workbook"
        If .FolderDevAndTest <> vbNullString Then
            If fso.FolderExists(.FolderDevAndTest) Then
                .FolderDevAndTestInfo = "Folder serviced by CompMan. I.e. CompMan provides its services for Workbooks " & _
                                          "located in this folder provided they have one or more services enabled."
                If FolderExportIsValid Then .ConfigInfo = vbNullString
                bValid = True
            Else
                .FolderDevAndTestInfo = "Invalid! The configured folder does not exist!."
            End If
        Else
            .FolderDevAndTestInfo = "Missing! CompMan's services are  n o t   a v a i l a b l e !" & vbLf & _
                                      "CompMan only serves Workbooks which are opened from within a configured 'Serviced Development and Test' folder."
        End If
        .ButtonWrkbkAutoOpen.Visible = bValid
    End With
    
xt: FolderDevAndTestIsValid = bValid
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub FolderExportConfigure()
    
    Dim fso     As New FileSystemObject
    Dim sNew    As String
    Dim sOld    As String
    Dim fld     As Folder
    
    sOld = wsConfig.FolderExport
    sNew = VBA.InputBox(Prompt:="Enter the name for the folder to which changed components are exported " & _
                             "by CompMan's 'Export Changed Components' service (for Workbooks which a " & _
                             "correspondingly configured and reside within the configured serrviced 'Development and Test' folder." & vbLf & vbLf & _
                             "Attention! When the name of the Export folder is changed, all old named folders are renamed correspondingly." _
                    , Title:="Export folder name" _
                    , Default:=sOld)
    If sNew <> vbNullString Then
        wsConfig.FolderExport = sNew
        If sNew <> sOld And wsConfig.FolderDevAndTestIsValid Then
            '~~ Provided a 'Serviced Development And Test' folder is configured,
            '~~ all export folders are renamed in accordance with the name change
            RenameSubFolders wsConfig.FolderDevAndTest, sOld, sNew
        End If
    Else
        wsConfig.FolderExport = sOld
    End If
    
End Sub

Public Function FolderExportIsValid(Optional ByRef iv_info As String) As Boolean
' ----------------------------------------------------------------------------
' Verification the name of the Export-Folder
' ----------------------------------------------------------------------------
    With Me
        .FolderExportUrl = "#the-export-folder"
        If .FolderExport <> vbNullString Then
            .FolderExportInfo = "Current configured name of the folder within a serviced Workbook's dedicated folder " & _
                                "into which CompMan exports new/modified components. The name may be changed however."
            FolderExportIsValid = True
        Else
            .FolderExport = "source"
            .FolderExportInfo = "Current configured (default) name of the folder within a serviced Workbook's dedicated " & _
                                "folder into which CompMan exports new/modified components. The name may be changed however."
            FolderExportIsValid = False
        End If
        iv_info = .FolderExportInfo
    End With
End Function

Public Sub FolderSyncArchiveClear()
    wsConfig.FolderSyncArchive = vbNullString
End Sub

Public Sub FolderSyncArchiveConfigure()
    wsConfig.FolderSyncArchive = mBasic.SelectFolder("Select the 'Synchronization Archive' folder for CompMan's 'Synchronize VB-Project' service")
    Me.FolderSyncArchiveIsValid
    CurrentStatus
End Sub

Public Function FolderSyncArchiveIsValid(Optional ByRef iv_info As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current configured Synchronization-Archive-Folder is
' valid.
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    
    With Me
        .FolderSyncArchiveUrl = "#the-synchronization-archive-folder"
        If .FolderSyncArchive <> vbNullString Then
            FolderSyncArchiveIsValid = fso.FolderExists(.FolderSyncTarget)
        Else
            If .FolderSyncTarget <> vbNullString Then
                .FolderSyncArchiveInfo = "Invalid! When a Sync-Target-Folder is configured the configuration of a Sync-Archive-Folder is obligatory!"
            End If
        End If
        If FolderSyncArchiveIsValid Then
            .FolderSyncArchiveInfo = "The Sync-Target-Workbook will be copied into this folder befor it gets synchronized."
        End If
    End With
    
    Set fso = Nothing

End Function

Public Sub FolderSyncTargetClear()
    wsConfig.FolderSyncTarget = vbNullString
End Sub

Public Sub FolderSyncTargetConfigure()
    wsConfig.FolderSyncTarget = mBasic.SelectFolder("Select the folder for the CompMan's 'Synchronize VB-Project' service")
End Sub

Public Function FolderSyncTargetIsValid() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current configured Synchronization-Folder is valid.
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    
    With wsConfig
        .FolderSyncTargetUrl = "#the-serviced-synchronization-target-folder"
        If .FolderSyncTarget <> vbNullString Then
            If fso.FolderExists(.FolderSyncTarget) Then
                FolderSyncTargetIsValid = True
                .FolderSyncTargetInfo = "CompMan's folder for the 'Synchronization VB-Project' service. The service is " & _
                                        "provided for any Workbook prepared for the service and opened from within this folder."
            End If
        Else
            .FolderSyncTargetInfo = "No 'Synchronization-Target' folder configured. CompMan's 'Synchronize VB-Project' service" & _
                                    "will not be provided even when a Workbook has it enabled." & vbLf & _
                                    "Important! The folder must not be a sub-folder of the 'Serviced Development and Test' folder!"
            FolderSyncTargetIsValid = False
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

Private Function MayBeConfirmed() As Boolean
    With Me
        MayBeConfirmed = (bAddinConfigObligatory And FolderAddInIsValid) _
                      Or (mAddin.Folder <> vbNullString And FolderAddInIsValid) _
                      And .FolderSyncTargetIsValid _
                      And .FolderExportIsValid _
                      And (bSyncConfigObligatory And .FolderSyncTargetIsValid) Or (.FolderSyncTarget <> vbNullString And .FolderSyncTargetIsValid) _
                      And (bSyncConfigObligatory And FolderSyncArchiveIsValid) Or (.FolderSyncArchive <> vbNullString And FolderSyncArchiveIsValid)
    End With
End Function

Public Sub MonitorRenewStep(Optional ByVal rn_action As String = vbNullString, _
                            Optional ByVal rn_result As String = vbNullString, _
                            Optional ByVal rn_last_step As Boolean = False, _
                            Optional ByVal rn_renew_failed As Boolean = False)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    Dim Title           As String:      Title = "Establish or renew CompMan as Add-in"
    Dim Header          As TypeMsgText
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

Private Sub VerifyConfig()
    With Me
        .FolderDevAndTestIsValid
        .FolderAddInIsValid
        .FolderSyncTargetIsValid
        .FolderSyncArchiveIsValid
        .CompManAddInStatusUrl = "#status-of-the-compman-add-in"
        .CompManWrkbkStatusUrl = "#setup-auto-open-for-the-compman-workbook"
    End With
End Sub

Private Sub Worksheet_Activate()
    
    If mMe.IsAddinInstnc Then GoTo xt
    With Application
        .ScreenUpdating = False
        .DisplayFormulaBar = False
    End With
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayHeadings = False
    End With
    
    VerifyConfig
    
xt: Exit Sub

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    With Me
        Select Case Target.Address
            Case .Range("FolderAddInInfo").Address, _
                 .Range("FolderDevAndTestInfo").Address, _
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
            Case .Range("FolderDevAndTest").Address:    .FolderDevAndTestIsValid
            Case .Range("FolderAddIn").Address:         .FolderAddInIsValid
            Case .Range("FolderSyncTarget").Address:    .FolderSyncTargetIsValid
            Case .Range("FolderSyncArchive").Address:   .FolderSyncArchiveIsValid
        End Select
    End With
End Sub
