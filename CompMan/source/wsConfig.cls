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
' Data Module wsConfig: Current configuration throughout all services.
' =====================
' When CompMan is initially setup (after a download) the configuration of the
' Export-Folder name is the default and as such written to the private
' configuration PrivateProfile file (CompMan.cfg). Once this file exists,
' any user's private configuration is written to this configuration Worksheet
' when a new version is downloaded replacing the existing version.
' ----------------------------------------------------------------------------
Private sRenewResult As String

Private Property Let AddInStatus(ByVal s As String):                MaintainMinRowHeight Me.Range("AddInStatus"), s:                        End Property

Private Property Let AutoOpenStatus(ByVal s As String):             MaintainMinRowHeight Me.Range("AutoOpenStatus"), s:                     End Property

Private Property Get ButtonAddInProvideGiveUp() As Button:          Set ButtonAddInProvideGiveUp = Me.Buttons("cmbAddInSetUpGiveUp"):       End Property

Private Property Get ButtonAddInRenew() As Button:                  Set ButtonAddInRenew = Me.Buttons("cmbAddInRenew"):                     End Property

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

Private Property Let ButtonHelpUrlAddinSetup(ByVal s As String):    UrlButton(Me.Range("ButtonHelpUrlAddinSetup")) = s:                     End Property

Private Property Let ButtonHelpUrlPauseContinue(ByVal s As String): UrlButton(Me.Range("ButtonHelpUrlPauseContinue")) = s:                  End Property

Private Property Get ButtonPauseAddIn() As Button:                  Set ButtonPauseAddIn = Me.Buttons("cmbAddInPause"):                     End Property

Private Property Get ButtonSetupRenewAddIn() As Button:             Set ButtonSetupRenewAddIn = Me.Buttons("cmbAddInRenew"):                End Property

Private Property Get ButtonWrkbkAutoOpen() As Button:               Set ButtonWrkbkAutoOpen = Me.Buttons("cmbWrkbkAutoOpen"):               End Property

Private Property Let CommonComponentsFolderUrl(ByVal s As String):  UrlLabel(Me.Range("CommonComponentsFolderUrl")) = s:                    End Property

Private Property Let CompManAddInStatusUrl(ByVal s As String):      UrlLabel(Me.Range("CompManAddInStatusUrl")) = s:                        End Property

Private Property Let CompManWrkbkStatusUrl(ByVal s As String):      UrlLabel(Me.Range("CompManWrkbkStatusUrl")) = s:                        End Property

Public Property Let ConfigInfo(ByVal s As String):                  mWsh.Value(Me, "ConfigInfo") = s:                                       End Property

Public Property Get FolderCommonComponentsName() As String
    FolderCommonComponentsName = fso.GetFolder(FolderCommonComponentsPath).Name
End Property

Public Property Get FolderCommonComponentsPath() As String:         FolderCommonComponentsPath = mWsh.Value(Me, "FolderCommonComponents"):  End Property

Public Property Let FolderCommonComponentsPath(ByVal s As String):  mWsh.Value(Me, "FolderCommonComponents") = s:                           End Property

Public Property Get FolderCompManServicedRoot() As String:          FolderCompManServicedRoot = mWsh.Value(Me, "FolderCompManRoot"):        End Property

Public Property Let FolderCompManServicedRoot(ByVal s As String):   mWsh.Value(Me, "FolderCompManRoot") = s:                                End Property

Private Property Let FolderCompManRootInfo(ByVal s As String):      MaintainMinRowHeight Me.Range("FolderCompManRootInfo"), s:              End Property

Private Property Let FolderCompManRootUrl(ByVal s As String):       UrlLabel(Me.Range("FolderCompManRootUrl")) = s:                         End Property

Public Property Get FolderExport() As String:                       FolderExport = mWsh.Value(Me, "FolderExport"):                          End Property

Public Property Let FolderExport(ByVal s As String):                mWsh.Value(Me, "FolderExport") = s:                                     End Property

Private Property Get FolderExportInfo() As String:                  FolderExportInfo = mWsh.Value(Me, "FolderExportInfo"):                  End Property

Private Property Let FolderExportInfo(ByVal s As String):           MaintainMinRowHeight Me.Range("FolderExportInfo"), s:                   End Property

Private Property Let FolderExportUrl(ByVal s As String):            UrlLabel(Me.Range("FolderExportUrl")) = s:                              End Property

Private Property Get MonitorAnchor() As Range:                      Set MonitorAnchor = Me.Range("ConfigInfo"):                             End Property

Private Property Let UrlButton(ByVal url_rng As Range, _
                               ByVal url_chapter As String)
                        
    mWsh.Url(Me, url_rng, xlUnderlineStyleNone, 11) = Replace(Replace(mCompMan.GITHUB_REPO_URL & url_chapter, "##", "#"), "?", vbNullString)
    
End Property

Private Property Let UrlLabel(ByVal url_rng As Range, _
                              ByVal url_chapter As String)
                        
    mWsh.Url(Me, url_rng, xlUnderlineStyleNone, 9) = Replace(Replace(mCompMan.GITHUB_REPO_URL & url_chapter, "##", "#"), "?", vbNullString)
    
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
    CurrentStatus
    
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

Public Sub CurrentStatus()
' ----------------------------------------------------------------------------
' Analysis the status of the CompMan Add-in and the CompMan Workbook.
' ----------------------------------------------------------------------------
    Const PROC = "CurrentStatus"
    
    On Error GoTo eh
    Dim sStatus     As String
    Dim cll         As New Collection
    Dim i           As Long
    Dim sShortcut   As String
    Dim sComma      As String
    
    If mMe.IsAddinInstnc Then GoTo xt
    
    '~~ CompMan Add-in setup status
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
                cll.Add "open"
            End If
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
    ButtonHelpUrlPauseContinue = "#compmanxlsb-versus-compman-as-add-in"

xt: Set cll = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Me.CodeName & "." & sProc
End Function

Public Function FolderCompManRootIsValid() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current configured Synchronization-Archive-Folder is
' valid. Considers that the configured folder does not or no longer exist, in
' case the Workbook is opened on a different computer for example.
' ----------------------------------------------------------------------------
    Const PROC = "FolderCompManRootIsValid"
    
    On Error GoTo eh
    Dim bValid  As Boolean
    
    FolderCompManRootUrl = "/blob/master/SpecsAndUse.md#configuration-changes-compmans-config-worksheet"
    CompManWrkbkStatusUrl = "/blob/master/SpecsAndUse.md#configuration-changes-compmans-config-worksheet"
    CommonComponentsFolderUrl = "/blob/master/SpecsAndUse.md#public-common-components"
    
    If FolderCompManServicedRoot <> vbNullString Then
        If fso.FolderExists(FolderCompManServicedRoot) Then
            FolderCompManRootInfo = "Folder serviced by CompMan. I.e. CompMan provides its services for any Workbook " & _
                                    "located (opened, saved) in this folder - provided the Workbook has one or more " & _
                                    "services enabled."
            If FolderExportIsValid Then ConfigInfo = vbNullString
            bValid = True
        Else
            FolderCompManRootInfo = "Not configured! The former configured folder does not exist and thus has been reset!." & vbLf & _
                                    "CompMan 's services are  n o t   a v a i l a b l e !" & vbLf & _
                                    "CompMan only serves Workbooks which are opened from within a configured 'Serviced Development and Test' folder."
            FolderCompManServicedRoot = vbNullString
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
    sNew = VBA.InputBox(Prompt:="Enter the desired name for the folder to which changed components are exported " & _
                                "by CompMan's 'Export service - for correspondigly configured Workbooks, i.e which " & _
                                "reside within the serviced root folder (""" & wsConfig.FolderCompManServicedRoot & """)." & vbLf & vbLf & _
                                "Please note! When the name of the Export folder is changed, any still in use old named " & _
                                "folder in a serviced Workbook will be renamed." _
                    , Title:="Export folder name" _
                    , Default:=sOld)
    If sNew <> vbNullString Then
        FolderExport = sNew
        If sNew <> sOld And FolderCompManRootIsValid Then
            '~~ Provided a 'Serviced Development And Test' folder is configured,
            '~~ all export folders are renamed in accordance with the name change
            RenameSubFolders wsConfig.FolderCompManServicedRoot, sOld, sNew
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
        FolderExportUrl = "/blob/master/SpecsAndUse.md#compmans-environment"
        If .FolderExport <> vbNullString Then
            FolderExportInfo = "Current configured name of the folder within a serviced Workbook's dedicated folder " & _
                               "into which CompMan exports new/modified components. The name may be changed only by " & _
                               "means of the corresponding ""Configure"" button exclusively!"
            FolderExportIsValid = True
        Else
            FolderExport = mConfig.DEFAULT_FOLDER_EXPORT
            FolderExportInfo = "Current configured (default) name of the folder within a serviced Workbook's dedicated " & _
                               "folder into which CompMan exports new/modified components. The name may be changed only by " & _
                               "means of the corresponding ""Configure"" button exclusively!"
            FolderExportIsValid = False
        End If
        iv_info = FolderExportInfo
    End With
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
    Dim Step            As mMsg.udtMsgText
    Dim Footer          As mMsg.udtMsgText
    Dim sProgressMsg    As String
    Dim lSecs           As Long
    
    If rn_result <> vbNullString Then sRenewResult = rn_result
    If sRenewResult <> vbNullString Then sProgressMsg = ProgressMsg(RenewStep, RenewAction, sRenewResult)
    Step.FontSize = 8
    With Footer
        .FontSize = 8
        .FontColor = rgbDarkGreen
        .FontBold = True
        .FontColor = rgbDarkGreen
    End With
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
             Footer.Text = "Renew Addin in progress. This may take a few seconds! Please wait and note " & _
                           "that this monitor message may vanish several times during the progress."
             mMsg.MonitorFooter Title, Footer
        Else
            mMsg.Monitor Title, Step
            If mMe.RenewTerminatedByUser Then
                Footer.Text = "Renew process terminated by user! Close window."
                mMsg.MonitorFooter Title, Footer
            Else
                If Not rn_renew_failed Then
                    For lSecs = 3 To 1 Step -1
                        Footer.Text = "Renew process  " & mBasic.Spaced("successfully") & "  finished! (Window closes in " & lSecs & " seconds)"
                        mMsg.MonitorFooter Title, Footer
                        mBasic.DelayedAction 1
                    Next lSecs
                    Servicing.MessageUnload Title
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
        ProgressMsg = mBasic.Align(pm_step, enAlignRight, 2) & " " & pm_result & " " & Split(pm_action, vbLf)(0)
        If UBound(Split(pm_action, vbLf)) > 0 Then
            ProgressMsg = ProgressMsg & vbLf & "          " & Split(pm_action, vbLf)(1)
        End If
    Else
        ProgressMsg = ProgressMsg & vbLf & mBasic.Align(pm_step, enAlignRight, 2) & " " & pm_result & " " & Split(pm_action, vbLf)(0)
        If UBound(Split(RenewAction, vbLf)) > 0 Then
            ProgressMsg = ProgressMsg & vbLf & "   " & Split(pm_action, vbLf)(1)
        End If
    End If
End Function

Public Function Verified() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when all configured folders are valid.
' ----------------------------------------------------------------------------
    Verified = FolderCompManRootIsValid
    CompManAddInStatusUrl = "/blob/master/SpecsAndUse.md#configuration-changes-compmans-config-worksheet"
    CompManWrkbkStatusUrl = "/blob/master/SpecsAndUse.md#configuration-changes-compmans-config-worksheet"
    CommonComponentsFolderUrl = "/blob/master/SpecsAndUse.md#public-common-components"
    
End Function

Private Sub Worksheet_Activate()
    Const PROC = "Worksheet_Activate"
    
    Dim b As Boolean
    
    
    mCompManClient.Events ErrSrc(PROC), False
    If mMe.IsAddinInstnc Then GoTo xt
    CurrentStatus
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
    
    With Me
        If .ProtectContents Then
            b = True
            .Unprotect
            .Range("RngConfig").Select
            ActiveWindow.Zoom = True
        End If
        With .Range("celParking")
            .Locked = False
            .Select ' selects a non visible cell
        End With
        If b Then .Protect
    End With
    
xt: mCompManClient.Events ErrSrc(PROC), True
    Exit Sub

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    With Me
        Select Case Target.Address
            Case .Range("FolderCompManRootInfo").Address, _
                 .Range("FolderExportInfo").Address, _
                 .Range("ConfigInfo").Address
                If InStr(Target.Value, "nvalid!") <> 0 _
                Or InStr(Target.Value, "issing!") <> 0 _
                Or InStr(Target.Value, "not configured!") <> 0 Then
                    Target.Font.Color = 255
                Else
                    Target.Font.ColorIndex = xlAutomatic
                End If
            Case .Range("FolderCompManRoot").Address:   FolderCompManRootIsValid
        End Select
    
        Select Case Target.Address
            Case .Range("FolderCompManRoot").Address
            Case .Range("FolderExport").Address
                If ConfigLocal Is Nothing Then
                    mEnvironment.Provide ThisWorkbook
                    Set ConfigLocal = New clsConfigLocal
                End If
                ConfigLocal.ExportFolderName = ConfigLocal.ExportFolderName & ">" & Target.Value
        End Select
    End With

End Sub
