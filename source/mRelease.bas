Attribute VB_Name = "mRelease"
Option Explicit
Public Module
' ------------------------------------------------------------------------------
' Standard  Module mRelease: Provides the "Release new version" services for
' ========================== Common Components.
'
' ------------------------------------------------------------------------------
Private cbcReleaseItem      As CommandBarControl
Private cbpMenu             As CommandBarPopup
Private dctPendingReleases  As New Dictionary

Public Sub HelpRelease()
    Debug.Print "HelpRelease performed"
'    mBasic.README r_url:=GITHUB_REPO_URL _
'                , r_bookmark:=README_RELEASE
End Sub

Public Sub ReleaseCommComps(Optional ByVal r_verbose As Boolean = True)
    Dim cllPending  As Collection
    Dim wbk         As Workbook
    
    Set wbk = ActiveWorkbook
    
    Set CommComps = New clsCommComps
    Set cllPending = CommComps.PendingReleases(wbk)
    If cllPending.Count <> 0 Then
        ReleaseService wbk, cllPending
    Else
        If r_verbose Then
            VBA.MsgBox Prompt:="No pending Common Component releases for the active Workbook (" & wbk.Name & ")" _
                     , Title:="CompMan Common Component Release Service"
        End If
    End If
End Sub

Public Sub ReleaseService(Optional ByVal r_wbk As Workbook = Nothing, _
                          Optional ByVal r_pending As Collection = Nothing)
' ------------------------------------------------------------------------------
' Presents a dialog for the release of all collected/pending Release candidates,
' each with the options:
' - Release used/hosted Common Component xxxx
' - Display changes
' - Skip for this time
' - Skip forever (comp becomes private)
' - Terminate
' When a Workbook is provided, the dialog presents only collected/pending
' Release candidates of this Workbook
' ------------------------------------------------------------------------------
    Const PROC              As String = "ReleaseService"
    Const BTTN_DSPLY_DIFF   As String = "Display changes/modifications" & vbLf & vbLf & "(close window when done to continue)"
    Const BTTN_RELEASE      As String = "Release"
    Const BTTN_SKIP_FOR_NOW As String = "Skip release for now"
    Const BTTN_SKIP_FOREVER As String = "Skip release forever"
    Const BTTN_TERMINATE    As String = "Terminate"
    
    On Error GoTo eh
    Dim sBttnRelease        As String
    Dim cllButtons          As Collection
    Dim Msg                 As mMsg.udtMsg
    Dim i                   As Long
    Dim sComp               As String
    Dim cllPending          As Collection
    Dim CodePublic          As clsCodeOfSource
    Dim CodePending         As clsCodeOfSource
    
    Set CommComps = New clsCommComps
    If r_pending Is Nothing _
    Then Set cllPending = CommComps.PendingReleases(r_wbk) _
    Else Set cllPending = r_pending
    
    Set LogService = New clsLog
    With LogService
        .WithTimeStamp
        .KeepDays = 10
        .FileFullName = ThisWorkbook.Path & "\" & fso.GetBaseName(ThisWorkbook.Name) & ".ServicesSummary.log"
        .NewLog True ' suppresses delimiter line
        mCompMan.LogFileServicedSummary = .FileName
    End With
    
    Do While cllPending.Count <> 0
        sComp = cllPending(1)
        CommComps.CompName = sComp
        Set CodePublic = New clsCodeOfSource
        Set CodePending = New clsCodeOfSource
        CodePublic.Source = CommComps.CurrentPublicExpFileFullName
        CodePending.Source = CommComps.PendingReleaseModExpFileFullName
        With Msg
            With .Section(1)
                If CodePublic.IsNone Then
                    .Label.Text = "Common Component:" & vbLf & "(initial release)"
                Else
                    .Label.Text = "Common Component:" & vbLf & "(pending release of modifications)"
                End If
                .Text.Text = CodePending.CompName
                .Text.FontBold = True
                .Text.MonoSpaced = True
            End With
            With .Section(2)
                .Label.Text = "Last modified in Workbook:"
                .Label.FontColor = rgbBlue
                .Text.Text = CommComps.PendingReleaseModInWbkName
                .Text.MonoSpaced = True
            End With
            With .Section(3)
                .Label.Text = "Last modified on Machine:"
                .Label.FontColor = rgbBlue
                .Text.Text = CommComps.PendingReleaseModOnMachine
                .Text.MonoSpaced = True
            End With
            With .Section(4)
                .Label.Text = "Last modified at Date-Time:"
                .Label.FontColor = rgbBlue
                .Text.Text = CommComps.PendingReleaseModAtDateTime
                .Text.MonoSpaced = True
            End With
            With .Section(5)
                .Label.Text = "Options:"
                .Label.FontBold = True
                .Text.Text = " "
            End With
            With .Section(6)
                .Label.Text = BTTN_SKIP_FOR_NOW & ":"
                .Label.FontColor = rgbBlue
                .Text.Text = "This dialog will continue with the next Common Component of which modifications are pending release (if any)."
            End With
            With .Section(7)
                .Label.Text = BTTN_SKIP_FOREVER & ":"
                .Label.FontColor = rgbRed
                .Text.Text = "Attention! This will declare the code modification made for  the Common Component ""private"", " & vbLf & _
                             "i.e. any modification made for the public Common Component will no longer update this component and no " & _
                             "modification will ever become public for others."
            End With
            With .Section(8)
                .Label.Text = BTTN_TERMINATE & ":"
                If r_wbk Is Nothing Then
                    .Text.Text = "Terminates the release process for any still pending release for any Workbook"
                Else
                    .Text.Text = "Terminates the release process for any still pending release for Workbook" & vbLf & _
                                 mBasic.Spaced(r_wbk.Name)
                End If
            End With
            
        End With
                
        Set cllButtons = New Collection
        cllButtons.Add BTTN_RELEASE:    cllButtons.Add vbLf
        If Not CodePublic.IsNone _
        Then cllButtons.Add BTTN_DSPLY_DIFF: cllButtons.Add vbLf
        cllButtons.Add BTTN_SKIP_FOR_NOW
        cllButtons.Add BTTN_SKIP_FOREVER: cllButtons.Add vbLf
        cllButtons.Add BTTN_TERMINATE
        
        Select Case mMsg.Dsply(dsply_title:="Release the modifications/changes of a Common Component" _
                             , dsply_msg:=Msg _
                             , dsply_Label_spec:="R150" _
                             , dsply_buttons:=cllButtons _
                             , dsply_width_min:=50)
            Case BTTN_RELEASE:      CommComps.ReleaseComp sComp
                                    cllPending.Remove 1
            Case BTTN_DSPLY_DIFF:   CodePublic.DiffFromDsply d_this_file_name:="CurrentPublic" _
                                                           , d_this_file_title:="Code in the current ""public"" Common Component in the Common-Components folder" _
                                                           , d_from_code:=CodePending _
                                                           , d_from_file_name:="PendingReleaseModifications" _
                                                           , d_from_file_title:="The Common Component's code recently modified in Workbook " & CommComps.PendingReleaseModInWbkName

            Case BTTN_SKIP_FOR_NOW: cllPending.Remove 1
            Case BTTN_SKIP_FOREVER: CompManDat.RegistrationState(sComp) = enRegStatePrivate
                                    cllPending.Remove 1
                                    If CommComps.PendingReleaseRegistered Then
                                        CommComps.PendingReleaseRemove
                                    End If
            Case BTTN_TERMINATE:    Exit Do
        End Select
        Set CodePublic = Nothing
        Set CodePending = Nothing
    Loop
    mCompManMenu.Setup
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRelease." & sProc
End Function

