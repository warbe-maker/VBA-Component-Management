Attribute VB_Name = "mRelease"
Option Explicit
Public Module
' ------------------------------------------------------------------------------
' Standard  Module mRelease: Provides the "Release new version" services for
' ========================== modified Common Components where the release is
' still pending.
'
' ------------------------------------------------------------------------------
Private Const README_PENDING_RELEASE    As String = "#pending-release-management"

Public Sub HelpRelease()
    mCompMan.README r_url_bookmark:=README_PENDING_RELEASE
End Sub

Public Sub ReleaseCommComps(Optional ByVal r_verbose As Boolean = True)
' ------------------------------------------------------------------------------
' Release any pending Common Component modification to public by moving the
' Export-File from the CommonPending folder into the Common-Components folder.
' ------------------------------------------------------------------------------
    
    mEnvironment.Provide ActiveWorkbook
    mCompMan.ServiceInitiate s_serviced_wbk:=ActiveWorkbook _
                           , s_service:="Release Modified Common Components"
        
    Set CommCompsPendingRelease = CommonPending.Components
    If CommCompsPendingRelease.Count <> 0 Then
        ReleaseService
    Else
        If r_verbose Then
            VBA.MsgBox Prompt:="No Common Component pending release for the active Workbook (" & Serviced.Wrkbk.Name & ")" _
                     , Title:="CompMan Common Component Release Service"
        End If
    End If
End Sub

Public Sub ReleaseService()
' ------------------------------------------------------------------------------
' Presents a dialog for the release of all pending Release candidates,
' each with the options:
' - Release used/hosted Common Component xxxx
' - Display changes
' - Skip for this time
' - Skip forever (comp becomes private)
' - Terminate
' Note: The service considers a Common Component may be released for the very
'       first time.
' ------------------------------------------------------------------------------
    Const PROC              As String = "ReleaseService"
    Const BTTN_DSPLY_DIFF   As String = "Display changes/modifications" & vbLf & vbLf & "(close window when done to continue)"
    Const BTTN_RELEASE      As String = "Release"
    Const BTTN_SKIP_FOR_NOW As String = "Skip release for now"
    Const BTTN_SKIP_FOREVER As String = "Skip release forever"
    Const BTTN_TERMINATE    As String = "Terminate"
    
    On Error GoTo eh
    Dim cllButtons      As Collection
    Dim CodePnding      As clsCode
    Dim CodePublic      As clsCode
    Dim Comp            As clsComp
    Dim i               As Long
    Dim lOtherPending   As Long
    Dim Msg             As mMsg.udtMsg
    Dim qPending        As New clsQ
    Dim sBttnRelease    As String
    Dim sComp           As String
    Dim sTitle          As String
    Dim v               As Variant
    
    mCompMan.CurrentServiceName = "Release pending Common Compoents"
    mCompMan.ServicedWrkbk = ActiveWorkbook
    mEnvironment.Provide ActiveWorkbook

    mBasic.BoP ErrSrc(PROC)
    mCompMan.ServiceInitiate s_serviced_wbk:=ActiveWorkbook _
                           , s_service:=PROC _
                           , s_do_housekeeping:=False
    For Each v In CommonPending.CommonComponentsPendingReadyForRelease
        Set Comp = New clsComp
        Comp.CompName = v
        qPending.EnQueue Comp
    Next v
    
    lOtherPending = CommCompsPendingRelease.Count - qPending.Size
    
    Do While Not qPending.IsEmpty
        qPending.First Comp
        sComp = Comp.CompName
        CommonPending.CompName = sComp
        If Comp.PendingLastModBy = mEnvironment.ThisComputersUser _
        Then sBttnRelease = BTTN_RELEASE _
        Else sBttnRelease = BTTN_RELEASE & " explicitely confirmed"
        If Comp.CodePublic.IsNone _
        Then sTitle = "Initial release of the Common Component  " _
        Else sTitle = "Release modifications/changes of the Common Component  "
        sTitle = sTitle & mBasic.Spaced(sComp)
        
        With Msg
            With .Section(1)
                .Label.Text = "Last modified:"
                .Label.FontColor = rgbBlue
                .Text.Text = "in: " & Comp.PendingLastModIn & vbLf & _
                             "at: " & Comp.PendingLastModAt & vbLf & _
                             "by: " & Comp.PendingLastModBy & vbLf & _
                             "on: " & Comp.PendingLastModOn
                .Text.MonoSpaced = True
            End With
            With .Section(2)
                .Label.Text = BTTN_SKIP_FOR_NOW & ":"
                .Label.FontColor = rgbBlue
                .Text.Text = "This dialog will continue with the next Common Component of which modifications are pending release (if any)."
            End With
            With .Section(3)
                .Label.Text = BTTN_SKIP_FOREVER & ":"
                .Label.FontColor = rgbRed
                .Text.Text = "Attention! This will change the Common Component to ""private""!" & vbLf & _
                             "I.e. the current and any subsequent modifications will no longer ever become public for other VB-Projects " & _
                             "since the component has become private, just by chance with the same name as a Common Component."
            End With
            With .Section(4)
                .Label.Text = BTTN_TERMINATE & ":"
                .Label.FontColor = rgbBlue
                .Text.Text = "Terminates releasing any pending Common Components' modification in any Workbook"
            End With
            With .Section(5)
                .Label.Text = sBttnRelease & ":"
                .Label.FontColor = rgbBlue
                If Comp.PendingLastModBy = mEnvironment.ThisComputersUser Then
                    .Text.Text = "Release will move the Export-File from the PendingReleases folder into the Common-Components folder."
                Else
                    .Text.Text = "Because the current logged-in user is not identical with the user who registered the modified " & _
                                 "Common Component ""pending release"", the release needs to be explicitely confirmed!"
                End If
            
            End With
            If lOtherPending > 0 Then
                With .Section(5)
                    .Label.Text = "Please note!"
                    .Label.FontColor = rgbBlue
                    If lOtherPending = 1 _
                    Then .Text.Text = "There is still another Common Component's modification pending release. " _
                    Else .Text.Text = "There are still " & lOtherPending & " Common Components' modification pending release. "
                    .Text.Text = .Text.Text & "But pending Common Component modifications can only be released by/from/within the " & _
                                              "Workbook in which it had been modified."
                End With
            End If
        End With
                
        Set cllButtons = mMsg.Buttons(sBttnRelease)
        If Not Comp.CodePublic.IsNone Then Set cllButtons = mMsg.Buttons(cllButtons, BTTN_DSPLY_DIFF)
        Set cllButtons = mMsg.Buttons(cllButtons, vbLf, BTTN_SKIP_FOR_NOW, BTTN_SKIP_FOREVER, vbLf, BTTN_TERMINATE)
        
        Select Case mMsg.Dsply(dsply_title:=sTitle _
                             , dsply_msg:=Msg _
                             , dsply_Label_spec:="R150" _
                             , dsply_buttons:=cllButtons _
                             , dsply_width_min:=60)
            Case sBttnRelease:      CommonPending.ReleaseComp sComp, True
                                    CommonPending.Remove sComp
                                    qPending.DeQueue
            Case BTTN_DSPLY_DIFF:   Comp.CodePublic.DsplyDiffs d_this_file_name:="CurrentPublic" _
                                                             , d_this_file_title:="The Common Component's current public code in the Common-Components folder's Export-File" _
                                                             , d_versus_code:=Comp.CodePnding _
                                                             , d_versus_file_name:="PendingReleaseModifications" _
                                                             , d_versus_file_title:="The Common Component's code recently modified in Workbook " & FSo.GetFileName(Comp.PendingLastModIn)

            Case BTTN_SKIP_FOR_NOW: qPending.DeQueue

            Case BTTN_SKIP_FOREVER: CommonServiced.KindOfComponent(sComp) = enCompCommonPrivate
                                    CommCompsPendingRelease.Remove 1
                                    If CommonPending.Exists(sComp) _
                                    Then CommonPending.Remove sComp
                                    qPending.DeQueue
            Case BTTN_TERMINATE:    Exit Do
        End Select
        mCompManMenuVBE.Setup
        Set CodePublic = Nothing
        Set CodePnding = Nothing
        If qPending.IsEmpty Then Exit Do
    Loop
    Application.StatusBar = " "
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function IsPendingModificationByThisServicedWorkbook(ByVal i_comp_name As String, _
                                                             ByRef i_comp As clsComp) As Boolean
    Dim Comp As clsComp
    
    With CommonPending
        If .LastModInWrkbkFullName(i_comp_name) = Serviced.Wrkbk.FullName Then
            Set Comp = New clsComp
            With Comp
                .CompName = i_comp_name
                If .CodeCrrent.Meets(.CodePnding) Then
                    IsPendingModificationByThisServicedWorkbook = True
                    Set i_comp = Comp
                    Set Comp = Nothing
                    GoTo xt
                End If
            End With
        End If
    End With
    
xt: Exit Function
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRelease." & sProc
End Function

