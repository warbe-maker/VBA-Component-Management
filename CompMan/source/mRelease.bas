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
    Const PROC = "ReleaseCommComps"
    
    mBasic.BoP ErrSrc(PROC)
    mEnvironment.Provide False
    mCompMan.ServiceInitiate s_serviced_wbk:=ActiveWorkbook _
                           , s_service:="Release Modified Common Components"
        
    Set CommCompsPendingRelease = CommonPending.Components
    If CommCompsPendingRelease.Count <> 0 Then
        ReleaseService
    Else
        If r_verbose Then
            VBA.MsgBox Prompt:="No Common Component pending release for the active Workbook (" & Serviced.Wrkbk.Name & ")" _
                     , Title:="CompMan's ""Release Common Components to public"" service"
        End If
    End If
    mBasic.EoP ErrSrc(PROC)
    
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
    Const BTTN_RELEASE      As String = "Release"
    Const BTTN_SKIP_FOR_NOW As String = "Skip release for now"
    Const BTTN_SKIP_FOREVER As String = "Skip release forever"
    Const BTTN_Terminate    As String = "Terminate"
    
    On Error GoTo eh
    Dim cllButtons      As Collection
    Dim lOtherPending   As Long
    Dim Msg             As mMsg.udtMsg
    Dim qPending        As New clsQ
    Dim sBttnRelease    As String
    Dim sComp           As String
    Dim sTitle          As String
    Dim v               As Variant
    
    mCompMan.CurrentServiceName = "Release pending Common Compoents"
    mCompMan.ServicedWrkbk = ActiveWorkbook
    mEnvironment.Provide True

    mBasic.BoP ErrSrc(PROC)
    mCompMan.ServiceInitiate s_serviced_wbk:=ActiveWorkbook _
                           , s_service:=PROC _
                           , s_do_housekeeping:=False
    For Each v In CommonPending.ReadyForRelease
        qPending.EnQueue v
    Next v
    
    lOtherPending = CommCompsPendingRelease.Count - qPending.Size
    
    Do While Not qPending.IsEmpty
        qPending.First sComp
        With CommonPending
            .CompName = sComp
            If .LastModBy(sComp) = mEnvironment.ThisComputersUser _
            Then sBttnRelease = BTTN_RELEASE _
            Else sBttnRelease = BTTN_RELEASE & " explicitely confirmed"
            If Not CommonPublic.Exists(sComp) _
            Then sTitle = "Initial release of the Common Component  " _
            Else sTitle = "Release modifications/changes of the Common Component  "
            sTitle = sTitle & mBasic.Spaced(sComp)
            
            With Msg
                With .Section(1)
                    .Label.Text = "Last modified:"
                    .Label.FontColor = rgbBlue
                    .Text.Text = "in: " & CommonPending.LastModInWrkbkName(sComp) & vbLf & _
                                 "at: " & CommonPending.LastModAt(sComp) & vbLf & _
                                 "by: " & CommonPending.LastModBy(sComp) & vbLf & _
                                 "on: " & CommonPending.LastModOn(sComp)
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
                    .Label.Text = BTTN_Terminate & ":"
                    .Label.FontColor = rgbBlue
                    .Text.Text = "Terminates releasing any pending Common Components' modification in any Workbook"
                End With
                With .Section(5)
                    .Label.Text = sBttnRelease & ":"
                    .Label.FontColor = rgbBlue
                    If CommonPending.LastModBy(sComp) = mEnvironment.ThisComputersUser Then
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
            If CommonPublic.Exists(sComp) Then Set cllButtons = mMsg.Buttons(cllButtons, mDiff.PublicVersusPendingReleaseBttn(sComp))
            If ActiveWorkbook.Name = Serviced.Wrkbk.Name Then
                Set cllButtons = mMsg.Buttons(cllButtons, vbLf, BTTN_SKIP_FOR_NOW, BTTN_SKIP_FOREVER, vbLf, BTTN_Terminate)
            Else
                Set cllButtons = mMsg.Buttons(cllButtons, vbLf, BTTN_SKIP_FOR_NOW, vbLf, BTTN_Terminate)
            End If
            
            Select Case mMsg.Dsply(d_title:=sTitle _
                                 , d_msg:=Msg _
                                 , d_label_spec:="R100" _
                                 , d_buttons:=cllButtons _
                                 , d_width_min:=40)
                Case sBttnRelease:      CommonPending.ReleaseComp sComp, True
                                        CommonPending.Remove sComp
                                        qPending.DeQueue
                
                Case mDiff.PublicVersusPendingReleaseBttn(sComp):   mDiff.PublicVersusPendingReleaseDsply sComp
                
                Case BTTN_SKIP_FOR_NOW:                             qPending.DeQueue
                
                '~~ This option is only available when the ActiveWorkbook is the current serviced Workbook
                Case BTTN_SKIP_FOREVER:                             CommonServiced.KindOfComponent(sComp) = enCompCommonPrivate
                                                                    CommCompsPendingRelease.Remove sComp
                                                                    If .Exists(sComp) _
                                                                    Then .Remove sComp
                                                                    qPending.DeQueue
                
                Case BTTN_Terminate:                                Exit Do
            End Select
            mCompManMenuVBE.Setup
            If qPending.IsEmpty Then Exit Do
        End With
    Loop
    Application.StatusBar = " "
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRelease." & sProc
End Function

