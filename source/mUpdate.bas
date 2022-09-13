Attribute VB_Name = "mUpdate"
Option Explicit

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mUpdate" & "." & es_proc
End Function

Private Function Spaced(ByVal s As String) As String
    Dim i       As Long
    Dim sSpaced As String
    sSpaced = " "
    For i = 1 To Len(s)
        sSpaced = sSpaced & Mid(s, i, 1) & " "
    Next i
    Spaced = sSpaced
End Function

Public Function UpdateConfirmed(ByRef uo_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the update of the outdated Used Common Component (uo_comp)
' has been confirmed.
' ----------------------------------------------------------------------------
    Const PROC = "UpdateConfirmed"
    
    On Error GoTo eh
    Dim cllButtons          As Collection
    Dim sMsg                As TypeMsg
    Dim vReply              As Variant
    Dim sBttnDsplyChanges   As String
    Dim sBttnUpdate         As String
    Dim sBttnSkip           As String
    Dim sCompName           As String
    
    mBasic.BoP ErrSrc(PROC)
    sCompName = uo_comp.CompName
    
    sBttnDsplyChanges = "Display code changes" & vbLf & vbLf & mBasic.Spaced(sCompName)
    sBttnUpdate = "Update" & vbLf & vbLf & mBasic.Spaced(sCompName)
    sBttnSkip = "Skip this update"
    Set cllButtons = mMsg.Buttons(sBttnDsplyChanges _
                                , vbLf _
                                , sBttnUpdate _
                                , vbLf _
                                , sBttnSkip _
                                 )
    With sMsg
        If uo_comp.DueModificationWarning Then
            With .Section(1)
                .Label.Text = "Attention!"
                .Label.FontColor = rgbRed
                .Text.Text = "The code of the 'Used Common Component'   " & mBasic.Spaced(sCompName) & _
                             "    had been modified within this Workbook/VBProject. This modification will be " & _
                             "reverted with this update. Displaying the difference will be the last chance to " & _
                             "modify the 'Raw Common Component' in its hosting Workbook (" & mComCompRawsGlobal.RawHostWbFullName(sCompName) & ")."
                .Text.FontColor = rgbRed
            End With
        End If
        With .Section(2)
            .Label.Text = "About this update:"
            .Text.Text = "When the update of the outdated 'Used Common Component'  " & mBasic.Spaced(sCompName) & "  " & _
                         "is skipped the request will be displayed again the next time this Workbook is opened " & _
                         "(from within the configured 'Serviced Folder' " & mConfig.ServicedDevAndTestFolder & "."
        End With
    End With
    
    Do
        With uo_comp
            If Not mMsg.IsValidMsgButtonsArg(cllButtons) Then Stop
            vReply = mMsg.Dsply(dsply_title:=Log.Service & "Update " & Spaced(.CompName) & "with changed raw" _
                              , dsply_msg:=sMsg _
                              , dsply_buttons:=cllButtons _
                               )
            Select Case vReply
                Case sBttnUpdate:       UpdateConfirmed = True:  Exit Do
                Case sBttnSkip:         UpdateConfirmed = False: Exit Do
                Case sBttnDsplyChanges: mService.ExpFilesDiffDisplay _
                                            fd_exp_file_left_full_name:=.ExpFileFullName _
                                          , fd_exp_file_left_title:="Outdated Used Common Component: '" & .ExpFileFullName & "'" _
                                          , fd_exp_file_right_full_name:=.Raw.SavedExpFileFullName _
                                          , fd_exp_file_right_title:="Up-to-date Raw Common Component: '" & .Raw.SavedExpFileFullName & "'"
            End Select
        End With
    Loop
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

