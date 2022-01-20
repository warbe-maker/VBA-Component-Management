Attribute VB_Name = "mUpdate"
Option Explicit

Public Sub Outdated(ByRef urc_wb As Workbook)
' ----------------------------------------------------------------------------
' Updates any Common Component used in Workbook (urc_wb). Note that Common
' Components are identifiied by equally named components in another workbook
' which is indicated the 'host' Workbook.
' ----------------------------------------------------------------------------
    Const PROC = "Outdated"
    
    On Error GoTo eh
    Dim dctUpdateDue    As Dictionary
    Dim vbc             As VBComponent
    Dim fso             As New FileSystemObject
    Dim sUpdated        As String
    Dim v               As Variant
    Dim bVerbose        As Boolean
    Dim bSkip           As Boolean
    Dim Comp            As clsComp
    Dim Comps           As New clsComps
    Dim lAll            As Long
    Dim lRemaining      As Long
    Dim wb              As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wb = mService.Serviced
    
    Set dctUpdateDue = Comps.Outdated
    With wb.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        bVerbose = True
        
        For Each vbc In .VBComponents
    
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wb
                .CompName = vbc.name
                Set .VBComp = vbc
                Log.ServicedItem = vbc
                If .KindOfComp = enCommCompUsed Then
                    If .Outdated Then
                        Log.Entry = "Outdated Used Common Component! (Export-Files '" & .ExpFileFullName & "' differs from '" & .Raw.SavedExpFileFullName & "')"
                        If OutdatedUsedCommonComponentConfirmed(Comp) Then
                            mRenew.ByImport rn_wb:=urc_wb _
                                 , rn_comp_name:=.CompName _
                                 , rn_raw_exp_file_full_name:=.Raw.SavedExpFileFullName
                            Log.Entry = "Used Common Component renewed/updated by (re-)import of the Raw's Export File (" & .Raw.SavedExpFileFullName & ")"
                            .RevisionNumber = .Raw.RevisionNumber
                            Stats.Count sic_used_comm_comp_updated
                            If sUpdated = vbNullString _
                            Then sUpdated = .CompName _
                            Else sUpdated = .CompName & ", " & sUpdated
                        End If
                    End If ' .Outdated
                End If ' Used Common Component
            End With
            Set Comp = Nothing
            lRemaining = lRemaining - 1
            Application.StatusBar = _
            mService.Progress(p_service:=Log.Service _
                            , p_result:=Stats.Total(sic_used_comm_comp_updated) _
                            , p_of:=lAll _
                            , p_op:="updated" _
                            , p_comps:=sUpdated _
                            , p_dots:=lRemaining _
                             )
        Next vbc
    End With

    Application.StatusBar = vbNullString
    Application.StatusBar = _
    mService.Progress(p_service:=Log.Service _
                    , p_result:=Stats.Total(sic_used_comm_comp_updated) _
                    , p_of:=lAll _
                    , p_op:="updated" _
                    , p_comps:=sUpdated _
                     )


xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

Public Function OutdatedUsedCommonComponentConfirmed(ByRef uo_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the update of the outdated Used Common Component (uo_comp)
' has been confirmed
' ----------------------------------------------------------------------------
    Const PROC = "OutdatedUsedCommonComponentConfirmed"
    
    On Error GoTo eh
    Dim cllButtons          As Collection
    Dim sMsg                As TypeMsg
    Dim vReply              As Variant
    Dim sBttnDsplyChanges   As String
    Dim sBttnUpdate         As String
    Dim sBttnSkip           As String
    
    mBasic.BoP ErrSrc(PROC)
    sBttnDsplyChanges = "Display code changes" & vbLf & vbLf & mBasic.Spaced(uo_comp.CompName)
    sBttnUpdate = "Update" & vbLf & vbLf & mBasic.Spaced(uo_comp.CompName)
    sBttnSkip = "Skip this update"
    mMsg.Buttons cllButtons, sBttnDsplyChanges _
                     , vbLf, sBttnUpdate _
                     , vbLf, sBttnSkip
    
    With sMsg
        .Section(1).Label.Text = "About this update:"
        .Section(1).Text.Text = "When the outdated 'Used Common Component'  " & mBasic.Spaced(uo_comp.CompName) & "  in " & _
                                "this Workbook is not updated the message will show up again the next time this " & _
                                "Workbook is opened - from within the configured 'Serviced Folder'."
        .Section(2).Text.Text = mConfig.FolderServiced
        .Section(2).Text.MonoSpaced = True
        With .Section(3)
            If Not uo_comp.RevisionNumbersDiffer Then
                .Label.Text = "Attention!"
                .Label.FontColor = rgbRed
                .Text.Text = "It appears that the code of the raw clone used in this Workbook has been modified. This modification will be " & _
                             "reverted with this update. Displaying the difference will be the last chance to modify the raw component in its " & _
                             "hosting Workbook (" & mComCompsRawsSaved.RawHostWbFullName(uo_comp.CompName) & ")."
            Else
                .Text.Text = vbNullString
            End If
        End With
    End With
    
    Do
        With uo_comp
            vReply = mMsg.Dsply(dsply_title:=Log.Service & "Update " & Spaced(.CompName) & "with changed raw" _
                              , dsply_msg:=sMsg _
                              , dsply_buttons:=cllButtons _
                               )
            Select Case vReply
                Case sBttnUpdate:       OutdatedUsedCommonComponentConfirmed = True:  Exit Do
                Case sBttnSkip:         OutdatedUsedCommonComponentConfirmed = False: Exit Do
                Case sBttnDsplyChanges: mService.ExpFilesDiffDisplay _
                                            fd_exp_file_left_full_name:=.ExpFileFullName _
                                          , fd_exp_file_left_title:="Used Common Component: '" & .ExpFileFullName & "'" _
                                          , fd_exp_file_right_full_name:=.Raw.SavedExpFileFullName _
                                          , fd_exp_file_right_title:="Raw Common Component: '" & .Raw.SavedExpFileFullName & "'"
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

