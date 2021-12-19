Attribute VB_Name = "mUpdate"
Option Explicit
' -----------------
'
' -----------------
Private sUpdateStayVerbose      As String
Private sBttnDsplyChanges       As String
Private sBttnSkipStayVerbose    As String


Private Property Let BttnUpdateStayVerbose(ByVal comp_name As String)
    sUpdateStayVerbose = "Update" & vbLf & vbLf & Spaced(comp_name) & vbLf & vbLf & "(stay verbose)"
End Property

Private Property Get BttnUpdateStayVerbose() As String
    BttnUpdateStayVerbose = sUpdateStayVerbose
End Property

Private Property Get BttnUpdateAll() As String
    BttnUpdateAll = "Update" & vbLf & "all"
End Property

Private Property Let BttnDsplyChanges(ByVal comp_name As String)
    sBttnDsplyChanges = "Display code changes" & vbLf & vbLf & Spaced(comp_name)
End Property
Private Property Get BttnDsplyChanges() As String
    BttnDsplyChanges = sBttnDsplyChanges
End Property

Private Property Let BttnSkipStayVerbose(ByVal comp_name As String)
    sBttnSkipStayVerbose = "Skip this component" & vbLf & "(stay verbose)"
End Property
Private Property Get BttnSkipStayVerbose() As String
    BttnSkipStayVerbose = sBttnSkipStayVerbose
End Property

Private Property Get BttnSkipAll() As String
    BttnSkipAll = "Skip" & vbLf & "all"
End Property

Public Sub RawClones(ByRef urc_wb As Workbook)
' --------------------------------------------------------
' Updates any clone in Workbook (urc_wb). Note that clones
' are identifiied by equally named components in another
' workbook which are indicated 'hosted'.
' --------------------------------------------------------
    Const PROC = "RawClones"
    
    On Error GoTo eh
    Dim dctClonesTheRawChanged  As Dictionary
    Dim vbc                     As VBComponent
    Dim fso                     As New FileSystemObject
    Dim sUpdated                As String
    Dim v                       As Variant
    Dim bVerbose                As Boolean
    Dim bSkip                   As Boolean
    Dim Clones                  As clsClones
    Dim RawComp                 As clsRaw
    Dim CloneComp               As clsComp
    
    Set Clones = New clsClones
    Set dctClonesTheRawChanged = Clones.RawChanged
    
    bVerbose = True
            
    For Each v In dctClonesTheRawChanged
        mCompMan.DsplyProgress p_result:=sUpdated _
                             , p_total:=Stats.Total(sic_clone_comps) _
                             , p_done:=Stats.Total(sic_clones_comps_updated)
        Set RawComp = dctClonesTheRawChanged(v)
        Set CloneComp = New clsComp
        Set CloneComp.Wrkbk = urc_wb
        CloneComp.CompName = RawComp.CompName
        Log.ServicedItem = CloneComp.VBComp
        Log.Entry = "Corresponding Raw-Component's code changed! (its Export-File differs from the Clone's Export-File)"
        With RawComp
            If bVerbose Then
                UpdateCloneConfirmed ucc_comp_name:=.CompName _
                                   , ucc_stay_verbose:=bVerbose _
                                   , ucc_skip:=bSkip _
                                   , ucc_clone:=CloneComp.ExpFileFullName _
                                   , ucc_raw:=.ExpFileFullName
            End If
            If Not bSkip Then
                mRenew.ByImport rn_wb:=urc_wb _
                     , rn_comp_name:=.CompName _
                     , rn_exp_file_full_name:=.ExpFileFullName
                Log.Entry = "Clone-Component renewed/updated by (re-)import of '" & RawComp.ExpFileFullName & "'"
                mRawClonesUsed.RevisionNumber = mRawsHosted.RevisionNumber
                Stats.Count sic_clones_comps_updated
                If sUpdated = vbNullString _
                Then sUpdated = .CompName _
                Else sUpdated = .CompName & ", " & sUpdated
            End If
        End With
        Set CloneComp = Nothing
        Set RawComp = Nothing
        
    Next v
    If sUpdated = vbNullString _
    Then Application.StatusBar = Log.Service & "None updated (" & Stats.Total(sic_clones_comps_updated) & " of " & Stats.Total(sic_clone_comps) & ")" _
    Else Application.StatusBar = Log.Service & sUpdated & " (" & Stats.Total(sic_clones_comps_updated) & " of " & Stats.Total(sic_clone_comps) & ")"

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function UpdateCloneConfirmed( _
                               ByVal ucc_comp_name As String, _
                               ByRef ucc_stay_verbose As Boolean, _
                               ByRef ucc_skip As Boolean, _
                               ByVal ucc_clone As String, _
                               ByVal ucc_raw As String)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    Const PROC = "UpdateCloneConfirmed"
    
    On Error GoTo eh
    Dim cllButtons      As Collection
    Dim sMsg            As TypeMsg
    Dim vReply          As Variant
    
    BttnDsplyChanges = ucc_comp_name
    BttnUpdateStayVerbose = ucc_comp_name
    mMsg.Buttons cllButtons, BttnDsplyChanges _
                     , vbLf, BttnUpdateStayVerbose, BttnSkipStayVerbose _
                     , vbLf, BttnUpdateAll, BttnSkipAll
    
    With sMsg
        .Section(1).Label.Text = "About this update:"
        .Section(1).Text.Text = "When the raw clone used in this Workbook is not updated the message will show up again " & _
                                "the next time this Workbook is opened - provided it is located/opened in the  in the " & _
                                "configured development root folder:"
        .Section(2).Text.Text = mMe.ServicedRootFolder
        .Section(2).Text.MonoSpaced = True
        With .Section(3)
            If mRawsHosted.RevisionNumber = mRawClonesUsed.RevisionNumber Then
                .Label.Text = "Attention!"
                .Label.FontColor = rgbRed
                .Text.Text = "It appears that the code of the raw clone used in this Workbook has been modified. This modification will be " & _
                             "reverted with this update. Displaying the difference will be the last chance to modify the raw component in its " & _
                             "hosting Workbook (" & mRawsHosted.HostFullName(ucc_comp_name) & ")."
            Else
                .Text.Text = vbNullString
            End If
        End With
    End With
    
    Do
        vReply = mMsg.Dsply(dsply_title:=Log.Service & "Update " & Spaced(ucc_comp_name) & "with changed raw" _
                          , dsply_msg:=sMsg _
                          , dsply_buttons:=cllButtons _
                           )
        Select Case vReply
            Case BttnDsplyChanges
                mFile.Compare fc_file_left:=ucc_clone _
                            , fc_left_title:="Cloned raw: '" & ucc_clone & "'" _
                            , fc_file_right:=ucc_raw _
                            , fc_right_title:="Current raw: '" & ucc_raw & "'"
            
            Case BttnUpdateStayVerbose
                ucc_skip = False
                Exit Do
            Case BttnSkipStayVerbose
                ucc_stay_verbose = True
                Exit Do
            Case BttnUpdateAll
                ucc_skip = False
                ucc_stay_verbose = False
                Exit Do
            Case BttnSkipAll
                ucc_skip = True
                Exit Do
        End Select
    Loop
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub DisplayStatus( _
                    ByVal sp_cloned_raws As Long, _
                    ByVal sp_no_replaced As Long, _
                    ByRef sp_replaced As String)
' -------------------------------------------------
' Display the services final result.
' -------------------------------------------------
    Dim sMsg        As String
    
    Select Case sp_cloned_raws
        Case 0: sMsg = "No 'Cloned Raw Component' in Workbook"
        Case 1
            Select Case sp_no_replaced
                Case 0:     sMsg = "1 Clone-Component identified but its corresponding Raw-Component had not changed."
                Case 1:     sMsg = "1 of 1 Clone-Component updated because its corresponding Raw-Component had changed (" & Left(sp_replaced, Len(sp_replaced) - 2) & ")."
            End Select
        Case Else
            Select Case sp_no_replaced
                Case 0:     sMsg = "No Clone-Component updated (0 of " & sp_cloned_raws & ")"
                Case 1:     sMsg = "1 of " & sp_cloned_raws & " Clone-Components has been updated because the corresponding Raw-Component has changed (" & Left(sp_replaced, Len(sp_replaced) - 2) & ")."
                Case Else:  sMsg = sp_no_replaced & " of " & sp_cloned_raws & " Clone-Components have been updated because their corresponding Raw-Component had changed (" & Left(sp_replaced, Len(sp_replaced) - 2) & ")."
            End Select
    End Select
    DsplyProgress sMsg

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

