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
    sBttnDsplyChanges = "Display changes for" & vbLf & vbLf & Spaced(comp_name)
End Property
Private Property Get BttnDsplyChanges() As String
    BttnDsplyChanges = sBttnDsplyChanges
End Property

Private Property Let BttnSkipStayVerbose(ByVal comp_name As String)
    sBttnSkipStayVerbose = "Skip update of " & vbLf & vbLf & Spaced(comp_name) & vbLf & vbLf & "(stay verbose)"
End Property
Private Property Get BttnSkipStayVerbose() As String
    BttnSkipStayVerbose = sBttnSkipStayVerbose
End Property

Private Property Get BttnSkipAll() As String
    BttnSkipAll = "Skip" & vbLf & "all"
End Property

Public Sub RawClones( _
               ByVal urc_wb As Workbook, _
               ByVal urc_comp_max_len As Long, _
               ByVal urc_service As String, _
               ByVal urc_clones As Dictionary, _
      Optional ByRef urc_log As clsLog = Nothing)
' --------------------------------------------------------
' Updates any raw clone in Workbook urc_wb
' --------------------------------------------------------
    Const PROC = "RawClones"
    
    On Error GoTo eh
    Dim sStatus     As String
    Dim vbc         As VBComponent
    Dim sServiced   As String
    Dim fso         As New FileSystemObject
    Dim lReplaced   As Long
    Dim sReplaced   As String
    Dim v           As Variant
'    Dim lKoCC       As enKindOfCodeChange
    Dim bVerbose    As Boolean
    Dim bSkip       As Boolean
    
    bVerbose = True
    sStatus = urc_service
    '~~ Prevent any action for a Workbook opened with any irregularity
    
'    Application.StatusBar = sStatus & "Resolve pending imports"
'    mPending.Resolve urc_wb
        
    For Each v In urc_clones
        Set vbc = v
        Set cComp = New clsComp
        cComp.Wrkbk = urc_wb
        cComp.CompName = vbc.name

        Application.StatusBar = sStatus & vbc.name & " "
        cComp.VBComp = vbc
        sServiced = cComp.Wrkbk.name & " Component """ & vbc.name & """"
        sServiced = sServiced & String(urc_comp_max_len - Len(vbc.name), ".")
        If Not urc_log Is Nothing Then urc_log.ServicedItem = sServiced

        With cComp
            Set cRaw = New clsRaw
            cRaw.CompName = .CompName
            cRaw.ExpFile = fso.GetFile(FilePath:=mHostedRaws.ExpFilePath(.CompName))
            cRaw.ExpFilePath = cRaw.ExpFile.Path
            cRaw.HostFullName = mHostedRaws.HostFullName(comp_name:=.CompName)
            cRaw.CloneExpFilePath = .ExpFilePath
            If cRaw.Changed Then
                Application.StatusBar = sStatus & vbc.name & " Renew of '" & .CompName & "' by import of '" & cRaw.ExpFilePath & "'"
                If bVerbose Then
                    UpdateCloneConfirmed ucc_comp_name:=vbc.name _
                                       , ucc_service:=urc_service _
                                       , ucc_stay_verbose:=bVerbose _
                                       , ucc_skip:=bSkip _
                                       , ucc_clone:=cComp.ExpFilePath _
                                       , ucc_raw:=cRaw.ExpFilePath
                End If
                If Not bSkip Then
                    mRenew.ByImport rn_wb:=.Wrkbk _
                         , rn_comp_name:=.CompName _
                         , rn_exp_file_full_name:=cRaw.ExpFilePath _
                         , rn_status:=sStatus & " " & .CompName & " "
                    If Not urc_log Is Nothing Then urc_log.Action = "Clone component renewed/updated by (re-)import of '" & cRaw.ExpFilePath & "'"
                    lReplaced = lReplaced + 1
                    sReplaced = .CompName & ", " & sReplaced
                    '~~ Register the update being used to identify a potentially relevant
                    '~~ change of the origin code
                End If
            End If
            sStatus = sStatus & " " & .CompName & ", "
        End With
        Set cComp = Nothing
        Set cRaw = Nothing
        
    Next v
    DsplyStatusUpdateClonesResult sp_cloned_raws:=urc_clones.Count _
                                , sp_no_replaced:=lReplaced _
                                , sp_replaced:=sReplaced _
                                , sp_service:=urc_service
        
    
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function UpdateCloneConfirmed( _
                               ByVal ucc_comp_name As String, _
                               ByVal ucc_service As String, _
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
    Dim bStayVerbose    As Boolean
    Dim sMsg            As tMsg
    Dim vReply          As Variant
    
    BttnDsplyChanges = ucc_comp_name
    BttnUpdateStayVerbose = ucc_comp_name
    Set cllButtons = _
    mMsg.Buttons(BttnDsplyChanges _
               , vbLf _
               , BttnUpdateStayVerbose _
               , BttnSkipStayVerbose _
               , vbLf _
               , BttnUpdateAll _
               , BttnSkipAll _
               )
    
    With sMsg
        .section(1).sLabel = "About"
        .section(1).sText = "When the cloned raw in this Workbook is not updated the message will show up again the next time this Workbook is opened in the configured development root:"
        .section(2).sText = mMe.RootServicedByCompMan
        .section(2).bMonspaced = True
    End With
    Do
        vReply = mMsg.Dsply(msg_title:=ucc_service & "Update " & Spaced(ucc_comp_name) & "with changed raw" _
                          , msg:=sMsg _
                          , msg_buttons:=cllButtons _
                           )
        Select Case vReply
            Case BttnDsplyChanges
                mFile.Compare file_left_full_name:=ucc_clone _
                            , file_left_title:="Cloned raw: '" & ucc_clone & "'" _
                            , file_right_full_name:=ucc_raw _
                            , file_right_title:="Current raw: '" & ucc_raw & "'"
            
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
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Sub DsplyStatusUpdateClonesResult( _
                           ByVal sp_cloned_raws As Long, _
                           ByVal sp_no_replaced As Long, _
                           ByRef sp_replaced As String, _
                           ByVal sp_service As String)
' --------------------------------------------------------
' Display the service progress
' --------------------------------------------------------
    Dim sMsg As String
    
    Select Case sp_cloned_raws
        Case 0: sMsg = sp_service & "No 'Cloned Raw Component' in Workbook"
        Case 1
            Select Case sp_no_replaced
                Case 0:     sMsg = sp_service & "1 has been identified as 'Cloned Raw Component' but has not been updated since the raw had not changed."
                Case 1:     sMsg = sp_service & "1 of 1 'Cloned Raw Component' has been updated because the raw had changed (" & left(sp_replaced, Len(sp_replaced) - 2) & ")."
            End Select
        Case Else
            Select Case sp_no_replaced
                Case 0:     sMsg = sp_service & "None of the " & sp_cloned_raws & " 'Cloned Raw Components' has been updated since none of the raws had changed."
                Case 1:     sMsg = sp_service & "1 of the " & sp_cloned_raws & "'Cloned Raw Components' has been updated since the raw's code had changed (" & left(sp_replaced, Len(sp_replaced) - 2) & ")."
                Case Else:  sMsg = sp_service & sp_no_replaced & " of the " & sp_cloned_raws & " 'Cloned Raw Components' have been updated because the raws had changed (" & left(sp_replaced, Len(sp_replaced) - 2) & ")."
            End Select
    End Select
    If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg

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

