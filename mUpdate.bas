Attribute VB_Name = "mUpdate"
Option Explicit
' -----------------
'
' -----------------
Private Property Get BTTN_UPDATE_STAY_VERBOSE(Optional ByVal comp_name As String) As String
    BTTN_UPDATE_STAY_VERBOSE = "Update '" & comp_name & "'" & vbLf & "stay verbose"
End Property

Private Property Get BTTN_UPDATE_ALL() As String
    BTTN_UPDATE_ALL = "Update all"
End Property

Private Property Get BTTN_DSPLY_CHANGES(Optional ByVal comp_name As String) As String
    BTTN_DSPLY_CHANGES = "Display changes for" & vbLf & "'" & comp_name & "'"
End Property

Private Property Get BTTN_SKIP_UPDATE_STAY_VERBOSE(Optional ByVal comp_name As String) As String
    BTTN_SKIP_UPDATE_STAY_VERBOSE = "Skip update of " & vbLf & "'" & comp_name & "'" & vbLf & "stay verbose"
End Property

Private Property Get BTTN_SKIP_UPDATE_ALL() As String
    BTTN_SKIP_UPDATE_ALL = "Skip update for all" & vbLf & "clone components"
End Property

Public Sub RawClones( _
               ByVal urc_wb As Workbook, _
               ByVal urc_comp_max_len As Long, _
               ByVal urc_service As String, _
      Optional ByRef urc_log As clsLog = Nothing)
' ----------------------------------------------
' Updates any raw clone in Workbook urc_wb
' ----------------------------------------------
    Const PROC = "RawClones"
    
    On Error GoTo eh
    Dim sStatus     As String
    Dim vbc         As VBComponent
    Dim sServiced   As String
    Dim fso         As New FileSystemObject
    Dim lReplaced   As Long
    Dim sReplaced   As String
    Dim v           As Variant
    Dim lKoCC       As enKindOfCodeChange
    Dim dct         As Dictionary
    
    sStatus = urc_service
    '~~ Prevent any action for a Workbook opened with any irregularity
    
    Application.StatusBar = sStatus & "Resolve pending imports"
    mPending.Resolve urc_wb
        
    Set dct = mCompMan.Clones(urc_wb)
    For Each v In dct
        Set vbc = v
        lKoCC = dct(v)
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
            cRaw.ExpFile = fso.GetFile(FilePath:=mRaw.ExpFileFullName(.CompName))
            cRaw.ExpFileFullName = .ExpFile.PATH
            cRaw.HostFullName = mRaw.HostFullName(comp_name:=.CompName)
            
            If lKoCC = enRawOnly _
            Or lKoCC = enRawAndClone Then
                '~~ Attention!! The cloned raw's code is updated disregarding any code changed in it.
                '~~ A code change in the cloned raw is only considered when the Workbook is about to
                '~~ be closed - where it may be ignored to make exactly this happens.
                Application.StatusBar = sStatus & vbc.name & " Renew of '" & .CompName & "' by import of '" & cRaw.ExpFileFullName & "'"
                mRenew.ByImport rn_wb:=.Wrkbk _
                     , rn_comp_name:=.CompName _
                     , rn_exp_file_full_name:=cRaw.ExpFileFullName _
                     , rn_status:=sStatus & " " & .CompName & " "
                If Not urc_log Is Nothing Then urc_log.Action = "Clone component renewed/updated by (re-)import of '" & cRaw.ExpFileFullName & "'"
                lReplaced = lReplaced + 1
                sReplaced = .CompName & ", " & sReplaced
                '~~ Register the update being used to identify a potentially relevant
                '~~ change of the origin code
            End If
            sStatus = sStatus & " " & .CompName & ", "
        End With
        Set cComp = Nothing
        Set cRaw = Nothing
        
    Next v
    DsplyStatusUpdateClonesResult sp_cloned_raws:=dct.Count _
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

Private Function UpdateCloneConfirmed(ByVal ucc_comp_name As String)
    
    Dim cllButtons      As Collection
    Dim bStayVerbose    As Boolean
    Dim sMsg            As tMsg
    
    Set cllButtons = _
    mMsg.Buttons(BTTN_DSPLY_CHANGES(ucc_comp_name) _
               , vbLf _
               , BTTN_UPDATE_STAY_VERBOSE(ucc_comp_name) _
               , BTTN_SKIP_UPDATE_STAY_VERBOSE _
               , vbLf _
               , BTTN_UPDATE_ALL _
               , BTTN_SKIP_UPDATE_ALL _
               )
    
    Select Case mMsg.Dsply(msg_title:="" _
                         , msg:=sMsg _
                         , msg_buttons:=cllButtons _
                          )
        Case ""
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

