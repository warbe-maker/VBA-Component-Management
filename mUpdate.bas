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


Public Sub ClonesTheRawHasChanged( _
                            ByVal uc_wb As Workbook, _
                            ByVal uc_comp_max_len As Long, _
                            ByVal uc_service As String, _
                   Optional ByRef uc_log As clsLog = Nothing)
' ---------------------------------------------------------
'
' ---------------------------------------------------------
    Const PROC = "ClonesTheRawHasChanged"
    
    On Error GoTo eh
    Dim sStatus     As String
    Dim vbc         As VBComponent
    Dim lComponents As Long
    Dim sServiced   As String
    Dim fso         As New FileSystemObject
    Dim lClonedRaws As Long
    Dim lReplaced   As Long
    Dim sReplaced   As String
    
    sStatus = uc_service
    '~~ Prevent any action for a Workbook opened with any irregularity
    
    Application.StatusBar = sStatus & "Resolve pending imports"
    mPending.Resolve uc_wb
        
    For Each vbc In uc_wb.VBProject.VBComponents
        Set cComp = New clsComp
        cComp.Wrkbk = uc_wb
        cComp.CompName = vbc.name

        Application.StatusBar = sStatus & vbc.name & " "
        cComp.VBComp = vbc
        lComponents = lComponents + 1
        sServiced = cComp.Wrkbk.name & " Component """ & vbc.name & """"
        sServiced = sServiced & String(uc_comp_max_len - Len(vbc.name), ".")
        If Not uc_log Is Nothing Then uc_log.ServicedItem = sServiced
            
        If cComp.KindOfComp = enRawClone Then
            '~~ Establish a component class object which represents the cloned raw's remote instance
            '~~ which is hosted in another Workbook
            Set cRaw = New clsRaw
            With cRaw ' Provide all available information rearding the remote raw component
                .CompName = cComp.CompName
                .ExpFile = fso.GetFile(FilePath:=mRaw.ExpFileFullName(.CompName))
                .ExpFileFullName = .ExpFile.PATH
                .HostFullName = mRaw.HostFullName(comp_name:=.CompName)
            End With

            With cComp
                If .KindOfComp = enRawClone Then lClonedRaws = lClonedRaws + 1
                If .KindOfCodeChange = enRawOnly _
                Or .KindOfCodeChange = enRawAndClone Then
                    '~~ Attention!! The cloned raw's code is updated disregarding any code changed in it.
                    '~~ A code change in the cloned raw is only considered when the Workbook is about to
                    '~~ be closed - where it may be ignored to make exactly this happens.
                    Application.StatusBar = sStatus & vbc.name & " Renew of '" & .CompName & "' by import of '" & cRaw.ExpFileFullName & "'"
                    mRenew.ByImport rn_wb:=.Wrkbk _
                         , rn_comp_name:=.CompName _
                         , rn_exp_file_full_name:=cRaw.ExpFileFullName _
                         , rn_status:=sStatus & " " & .CompName & " "
                    If Not uc_log Is Nothing Then uc_log.Action = "Clone component renewed/updated by (re-)import of '" & cRaw.ExpFileFullName & "'"
                    lReplaced = lReplaced + 1
                    sReplaced = .CompName & ", " & sReplaced
                    '~~ Register the update being used to identify a potentially relevant
                    '~~ change of the origin code
                End If
                sStatus = sStatus & " " & .CompName & ", "
            End With
        End If
        Set cComp = Nothing
        Set cRaw = Nothing
        
    Next vbc
    DsplyStatusUpdateClonesResult sp_total_comps:=lComponents _
                                , sp_cloned_raws:=lClonedRaws _
                                , sp_no_replaced:=lReplaced _
                                , sp_replaced:=sReplaced _
                                , sp_service:=uc_service
        
    
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function UpdateCloneConfirmed(ByVal ucc_comp_name As String)
    
    Dim cllButtons As Collection
    Dim bStayVerbose As Boolean
    
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
                           ByVal sp_total_comps As Long, _
                           ByVal sp_cloned_raws As Long, _
                           ByVal sp_no_replaced As Long, _
                           ByRef sp_replaced As String, _
                           ByVal sp_service As String)
' --------------------------------------------------------
' Display the service progress
' --------------------------------------------------------
    Dim sMsg As String
    
    Select Case sp_cloned_raws
        Case 0: sMsg = sp_service & "None of " & sp_total_comps & " components had been identified as 'Cloned Raw Component'."
        Case 1
            Select Case sp_no_replaced
                Case 0:     sMsg = sp_service & "1 has been identified as 'Cloned Raw Component' but has not been updated since the raw had not changed."
                Case 1:     sMsg = sp_service & "1 of 1 'Cloned Raw Component' has been updated because the raw had changed (" & left(sp_replaced, Len(sp_replaced) - 2) & ")."
            End Select
        Case Else
            Select Case sp_no_replaced
                Case 0:     sMsg = sp_service & "None of the " & sp_cloned_raws & " 'Cloned Raw Components' has been updated since none of the raws had changed."
                Case 1:     sMsg = sp_service & "1 of the " & sp_cloned_raws & "'Cloned Raw Components' has been updated since the raw's code had changed (" & left(sp_replaced, Len(sp_replaced) - 2) & ")."
                Case Else:  sMsg = sp_service & sp_cloned_raws & " of the " & sp_total_comps & " 'Cloned Raw Components' have been updated because the raws had changed (" & left(sp_replaced, Len(sp_replaced) - 2) & ")."
            End Select
    End Select
    If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg

End Sub

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mUpdate" & "." & es_proc
End Function

