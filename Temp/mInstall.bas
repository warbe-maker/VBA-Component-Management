Attribute VB_Name = "mInstall"
Option Explicit

Private Property Get INST_FINISHED() As String: INST_FINISHED = "Installation" & vbLf & "Finished": End Property

Public Sub Component(Optional ByVal ic_wb As Workbook)
' ----------------------------------------------------
' Installs one or more raw components by importing
' their Export File.
' ----------------------------------------------------
    Const PROC = "Install"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim cll     As New Collection
    Dim i       As Long
    Dim vReply  As Variant
    
    If ic_wb Is Nothing Then Set ic_wb = ActiveWorkbook
    If mMe.IsAddinInstnc Then GoTo xt
    
    For Each v In mHostedRaws.Components
        If i >= 7 Then
            cll.Add vbLf
            i = 0
        End If
        cll.Add v
        i = i + 1
    Next v
    cll.Add vbLf
    cll.Add INST_FINISHED
    
    Do
        vReply = mMsg.Box(msg_title:="Please select the component to be installed in '" & ic_wb.name & "' or press '" & VBA.Replace(INST_FINISHED, vbLf, " ") & "'" _
                        , msg_buttons:=cll _
                         )
        Select Case vReply
            Case INST_FINISHED: Exit Do
            Case Else
                mRenew.ByImport rn_wb:=ic_wb _
                              , rn_comp_name:=vReply _
                              , rn_exp_file_full_name:=mHostedRaws.ExpFilePath(vReply)
        End Select
    Loop
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mInstall" & "." & sProc
End Function

