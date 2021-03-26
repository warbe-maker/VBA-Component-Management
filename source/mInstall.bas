Attribute VB_Name = "mInstall"
Option Explicit

Private Const BTT_INST_DONE = "Done"

Public Sub CloneRaws(Optional ByRef ic_wb As Workbook)
' ----------------------------------------------------
' Installs one or more raw components by importing
' their Export-File.
' ----------------------------------------------------
    Const PROC = "CloneRaws"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim cll     As New Collection
    Dim i       As Long
    Dim vReply  As Variant
    Dim sMsg    As tMsg
    
    If ic_wb Is Nothing Then Set ic_wb = ActiveWorkbook
    If mMe.IsAddinInstnc Then GoTo xt
    
    For Each v In mRawsHosted.Components
        If i >= 7 Then
            cll.Add vbLf
            i = 0
        End If
        If Not mCompMan.CompExists(ce_wb:=ic_wb, ce_comp_name:=v) Then
            cll.Add v
            i = i + 1
        End If
    Next v
    cll.Add vbLf
    cll.Add BTT_INST_DONE
    sMsg.Section(1).sText = ""
    sMsg.Section(2).sLabel = "Please note!"
    sMsg.Section(2).sText = "The selection contains all known 'raw' components hosted in another Workbook " & _
                            "which are not already installed (i.e. imported). Any components missed may not be " & _
                            "indicated 'hosted' in any Workbook maintained within the 'manged root'"
    sMsg.Section(3).sText = mMe.ServicedRoot
    sMsg.Section(3).bMonspaced = True
    
    Do
        vReply = mMsg.Dsply(msg_title:="Please select the 'raw' component to be imported into '" & ic_wb.Name & "' or press '" & VBA.Replace(BTT_INST_DONE, vbLf, " ") & "'" _
                        , msg:=sMsg _
                        , msg_buttons:=cll _
                         )
        Select Case vReply
            Case BTT_INST_DONE: Exit Do
            Case Else
                mRenew.ByImport rn_wb:=ic_wb _
                              , rn_comp_name:=vReply _
                              , rn_exp_file_full_name:=mRawsHosted.ExpFileFullName(vReply)
        End Select
    Loop
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mInstall" & "." & sProc
End Function

