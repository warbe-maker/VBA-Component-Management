Attribute VB_Name = "mInstall"
Option Explicit

Private Const BTT_INST_DONE = "Done"

Public Sub CommonComponents(Optional ByRef ic_wbk As Workbook)
' -----------------------------------------------------------
' Installs one or more 'Common-Componernts' by importing the
' selected 'Raw-Component's Export-File.
' -----------------------------------------------------------
    Const PROC = "CloneRaws"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim cll     As New Collection
    Dim i       As Long
    Dim vReply  As Variant
    Dim sMsg    As TypeMsg
    
    If ic_wbk Is Nothing Then Set ic_wbk = ActiveWorkbook
    If mMe.IsAddinInstnc Then GoTo xt
    
    For Each v In mComCompRawsGlobal.Components
        If i >= 7 Then
            cll.Add vbLf
            i = 0
        End If
        If Not mComp.Exists(ic_wbk, v) Then
            cll.Add v
            i = i + 1
        End If
    Next v
    cll.Add vbLf
    cll.Add BTT_INST_DONE
    sMsg.Section(1).Text.Text = ""
    sMsg.Section(2).Label.Text = "Please note!"
    sMsg.Section(2).Text.Text = "The selection contains all known 'Raw-Components/Common-Components' which are not already installed " & _
                            "(i.e. imported). Any components missed may either not be indicated 'hosted' in any Workbook or the Workbook " & _
                            "does not reside within the configured 'Serviced-Root-Folder'" & vbLf & _
                            "(currently  " & mBasic.Spaced(mConfig.ServicedDevAndTestFolder) & "  )."
    
    Do
        If Not mMsg.IsValidMsgButtonsArg(cll) Then Stop
        vReply = mMsg.Dsply(dsply_title:="Select one of the available 'Raw-Components/Common-Components') yet not installed in '" & ic_wbk.Name & "' or press '" & VBA.Replace(BTT_INST_DONE, vbLf, " ") & "'" _
                        , dsply_msg:=sMsg _
                        , dsply_buttons:=cll _
                         )
        Select Case vReply
            Case BTT_INST_DONE: Exit Do
            Case Else
                mRenew.ByImport bi_wbk_serviced:=ic_wbk _
                              , bi_vbc_name:=vReply _
                              , bi_exp_file:=mComCompRawsGlobal.SavedExpFileFullName(vReply)
        End Select
    Loop
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mInstall" & "." & sProc
End Function

