Attribute VB_Name = "mInstall"
Option Explicit

Private Const BTT_INST_DONE = "Done"

Public Sub CommonComponents(ByVal cc_wbk As Workbook, _
                   Optional ByVal cc_names As String = vbNullString)
' ------------------------------------------------------------------------------
' Installs one or more 'Common-Componernts' by importing the selected
' Raw-Component's Export-File.
' ------------------------------------------------------------------------------
    Const PROC = "CloneRaws"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim cll     As New Collection
    Dim i       As Long
    Dim vReply  As Variant
    Dim sMsg    As TypeMsg
    
    If mMe.IsAddinInstnc Then GoTo xt
    If cc_names <> vbNullString Then
        cc_wbk.VBProject.VBComponents.Import mCommComps.SavedExpFileFullName(cc_names)
    Else
        For Each v In mCommComps.Components
            If i >= 7 Then
                cll.Add vbLf
                i = 0
            End If
            If Not mComp.Exists(cc_wbk, v) Then
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
                                "(currently  " & mBasic.Spaced(wsConfig.FolderCompManRoot) & "  )."
        
        Do
            If Not mMsg.IsValidMsgButtonsArg(cll) Then Stop
            vReply = mMsg.Dsply(dsply_title:="Select one of the available 'Raw-Components/Common-Components') yet not installed in '" & cc_wbk.Name & "' or press '" & VBA.Replace(BTT_INST_DONE, vbLf, " ") & "'" _
                            , dsply_msg:=sMsg _
                            , dsply_buttons:=cll _
                             )
            Select Case vReply
                Case BTT_INST_DONE: Exit Do
                Case Else
                    mChanged.ReImport bi_wbk_serviced:=cc_wbk _
                                    , bi_vbc_name:=vReply _
                                    , bi_exp_file:=mCommComps.SavedExpFileFullName(vReply)
            End Select
        Loop
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mInstall" & "." & sProc
End Function

