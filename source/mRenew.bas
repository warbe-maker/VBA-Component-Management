Attribute VB_Name = "mRenew"
Option Explicit

Private Property Get NowMsec() As String
    NowMsec = Format(Now(), "hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Property

Public Sub ByImport( _
              ByRef rn_wb As Workbook, _
              ByVal rn_comp_name As String, _
              ByVal rn_exp_file_full_name As String)
' -----------------------------------------------------
' Renews/replaces the component (rn_comp_name) in
' Workbook (rn_wb) by importing the Export-File
' (rn_exp_file_full_name).
' Note: Because a module cannot be deleted it is
'       renamed and deleted. The rename puts it out of
'       the way, deletion is done by the system when
'       the process has ended.
' -----------------------------------------------------
    Const PROC = "ByImport"
    
    On Error GoTo eh
    Dim sTempName       As String
    Dim fso             As New FileSystemObject
    Dim Comp           As clsComp
    
    Debug.Print NowMsec & " =========================="
    SaveWbk rn_wb
    DoEvents:  Application.Wait Now() + 0.0000001 ' wait for 10 milliseconds
    With rn_wb.VBProject
        If mComp.Exists(wb:=rn_wb, comp_name:=rn_comp_name) Then
            '~~ Find a free/unused temporary name
            sTempName = mComp.TempName(wb:=rn_wb, comp_name:=rn_comp_name)
            '~~ Rename the component when it already exists
            .VBComponents(rn_comp_name).Name = sTempName
            Debug.Print NowMsec & " '" & rn_comp_name & "' renamed to '" & sTempName & "'"
'           DoEvents:  Application.Wait Now() + 0.0000001 ' wait for 10 milliseconds
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Debug.Print NowMsec & " '" & sTempName & "' removed (may be postponed by the system however)"
        End If
    
        '~~ (Re-)import the component
        .VBComponents.Import rn_exp_file_full_name
        Debug.Print NowMsec & " '" & rn_comp_name & "' (re-)imported from '" & rn_exp_file_full_name & "'"
        Set Comp = New clsComp
        Set Comp.Wrkbk = rn_wb
        Comp.CompName = rn_comp_name
        .VBComponents(rn_comp_name).Export Comp.ExpFileFullName
        Debug.Print NowMsec & " '" & rn_comp_name & "' exported to '" & Comp.ExpFileFullName & "'"
    End With
          
xt: Set fso = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mRenew." & s
End Function

Private Sub SaveWbk(ByRef rs_wb As Workbook)
    Application.EnableEvents = False
    rs_wb.Save
    Application.EnableEvents = True
End Sub

