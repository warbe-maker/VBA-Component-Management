Attribute VB_Name = "mRenew"
Option Explicit

Private Property Get NowMsec() As String
    NowMsec = Format(Now(), "hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Property

Public Sub ByImport( _
              ByRef rn_wb As Workbook, _
              ByVal rn_comp_name As String, _
              ByVal rn_raw_exp_file_full_name As String)
' -----------------------------------------------------
' Renews/replaces the component (rn_comp_name) in
' Workbook (rn_wb) by importing the Export-File
' (rn_raw_exp_file_full_name).
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
    
    mBasic.BoP ErrSrc(PROC)
    SaveWbk rn_wb
    
    With rn_wb.VBProject
        If mComp.Exists(wb:=rn_wb, comp_name:=rn_comp_name) Then
            '~~ Find a free/unused temporary name
            sTempName = mComp.TempName(tn_wb:=rn_wb, tn_comp_name:=rn_comp_name)
            '~~ Rename the component when it already exists
            .VBComponents(rn_comp_name).name = sTempName
            Log.Entry = NowMsec & " '" & rn_comp_name & "' renamed to '" & sTempName & "'"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Log.Entry = NowMsec & " '" & sTempName & "' removed (may be postponed by the system however)"
        End If
    
        '~~ (Re-)import the component
        .VBComponents.Import rn_raw_exp_file_full_name
        Log.Entry = NowMsec & " '" & rn_comp_name & "' (re-)imported from '" & rn_raw_exp_file_full_name & "'"
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = rn_wb
            .CompName = rn_comp_name
        End With
        .VBComponents(rn_comp_name).Export Comp.ExpFileFullName
        Log.Entry = NowMsec & " '" & rn_comp_name & "' exported to '" & Comp.ExpFileFullName & "'"
        SaveWbk rn_wb
    End With
          
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mRenew." & s
End Function

Private Sub SaveWbk(ByRef rs_wb As Workbook)
    Const PROC = "SaveWbk"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    Application.EnableEvents = False
    rs_wb.Save
    Log.Entry = "Workbook saved (DoEvents delayed continuation for " & TimedDoEvents & " msec)"
    Application.EnableEvents = True

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function TimedDoEvents() As String
    mBasic.TimerBegin
    DoEvents
    TimedDoEvents = Format(mBasic.TimerEnd, "00000")
End Function

