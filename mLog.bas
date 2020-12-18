Attribute VB_Name = "mLog"
Option Explicit
' ------------------------------------------------
' Standard Module mLog:
' Logging of any of the CompMan AddIn's activities
'
' Services: LogAction
'
' W. Raschenberger, Berlin Nov 2020
' ------------------------------------------------
Private Declare PtrSafe Function WritePrivateProfileString _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal lpw_ApplicationName As String, _
                ByVal lpw_KeyName As String, _
                ByVal lpw_String As String, _
                ByVal lpw_FileName As String) As Long
                
Private Declare PtrSafe Function GetPrivateProfileString _
                Lib "kernel32" Alias "GetPrivateProfileStringA" _
               (ByVal lpg_ApplicationName As String, _
                ByVal lpg_KeyName As String, _
                ByVal lpg_Default As String, _
                ByVal lpg_ReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpg_FileName As String) As Long

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompLog." & sProc
End Function

Public Sub LogAction(ByVal log_action As String, _
            Optional ByVal log_wb As String = vbNullString)
' ---------------------------------------------------------
' Log the action (sLog) started at (dtLog).
' ----------------------------------------------
    If wbAddIn.IsDevlpInstance Then Exit Sub
    
    If log_wb = vbNullString Then log_wb = ActiveWorkbook.name
    ValueLet vl_section:=log_wb _
           , vl_value_name:=Format(Now, "yyyy-mm-dd hh:mm:ss") _
           , vl_value:=log_action _
           , vl_file:=mCfg.CompManAddinPath & "\CompMan.log"
End Sub

Private Function ValueGet( _
                    ByVal vg_section As String, _
                    ByVal vg_value_name As String, _
                    ByVal vg_file As String) As String
' -----------------------------------------------------
'
' -----------------------------------------------------
    Const PROC  As String = "ValueGet"
    
    On Error GoTo eh
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant

    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString( _
                                      lpg_ApplicationName:=vg_section _
                                    , lpg_KeyName:=vg_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=vg_file _
                                     )
    vValue = left$(sRetVal, lResult)
    ValueGet = vValue
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Sub ValueLet( _
               ByVal vl_section As String, _
               ByVal vl_value_name As String, _
               ByVal vl_value As Variant, _
               ByVal vl_file As String)
' -------------------------------------------
'
' -------------------------------------------
    Const PROC = "ValueLet"
        
    On Error GoTo eh
    Dim lChars As Long
    
    Select Case VarType(vl_value)
        Case vbBoolean: lChars = WritePrivateProfileString(lpw_ApplicationName:=vl_section _
                                                         , lpw_KeyName:=vl_value_name _
                                                         , lpw_String:=VBA.CStr(VBA.CLng(vl_value)) _
                                                         , lpw_FileName:=vl_file _
                                                         )
        Case Else:      lChars = WritePrivateProfileString(vl_section, vl_value_name, CStr(vl_value), vl_file)
    End Select
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & vl_section & "'" & vbLf & _
               "Value name = '" & vl_value_name & "'" & vbLf & _
               "Value      = '" & CStr(vl_value) & "'" & vbLf & _
               "Value file = '" & vl_file & "'"
    End If

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

