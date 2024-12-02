VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsConfig: Maintains and provides configuration items in a
' ======================= Private Profile File called CompMan.cfg.
'
' Public Properties:
' ------------------
'
' Uses the Common VBA Private Prifile services (clsPrivProf)
'
' W. Rauschenberger Berlin, Nov 2024
' ---------------------------------------------------------------------------
Private Const VALUE_NAME_        As String = ""
Private sSectionName             As String

Public PPFile                    As New clsPrivProf

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
        
    On Error GoTo eh
    sSectionName = "CompManConfig"
    With PPFile
        .FileName = mEnvironment.CompManConfigFileFullName
        .FileHeader = FileHeader
        .FileFooter = mEnvironment.PrivateProfileFileFooter
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Class_Terminate()
    Set PPFile = Nothing
End Sub

Private Function FileHeader() As String
    FileHeader = ""
End Function

Friend Property Let x(ByVal s As String):   Value(VALUE_NAME_) = s:                 End Property

Friend Property Get x() As String:          x = Value(VALUE_NAME_REG_STAT_COMP):    End Property


Public Property Get NameHskpngSection() As String:         NameHskpngSection = PPFile.NameHskpngSection:    End Property

Private Property Get Value(Optional ByVal v_value_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    Value = PPFile.Value(v_value_name, sSectionName)

End Property

Private Property Let Value(Optional ByVal v_value_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    PPFile.Value(v_value_name, sSectionName) = v_value
                                    
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommonServiced" & "." & sProc
End Function

