Attribute VB_Name = "mReg"
Option Explicit

Public Const DWordRegKeyEnabled As Integer = 1
Public Const DWordRegKeyDisabled As Integer = 0

'Private RegKeyStr                As String
'Private RegKeyLocStr             As String
'Private RegKeyNameStr            As String
'Private iRegKeyDesiredStateInt   As Integer
'Private RegKeyCurrentStateInt    As Integer
'Private RegKeyFoundBool          As Boolean

Public Property Let Value( _
                 Optional ByVal v_name As String, _
                          ByVal v As String)
' ----------------------------------------------------------------------------
' object.RegWrite(strName, anyValue [,strType])
' Note: To have reg_name interpreted as value name it must not end with a \.
' strType: Type          Description                 in the form of
'          REG_SZ        A string                    A string
'          REG_DWORD     a Number                    An integer
'          REG_BINARY    A binary value              An integer
'          REG_EXPAND_SZ An expandable string        A string
'                       (e.g., "%windir%\\calc.exe")
' ----------------------------------------------------------------------------
    CreateObject("WScript.Shell").RegWrite v_name, v, "REG_SZ"
End Property

Public Property Get Value( _
                 Optional ByVal v_name As String) As String
' ----------------------------------------------------------------------------
' Returns the value of the registry kex (reg_key).
' Note: To have reg_name interpreted as value name it must not end with a \.
' ----------------------------------------------------------------------------
    Value = CreateObject("WScript.Shell").RegRead(v_name)
End Property

Public Function NameExists(ByVal REG_NAME As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the key (reg_key_string) exists in the Registry.
' Note: To have reg_name interpreted as value name it must not end with a \.
' ----------------------------------------------------------------------------
    NameExists = False
    On Error GoTo xt
    CreateObject("WScript.Shell").RegRead (REG_NAME)
    NameExists = True
xt: Exit Function
End Function

Public Function Exists(ByVal REG_KEY As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the key (reg_key_string) exists in the Registry.
' Note: To have reg_key interpreted as key it has to end with a \.
' ----------------------------------------------------------------------------
    Exists = False
    On Error GoTo xt
    CreateObject("WScript.Shell").RegRead (REG_KEY)
    Exists = True
xt: Exit Function
End Function

Public Sub Delete(ByVal reg_item As String)
' ----------------------------------------------------------------------------
' When reg_item ends with a \ the respective key is deleted when it does not
' end with a \ the respective name is deleted.
' ----------------------------------------------------------------------------
    CreateObject("WScript.Shell").RegDelete reg_item
End Sub
