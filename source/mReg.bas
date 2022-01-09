Attribute VB_Name = "mReg"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mReg
' Simplified services for working with the Registry for configuration,
' initiation or other application values. All services only use the
' named arguments: reg_key, reg_value_name, reg_value whereby re_key is
' allways a string which includes the HKEY. The services provided are
' pretty much the same as for the same purpose when written to a file as
' section, name=value entries with the great difference that the registry
' allows any number of "sub-sections".
' Note: Not as universal as http://www.cpearson.com/excel/registry.htm but
' by far more simple to be used.
'
' Public services:
' Delete        Deletes a key or only a name/value
' Value-Get     Returns the value of a named entry
' Value-Let     Writes a provided value under the provided key/name
' Exists        Returns TRUE when the provided key or name exists
' Keys
' Names
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin Jan 2022
' See: https://github.com/warbe-maker/Common-VBA-Registry-Services
' ----------------------------------------------------------------------------
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As _
    Long
Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal numBytes As Long)

Private Const HKEY_CURRENT_USER      As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE     As Long = &H80000002
Private Const HKEY_CLASSES_ROOT      As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG    As Long = &H80000005
Private Const HKEY_DYN_DATA          As Long = &H80000006
Private Const HKEY_PERFORMANCE_DATA  As Long = &H80000004
Private Const HKEY_USERS             As Long = &H80000003

Private Const HKCU                   As Long = HKEY_CURRENT_USER
Private Const HKLM                   As Long = HKEY_LOCAL_MACHINE

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7
Private Const ERROR_MORE_DATA = 234
                                  
Private Const REG_VALUE_MIN_VALUE       As Long = -2147483648#
Private Const REG_VALUE_MAX_VALUE       As Long = 2147483647
Private Const REG_VALUE_MAX_LENGTH      As Long = &H2048
Private Const REG_KEY_MAX_LENGTH        As Long = &H255

Private Const KEY_WRITE                 As Long = 131078
Private Const KEY_READ                  As Long = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                                                           ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                                                           ' SYNCHRONIZE))
Dim oReg                                As Object
                        
Public Property Get Value( _
                 Optional ByVal reg_key As String, _
                 Optional ByVal reg_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value of the registry key's (reg_key) value identified by its
' name (reg_value_name).
' ----------------------------------------------------------------------------
    Const PROC = "Value-Get"
    
    On Error GoTo eh
    Dim rString As String
    
    rString = RegString(reg_key, reg_value_name)
    If Not IsValidKey(rString) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided reg_key | reg_value_name ('" & rString & "') lenght exceeds the maximum length of '" & REG_KEY_MAX_LENGTH & "'!"
    If Not HasAccess(reg_key, KEY_READ) Then
        ' Err.Raise AppErr(2), ErrSrc(PROC), "No access right to read '" & rString & "'!"
        ' Debug.Print "KEY_READ denied for '" & rString & "'"
    End If
    Value = CreateObject("WScript.Shell").RegRead(rString)
'    Debug.Print "Read value from '" & rString & "' successful"

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Value( _
                 Optional ByVal reg_key As String, _
                 Optional ByVal reg_value_name As String, _
                          ByVal reg_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (re_value) to the key (reg_key) value name (reg_value_name)
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
    
    On Error GoTo eh
    Dim rString As String
    
    rString = RegString(reg_key, reg_value_name)
'    Debug.Print rString
    
    If Not IsValidKey(rString) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided reg_key | reg_value_name ('" & rString & "') lenght exceeds the maximum length of '" & REG_KEY_MAX_LENGTH & "'!"
'    If Not HasAccess(reg_key, KEY_WRITE) _
'    Then Err.Raise AppErr(2), ErrSrc(PROC), "No access right to write a value for '" & rString & "'!"
    
    CreateObject("WScript.Shell").RegWrite rString, reg_value, RegType(reg_value)

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ArrayIsAllocated(arr As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = IsArray(arr) _
                       And Not IsError(LBound(arr, 1)) _
                       And LBound(arr, 1) <= UBound(arr, 1)
    
End Function

Public Function Delete(ByVal reg_key As String, _
              Optional ByVal reg_value_name As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' When no value name (reg_value_name) is provided the provided key is deleted
' else only the name. A possibly missing \ ate the end of the key (reg_key)
' ----------------------------------------------------------------------------
    Const PROC = "Delete"
    
    On Error GoTo eh
    Dim s       As String
    Dim rString As String
    Dim wsh     As New wshShell
    Dim sCmd    As String
    
    If reg_value_name <> vbNullString Then
        rString = RegString(reg_key, reg_value_name)
        If Right(rString, 1) = "\" Then
            rString = Left(rString, Len(rString) - 1)
        End If
        On Error GoTo eh
        CreateObject("WScript.Shell").RegDelete rString ' reg_key & reg_value_name
        Delete = True
    Else
        mReg.DeleteSubKeys reg_key:=reg_key
    End If
    
xt: Set wsh = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DeleteSubKeys(ByVal reg_key As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "DeleteSubKeys"
    
    On Error GoTo eh
    Dim SubKeys()   As Variant
    Dim SubKey      As Variant
    Dim oReg        As Object
    Static HKey     As Long
    
    RegHKey reg_key, HKey   ' Extract and transform hKey and unstrip it frrom reg_key
    If Right(reg_key, 1) = "\" Then reg_key = Left(reg_key, Len(reg_key) - 1)
    
    Set oReg = GetObject("winmgmts:\\" & "." & "\root\default:StdRegProv")
    oReg.EnumKey HKey, reg_key, SubKeys
    If ArrayIsAllocated(SubKeys) Then
        For Each SubKey In SubKeys
            DeleteSubKeys reg_key & "\" & SubKey
        Next
    End If
    oReg.DeleteKey HKey, reg_key

xt: Set oReg = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.name & " mReg." & sProc
End Function

Public Function Exists(ByVal reg_key As String, _
         Optional ByVal reg_value_name As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the key (reg_key) exists and when a name (reg_value_name)
' is provided when the name exists. No worry about \! Missing ones are added.
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim rString As String
    
    rString = RegString(reg_key, reg_value_name)
    
    If reg_value_name <> vbNullString Then
        If Right(reg_value_name, 1) = "\" _
        Then reg_value_name = Left(reg_value_name, Len(reg_value_name) - 1)
    End If
    Exists = False
    On Error GoTo xt
    '~~ To have reg_key interpreted as key it has to end with a \.
    CreateObject("WScript.Shell").RegRead rString ' reg_key & reg_value_name
    Exists = True
xt:
End Function

Public Function Export(ByVal reg_key As String) As String
' ----------------------------------------------------------------------------
' Returns a string in the form [section]
'                              name=value
' When the key (reg_key) contains sub-keys the returned string will have as
' many [section] lines. The result of the function written to a file is a
' typical .ini, .dat, .cfg file.
' ----------------------------------------------------------------------------
    Dim Keys        As Collection
    Dim Key         As Variant
    Dim ValueName   As Variant
    Dim Values      As Dictionary
    
    Set Keys = mReg.Keys(reg_key)
    If Keys.Count = 0 Then
        Set Values = mReg.Values(reg_key)
        If Values.Count > 0 Then
            If Right(reg_key, 1) = "\" Then
                Export = "[" & Split(reg_key, "\")(UBound(Split(reg_key, "\")) - 1) & "]"
            Else
                Export = "[" & Split(reg_key, "\")(UBound(Split(reg_key, "\"))) & "]"
            End If
            For Each ValueName In Values
                Export = Export & vbLf & ValueName & "=" & Values(ValueName)
            Next ValueName
        End If
    Else
        For Each Key In Keys
            Set Values = mReg.Values(reg_key & "\" & Key)
            If Values.Count > 0 Then
                If Len(Export) = 0 _
                Then Export = "[" & Key & "]" _
                Else Export = Export & vbLf & "[" & Key & "]"
                For Each ValueName In Values
                    Export = Export & vbLf & ValueName & "=" & Values(ValueName)
                Next ValueName
            End If
        Next Key
    End If
    
End Function

Private Function HasAccess(ByVal reg_key As String, ByVal reg_access As Long)
' ------------------------------------------------------------------------------
' Returns TRUE when access right (reg_access) is granted for the key (reg_key).
' ------------------------------------------------------------------------------
    Dim sComputer   As String
    Dim sMethod     As String
    Dim HKey        As Long
    Dim oRegistry   As Object
    Dim oInParam    As Object
    Dim hAccess     As Long
    Dim oOutParam   As Object
    Dim oMethod     As Object
    
    sComputer = "."
    sMethod = "CheckAccess"
    RegHKey reg_key, HKey ' extract and unstrip the hKey from the reg_key
    
    Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}//" & _
            sComputer & "/root/default:StdRegProv")
    
    Set oMethod = oRegistry.Methods_(sMethod)
    Set oInParam = oMethod.inParameters.SpawnInstance_()
    
    oInParam.hDefKey = HKey
    oInParam.sSubKeyName = reg_key
    oInParam.uRequired = reg_access
    Set oOutParam = oRegistry.ExecMethod_(sMethod, oInParam)
    
    HasAccess = oOutParam.Properties_("bGranted") = 0
End Function

Private Function IsValidDataType(ByVal vdt_var As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the data type of the variable (vdt_var) is valid for
' a registry entry.
' ----------------------------------------------------------------------------
    IsValidDataType = False
    If VarType(vdt_var) >= vbArray Then GoTo xt
    If IsArray(vdt_var) Then GoTo xt
    If IsObject(vdt_var) Then GoTo xt

    Select Case VarType(vdt_var)
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbDouble, vbInteger, vbLong, vbSingle, vbString
            IsValidDataType = True
        Case Else
            GoTo xt
    End Select
xt:
End Function

Private Function IsValidKey(ByVal reg_key As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the length of the key (reg_key) is LE REG_MAX_KEYLENGTH and
' all spaces unstripped longer 0.
' ------------------------------------------------------------------------------
    IsValidKey = (Len(reg_key) <= REG_KEY_MAX_LENGTH) And (Len(Trim(reg_key)) > 0)
End Function

Private Function IsValidValue(Optional ByVal reg_value As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when a numeric value (reg_value) ranges
' from -2,147,483,648 to 2,147,483,647 or a string - which is the maximum
' because a Registry-Type REG_DWORD stores a Long type value.
' ------------------------------------------------------------------------------
    If IsNumeric(reg_value) Then
        IsValidValue = reg_value >= REG_VALUE_MIN_VALUE And reg_value <= REG_VALUE_MAX_VALUE
    ElseIf VarType(reg_value) = vbString Then
        IsValidValue = Len(reg_value) <= REG_VALUE_MAX_LENGTH And Len(Trim(reg_value)) > 0
    End If
End Function

Public Function Keys(ByVal reg_key As String) As Collection
' ----------------------------------------------------------------------------
' Returns a collection of all sub-keys under the key (reg_key). When the key
' does not have sub-keys the returned Collection is empty
' ----------------------------------------------------------------------------
    Dim Computer    As String
    Dim SubKeys()   As Variant
    Dim SubKey      As Variant
    Dim HKey        As Long
    Dim oReg        As Object
    
    Set Keys = New Collection
    If Not mReg.Exists(reg_key) Then GoTo xt
    RegHKey reg_key, HKey ' extract/transform HKey and unstrip it from re_key
    If Right(reg_key, 1) = "\" Then reg_key = Left(reg_key, Len(reg_key) - 1)
    Computer = "."

    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Computer & "\root\default:StdRegProv")
    oReg.EnumKey HKey, reg_key, SubKeys
    If ArrayIsAllocated(SubKeys) Then
        For Each SubKey In SubKeys
            Keys.Add SubKey
        Next
    End If
xt:
End Function

Private Sub RegHKey(ByRef reg_key As String, _
                    ByRef reg_hkey As Long)
' ----------------------------------------------------------------------------
' When the key (reg_key) begins with a HKEY it is extracted/unstripped and
' transformed to Long.
' The procedure is used wherever the HKey value is required rather than the
' full key as a string. However, any other HKey but HKEY_CURRENT_USER (HKCU)
' or HKEY_LOCAL_MACHINE (or HKML) is invalid and raises an application error.
' ----------------------------------------------------------------------------
    Const PROC = "RegHKey"
    
    On Error GoTo eh
    
    Select Case Split(reg_key, "\")(0)
        Case "HKEY_CURRENT_USER":       reg_hkey = HKEY_CURRENT_USER:       reg_key = Replace(reg_key, "HKEY_CURRENT_USER\", vbNullString)
        Case "HKCU":                    reg_hkey = HKEY_CURRENT_USER:       reg_key = Replace(reg_key, "HKCU\", vbNullString)
        Case "HKEY_LOCAL_MACHINE":      reg_hkey = HKEY_LOCAL_MACHINE:      reg_key = Replace(reg_key, "HKEY_LOCAL_MACHINE\", vbNullString)
        Case "HKLM":                    reg_hkey = HKEY_LOCAL_MACHINE:      reg_key = Replace(reg_key, "HKLM\", vbNullString)
        Case Else
'            Err.Raise AppErr(1), ErrSrc(PROC), "The provided key '" & reg_key & "' does not begin with " & _
'                                               "HKEY_CURRENT_USER (or HKCU) or HKEY_LOCAL_MACHINE (or HKLM)!"
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function RegString(ByRef reg_key As String, ByVal reg_value_name As String) As String
    RegString = VBA.Replace(reg_key & "\" & reg_value_name, "\\", "\")
    RegString = VBA.Replace(RegString, "HKCU", "HKEY_CURRENT_USER")
    RegString = VBA.Replace(RegString, "HKLM", "HKEY_LOCAL_MACHINE")
End Function

Private Function RegType(ByVal reg_value As Variant) As String
' ----------------------------------------------------------------------------
' Returns the Registration Type for the value (reg_value).
' ----------------------------------------------------------------------------
    Const PROC = "RegType"
    
    On Error GoTo eh
    Select Case VarType(reg_value)
        ' Type      To store
        ' --------- --------------------------------
        ' REG_SZ    A string
        ' REG_DWORD A 32-bit number (4 Bytes = Long)
        ' ------------------------------------------
        Case vbBoolean:     RegType = "REG_SZ":     reg_value = CStr(Abs(CInt(reg_value))) ' reads as 0 or 1
        Case vbByte:        RegType = "REG_DWORD"
        Case vbCurrency:    RegType = "REG_DWORD"
        Case vbDate:        RegType = "REG_SZ"
        Case vbDecimal:     RegType = "REG_DWORD"
        Case vbDouble:      RegType = "REG_DWORD"
        Case vbInteger:     RegType = "REG_DWORD"
        Case vbLong:        RegType = "REG_DWORD"
        Case vbSingle:      RegType = "REG_DWORD"
        Case vbString:      RegType = "REG_SZ"
        Case Else:          Err.Raise AppErr(1), ErrSrc(PROC), "The VarType of the value cannot be written to the registry!"
    End Select


xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Values(ByVal reg_key As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary of all values under a given registry key (reg_key)
' where each entry's key is the value name and the item is the value.
' ----------------------------------------------------------------------------
    Dim dataLen         As Long
    Dim handle          As Long
    Dim Index           As Long
    Dim name            As String
    Dim ValueNameLength         As Long
    Dim resLong         As Long
    Dim resString       As String
    Dim RetVal          As Long
    Dim valueType       As Long
    Dim HKey            As Long
    Dim ValueName       As String
    Dim Value           As Variant
    
    Set Values = New Dictionary
       
    ' Open the key, exit if not found.
    RegHKey reg_key, HKey
    If Len(reg_key) Then
        If RegOpenKeyEx(HKey, reg_key, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use reg_hkey
        HKey = handle
    End If
    
    Do
        ValueNameLength = 260           ' this is the max length for a key name
        name = Space$(ValueNameLength)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        RetVal = RegEnumValue(HKey, Index, name, ValueNameLength, ByVal 0&, valueType, resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            RetVal = RegEnumValue(HKey, Index, name, ValueNameLength, ByVal 0&, valueType, resBinary(0), dataLen)
        End If
        
        If RetVal Then Exit Do              ' exit the loop if any other error (typically, no more values)
        ValueName = Left$(name, ValueNameLength)    ' retrieve the value's name
        
        '~~ Return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                Value = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                Value = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                Value = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                Value = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        Values.Add ValueName, Value
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

