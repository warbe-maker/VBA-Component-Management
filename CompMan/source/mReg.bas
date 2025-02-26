Attribute VB_Name = "mReg"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mReg: Services for working with the Registry, i.e. write
' ===================== values to and read values from the Registry within a
' VB-Project. For a maximum simplicity all services use the following named
' arguments:
' ._reg_key           String expression, obligatory, starts with
'                   "HKEY_CURRENT_USER" or "HKEY_LOCAL_MACHINE" followed by
'                   sub-keys. The reg_key may well function as the base-key
'                   for a VB-Project. It sjhould be noticed that writing to
'                   "HKEY_LOCAL_MACHINE" requires elevated permissions which
'                   are not recommendable at all.
' ._reg_value_name  may be prefixed by sub-keys
' ._reg_value       any format and any length. Values unable to be stored in
'                   the registry raise an error
'
' Note: Because the services aim for ease of use they are not at all as
'       universal and complet as provided via
'       http://www.cpearson.com/excel/registry.htm.
'
' Public services:
' ----------------
' Delete        Deletes a reg_key with all its sub-keys or only an
'               optional reg_value_name
' Value (Get)   Returns the value of an entry idetified by reg_key and
'               reg_value_namey
' Value (Let)   Writes a reg_value under the provided reg_value_name under
'               the provided reg_key
' Exists        Returns TRUE when the provided reg_key or reg_value_name
'               exists
' Keys          Returns a Dictionary with a Dictionary of values as item
'               and the sub-key as the key (ordered ascending by key)
' Values        Returns a Dictionary of values under a provided key with
'               the value-name as the key and the value as the item.
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
    Any, Source As Any, ByVal numBytes As Long)

Private Const HKEY_CURRENT_USER      As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE     As Long = &H80000002

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

Private Const KEY_READ                  As Long = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                                                           ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                                                           ' SYNCHRONIZE))
Dim oReg                                As Object

Private Sub BoP(ByVal b_proc As String, _
                ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' service. When neither the Common Execution Trace
' Component (mTrc) nor the Common Error Handling Component (mErH) is installed
' (indicated by the Conditional Compile Arguments `mTrc = 1' or `clsTrc = 1` and/or the
' Conditional Compile Argument 'mErH = 1') this procedure does nothing.
' Else the service is handed over to the corresponding procedures.
' May be copied as Private Sub into any module or directly used when mBasic is
' installed.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc component when `mTrc = 1' or `clsTrc = 1`
    '~~ so the Else is only for the case only the mTrc is installed but not the merH.
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub
                        
Public Property Get Value(Optional ByVal v_reg_key As String, _
                          Optional ByVal v_reg_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value of the registry key's (v_reg_key) value identified by its
' name (v_reg_value_name).
' ----------------------------------------------------------------------------
    Const PROC = "Value-Get"
    
    On Error GoTo eh
    Dim rString As String
    
    rString = RegString(v_reg_key, v_reg_value_name)
    If Not IsValidKey(rString) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided v_reg_key | v_reg_value_name ('" & rString & "') lenght exceeds the maximum length of '" & REG_KEY_MAX_LENGTH & "'!"
    If Not HasAccess(v_reg_key, KEY_READ) Then
        ' Err.Raise AppErr(2), ErrSrc(PROC), "No access right to read '" & rString & "'!"
        ' Debug.Print ErrSrc(PROC) & ": " &  "KEY_READ denied for '" & rString & "'"
    End If
    Value = CreateObject("WScript.Shell").RegRead(rString)

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' service. When neither the Common Execution Trace
' Component (mTrc) nor the Common Error Handling Component (mErH) is installed
' (indicated by the Conditional Compile Arguments 'mTrc = 1' and/or the
' Conditional Compile Argument 'mErH = 1') this procedure does nothing.
' Else the service is handed over to the corresponding procedures.
' May be copied as Private Sub into any module or directly used when mBasic is
' installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc component when `mTrc = 1' or `clsTrc = 1`
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf mTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#ElseIf clsTrace = 1 Then
    Trc.EoP e_proc, e_inf
#End If
End Sub

Public Property Let Value(Optional ByVal v_reg_key As String, _
                          Optional ByVal v_reg_value_name As String, _
                                   ByVal v_reg_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (re_value) to the key (v_reg_key) value name (v_reg_value_name)
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
    
    Dim rString As String
    
    rString = RegString(v_reg_key, v_reg_value_name)
    If Not IsValidKey(rString) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The argument for the parameter ""v_reg_key"" is an invalid key string!"
    If Not IsValidValue(v_reg_value) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The the provided argument for the parameter ""v_reg_value"" is invalid! " & _
                                            "It exceeds the maximum allowed value (if numeric) or length if a string)"

    CreateObject("WScript.Shell").RegWrite rString, v_reg_value, RegType(v_reg_value)

End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ArrayIsAllocated(a_arr As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = UBound(a_arr) >= LBound(a_arr, 1)
    On Error GoTo 0
    Err.Clear
    
End Function

Private Sub DctAddAscByKey(ByRef add_dct As Dictionary, _
                           ByVal add_key As Variant, _
                           ByVal add_item As Variant)
' ----------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by
' the key (add_key). When the key is an object with no Name property an error
' is raised.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be
'       copied into any VBProject's module in order to have it independant
'       from this Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ----------------------------------------------------------------------------
    Const PROC = "DctAdd"
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    On Error GoTo eh
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    bOrderByItem = False
    bOrderByKey = True
    bSeqAscending = True
    bCaseIgnored = False
    bCaseSensitive = True
    bStayWithFirst = True
    bEntrySequence = False
    
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not bStayWithFirst Then
            If .Exists(add_key) Then
                If VarType(add_item) = vbObject Then Set .Item(add_key) = add_item Else .Item(add_key) = add_item
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If VarType(add_key) = vbObject Then
            On Error Resume Next
            add_key.Name = add_key.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If VarType(add_item) = vbObject Then
            On Error Resume Next
            add_item.Name = add_item.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = DctAddOrderValue(add_key)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = DctAddOrderValue(.Keys()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If VarType(add_dct.Item(vKeyExisting)) = vbObject _
        Then Set vItemExisting = add_dct.Item(vKeyExisting) _
        Else vItemExisting = add_dct.Item(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = DctAddOrderValue(vKeyExisting)
            
                If vValueExisting = vValueNew And bOrderByItem And bSeqAscending And Not .Exists(add_key) Then
                    If bStayWithFirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and bStayWithFirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function DctAddOrderValue(ByVal dctkey As Variant) As Variant
' ----------------------------------------------------------------------------
' When key is an object its name becomes the order value else the key as is.
' ----------------------------------------------------------------------------
    If VarType(dctkey) = vbObject _
    Then DctAddOrderValue = dctkey.Name _
    Else DctAddOrderValue = dctkey
End Function

Public Function Delete(ByVal d_reg_key As String, _
              Optional ByVal d_reg_value_name As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' When no value name (d_reg_value_name) is provided the provided key is deleted
' else only the name. A possibly missing \ ate the end of the key (d_reg_key)
' ----------------------------------------------------------------------------
    Const PROC = "Delete"
    
    On Error GoTo eh
    Dim rString As String
    
    If d_reg_value_name <> vbNullString Then
        rString = RegString(d_reg_key, d_reg_value_name)
        If Not IsValidKey(rString) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The argument d_reg_key (" & d_reg_key & _
                                                ") and the argument d_reg_value_name (" & d_reg_value_name & _
                                                ") resulted in '" & rString & "' which is not a valid key string!"
        
        
        If Right(rString, 1) = "\" _
        Then rString = Left(rString, Len(rString) - 1)
        CreateObject("WScript.Shell").RegDelete rString ' d_reg_key & d_reg_value_name
        Delete = True
    Else
        mReg.DeleteSubKeys d_reg_key:=d_reg_key
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DeleteSubKeys(ByVal d_reg_key As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "DeleteSubKeys"
    
    On Error GoTo eh
    Dim SubKeys()   As Variant
    Dim SubKey      As Variant
    Dim oReg        As Object
    Static HKey     As Long
    
    RegHKey d_reg_key, HKey   ' Extract and transform hKey and unstrip it frrom d_reg_key
    If Right(d_reg_key, 1) = "\" Then d_reg_key = Left(d_reg_key, Len(d_reg_key) - 1)
    
    Set oReg = GetObject("winmgmts:\\" & "." & "\root\default:StdRegProv")
    oReg.EnumKey HKey, d_reg_key, SubKeys
    If ArrayIsAllocated(SubKeys) Then
        For Each SubKey In SubKeys
            DeleteSubKeys d_reg_key & "\" & SubKey
        Next
    End If
    oReg.DeleteKey HKey, d_reg_key

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
' may be installed (Conditional Compile Argument 'mErH = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'mMsg = 1'). Only when none of the two is installed the error
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
#If mErH = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
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
    If err_source = vbNullString Then err_source = Err.Source
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
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mReg." & sProc
End Function

Public Function Exists(ByVal e_reg_key As String, _
         Optional ByVal e_reg_value_name As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the key (e_reg_key) exists and when a name (e_reg_value_name)
' is provided when the name exists. No worry about \! Missing ones are added.
' ----------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim rString As String
    
    rString = RegString(e_reg_key, e_reg_value_name)
    If Not IsValidKey(rString) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The argument e_reg_key (" & e_reg_key & _
                                            ") and the argument e_reg_value_name (" & e_reg_value_name & _
                                            ") resulted in '" & rString & "' which is not a valid key string!"
    
    If e_reg_value_name <> vbNullString Then
        If Right(e_reg_value_name, 1) = "\" _
        Then e_reg_value_name = Left(e_reg_value_name, Len(e_reg_value_name) - 1)
    End If
    Exists = False
    On Error GoTo xt
    '~~ To have e_reg_key interpreted as key it has to end with a \.
    CreateObject("WScript.Shell").RegRead rString ' e_reg_key & e_reg_value_name
    Exists = True

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Export(ByVal e_reg_key As String) As String
' ----------------------------------------------------------------------------
' Returns a string in the form [section]
'                              name=value
' When the key (e_reg_key) contains sub-keys the returned string will have as
' many [section] lines. The result of the function written to a file is a
' typical .ini, .dat, .cfg file.
' ----------------------------------------------------------------------------
    Dim Keys        As Collection
    Dim key         As Variant
    Dim ValueName   As Variant
    Dim Values      As Dictionary
    
    Set Keys = mReg.Keys(e_reg_key)
    If Keys.Count = 0 Then
        Set Values = mReg.Values(e_reg_key)
        If Values.Count > 0 Then
            If Right(e_reg_key, 1) = "\" Then
                Export = "[" & Split(e_reg_key, "\")(UBound(Split(e_reg_key, "\")) - 1) & "]"
            Else
                Export = "[" & Split(e_reg_key, "\")(UBound(Split(e_reg_key, "\"))) & "]"
            End If
            For Each ValueName In Values
                Export = Export & vbLf & ValueName & "=" & Values(ValueName)
            Next ValueName
        End If
    Else
        For Each key In Keys
            Set Values = mReg.Values(e_reg_key & "\" & key)
            If Values.Count > 0 Then
                If Len(Export) = 0 _
                Then Export = "[" & key & "]" _
                Else Export = Export & vbLf & "[" & key & "]"
                For Each ValueName In Values
                    Export = Export & vbLf & ValueName & "=" & Values(ValueName)
                Next ValueName
            End If
        Next key
    End If
    
End Function

Private Function HasAccess(ByVal h_reg_key As String, ByVal h_reg_access As Long)
' ------------------------------------------------------------------------------
' Returns TRUE when access right (h_reg_access) is granted for the key (h_reg_key).
' ------------------------------------------------------------------------------
    Dim sComputer   As String
    Dim sMethod     As String
    Dim HKey        As Long
    Dim oRegistry   As Object
    Dim oInParam    As Object
    Dim oOutParam   As Object
    Dim oMethod     As Object
    
    sComputer = "."
    sMethod = "CheckAccess"
    RegHKey h_reg_key, HKey ' extract and unstrip the hKey from the h_reg_key
    
    Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}//" & _
            sComputer & "/root/default:StdRegProv")
    
    Set oMethod = oRegistry.Methods_(sMethod)
    Set oInParam = oMethod.inParameters.SpawnInstance_()
    
    oInParam.hDefKey = HKey
    oInParam.sSubKeyName = h_reg_key
    oInParam.uRequired = h_reg_access
    Set oOutParam = oRegistry.ExecMethod_(sMethod, oInParam)
    
    HasAccess = oOutParam.Properties_("bGranted") = 0
End Function

Private Function IsValidKey(ByVal i_reg_key As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the length of the key (i_reg_key) is LE REG_MAX_KEYLENGTH and
' all spaces unstripped longer 0.
' ------------------------------------------------------------------------------
    IsValidKey = (Len(i_reg_key) <= REG_KEY_MAX_LENGTH) And (Len(Trim(i_reg_key)) > 0)
    If IsValidKey Then IsValidKey = Split(i_reg_key, "\")(0) Like "HK*"
End Function

Private Function IsValidValue(Optional ByVal i_reg_value As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when a numeric value (i_reg_value) ranges
' from -2,147,483,648 to 2,147,483,647 or a string - which is the maximum
' because a Registry-Type REG_DWORD stores a Long type value.
' ------------------------------------------------------------------------------
    If IsNumeric(i_reg_value) Then
        IsValidValue = i_reg_value >= REG_VALUE_MIN_VALUE And i_reg_value <= REG_VALUE_MAX_VALUE
    ElseIf VarType(i_reg_value) = vbString Then
        IsValidValue = Len(i_reg_value) <= REG_VALUE_MAX_LENGTH And Len(Trim(i_reg_value)) > 0
    End If
End Function

Public Function Keys(ByVal k_reg_key As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all sub-keys under the key (k_reg_key). The returned
' Dictionary contains an item for each sub-key under the key with the sub-key
' as the key and a Dictionary of all values with the Value-Name as the key.
' ----------------------------------------------------------------------------
    Dim Computer    As String
    Dim SubKeys()   As Variant
    Dim SubKey      As Variant
    Dim HKey        As Long
    Dim oReg        As Object
    
    If Not mReg.Exists(k_reg_key) Then GoTo xt
    RegHKey k_reg_key, HKey ' extract/transform HKey and unstrip it from re_key
    If Right(k_reg_key, 1) = "\" Then k_reg_key = Left(k_reg_key, Len(k_reg_key) - 1)
    Computer = "."

    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Computer & "\root\default:StdRegProv")
    oReg.EnumKey HKey, k_reg_key, SubKeys
    If ArrayIsAllocated(SubKeys) Then
        For Each SubKey In SubKeys
            DctAddAscByKey add_dct:=Keys, add_key:=SubKey, add_item:=mReg.Values(SubKey)
        Next
    End If
xt:
End Function

Private Sub RegHKey(ByRef r_reg_key As String, _
                    ByRef r_reg_hkey As Long)
' ----------------------------------------------------------------------------
' When the key (r_reg_key) begins with a HKEY it is extracted/unstripped and
' transformed to Long.
' The procedure is used wherever the HKey value is required rather than the
' full key as a string. However, any other HKey but HKEY_CURRENT_USER (HKCU)
' or HKEY_LOCAL_MACHINE (or HKML) is invalid and raises an application error.
' ----------------------------------------------------------------------------
    Const PROC = "RegHKey"
    
    On Error GoTo eh
    
    Select Case Split(r_reg_key, "\")(0)
        Case "HKEY_CURRENT_USER":       r_reg_hkey = HKEY_CURRENT_USER:       r_reg_key = Replace(r_reg_key, "HKEY_CURRENT_USER\", vbNullString)
        Case "HKCU":                    r_reg_hkey = HKEY_CURRENT_USER:       r_reg_key = Replace(r_reg_key, "HKCU\", vbNullString)
        Case "HKEY_LOCAL_MACHINE":      r_reg_hkey = HKEY_LOCAL_MACHINE:      r_reg_key = Replace(r_reg_key, "HKEY_LOCAL_MACHINE\", vbNullString)
        Case "HKLM":                    r_reg_hkey = HKEY_LOCAL_MACHINE:      r_reg_key = Replace(r_reg_key, "HKLM\", vbNullString)
        Case Else
'            Err.Raise AppErr(1), ErrSrc(PROC), "The provided key '" & r_reg_key & "' does not begin with " & _
'                                               "HKEY_CURRENT_USER (or HKCU) or HKEY_LOCAL_MACHINE (or HKLM)!"
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function RegString(ByRef r_reg_key As String, ByVal r_reg_value_name As String) As String
    RegString = VBA.Replace(r_reg_key & "\" & r_reg_value_name, "\\", "\")
    RegString = VBA.Replace(RegString, "HKCU", "HKEY_CURRENT_USER")
    RegString = VBA.Replace(RegString, "HKLM", "HKEY_LOCAL_MACHINE")
End Function

Private Function RegType(ByVal r_reg_value As Variant) As String
' ----------------------------------------------------------------------------
' Returns the Registration Type for the value (r_reg_value).
' ----------------------------------------------------------------------------
    Const PROC = "RegType"
    
    On Error GoTo eh
    Select Case VarType(r_reg_value)
        ' Type      To store
        ' --------- --------------------------------
        ' REG_SZ    A string
        ' REG_DWORD A 32-bit number (4 Bytes = Long)
        ' ------------------------------------------
        Case vbBoolean:     RegType = "REG_SZ":     r_reg_value = CStr(Abs(CInt(r_reg_value))) ' reads as 0 or 1
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

Public Function Values(ByVal v_reg_key As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary of all values under a given registry key (v_reg_key)
' where each entry's key is the value name and the item is the value. All
' value-names are sorted in ascending order.
' ----------------------------------------------------------------------------
    Dim dataLen         As Long
    Dim handle          As Long
    Dim Index           As Long
    Dim sName           As String
    Dim ValueNameLength As Long
    Dim resLong         As Long
    Dim resString       As String
    Dim retVal          As Long
    Dim valueType       As Long
    Dim HKey            As Long
    Dim ValueName       As String
    Dim Value           As Variant
    
    Set Values = New Dictionary
       
    '~~ Open the key, exit if not found.
    RegHKey v_reg_key, HKey
    If Len(v_reg_key) Then
        If RegOpenKeyEx(HKey, v_reg_key, 0, KEY_READ, handle) Then Exit Function
        HKey = handle   ' subsequent functions use v_reg_hkey
    End If
    
    Do
        ValueNameLength = 260           ' this is the max length for a key name
        sName = Space$(ValueNameLength)
        '~~ prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        '~~ Read the value's name and data
        retVal = RegEnumValue(HKey, Index, sName, ValueNameLength, ByVal 0&, valueType, resBinary(0), dataLen)
        
        '~~ Enlarge the buffer if more space is required
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(HKey, Index, sName, ValueNameLength, ByVal 0&, valueType, resBinary(0), dataLen)
        End If
        
        If retVal Then Exit Do                      ' exit the loop if any other error (typically, no more values)
        ValueName = Left$(sName, ValueNameLength)    ' retrieve the value's name
        
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
        DctAddAscByKey add_dct:=Values, add_key:=ValueName, add_item:=Value
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

