Attribute VB_Name = "mWbk"
Option Explicit
Option Private Module
Option Compare Text
' -----------------------------------------------------------------------------------
' Standard Module mWbk: Provides basic Workbook services.
'
' Public services:
' - GetOpen     Opens a provided Workbook if possible, returns the Workbook object of
'               the openend or an already open Workbook
' - IsFullName  Returns TRUE when a provided string is the full name of a Workbook
' - IsName      Returns TRUE when a provided string is the name of a Workbook
' - IsWbObject    Returns TRUE when the provided variant is a Workbook (not necessarily
'               also open!)
' - IsOpen      Returns TRUE when the provided Workbook is open
' - Opened      Returns a Dictionary of all open Workbooks in any application
'               instance with the Workbook's Name as the key and the Workbook object
'               as the item.
'
' Uses:         mErH, fMsg, mMsg, mTrc (in mWbkTest module only!)
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' See: https://github.com/warbe-maker/Common-VBA-Excel-Workbook-Services
'
' W. Rauschenberger, Berlin June 2022
' -----------------------------------------------------------------------------------
#If Not MsgComp = 1 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed which must
    ' be indicated by the Conditional Compile Argument MsgComp = 1.
    ' See https://github.com/warbe-maker/Common-VBA-Message-Service
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

' --- Begin of declarations to get all Workbooks of all running Excel instances
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As LongPtr, ByVal dwId As LongPtr, ByRef riid As UUID, ByRef ppvObject As Object) As LongPtr

Type UUID 'GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Private Const OBJID_NATIVEOM As LongPtr = &HFFFFFFF0
' --- End of declarations to get all Workbooks of all running Excel instances

Public Property Get Value(Optional ByVal v_wsh As Worksheet, _
                          Optional ByVal v_name As Variant) As String
' ----------------------------------------------------------------------------
' Returns a Value from a Worksheet (v_wsh) identified by (v_name) which may be
' a Range-Name or a Range.
' ----------------------------------------------------------------------------
    Const PROC = "Value-Get"
    
    On Error Resume Next
    Select Case TypeName(v_name)
        Case "String":  Value = v_wsh.Range(v_name).Value
        Case "Range":   Value = v_name.Value
        Case Else:      Err.Raise AppErr(1), ErrSrc(PROC), "The argument 'v_name is neither a string (RangeName) nor a Range!"
    End Select
    If Err.Number <> 0 _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Worksheet '" & v_wsh.Name & "' has no range with a name '" & v_name & "'!"
    
xt: Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Value(Optional ByVal v_wsh As Worksheet, _
                          Optional ByVal v_name As Variant, _
                                   ByVal v_value As String)
' ----------------------------------------------------------------------------
' Saves a Value to a Werksheet (v_wsh) identified by (v_name) which may be a
' Range-Name or a Range.
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
    
    Dim rng As Range
    
    On Error Resume Next
    Select Case TypeName(v_name)
        Case "String": Set rng = v_wsh.Range(v_name)
        Case "Range":  Set rng = v_name
    End Select
    If Err.Number <> 0 _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Worksheet '" & v_wsh.Name & "' has no range with a name '" & v_name & "'!"
    
    If v_wsh.ProtectContents And rng.Locked Then
        On Error Resume Next
        v_wsh.Unprotect
        If Err.Number <> 0 _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "The Worksheet '" & v_wsh.Name & "' is apparently password protected which is not supported by the Value service!"
        rng.Value = v_value
        v_wsh.Protect
    Else
        rng.Value = v_value
    End If
    
xt: Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Private Function checkHwnds(ByRef xlApps() As Application, hWnd As LongPtr) As Boolean
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "checkHwnds"            ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim i As Long
    
    If UBound(xlApps) = 0 Then GoTo xt

    For i = LBound(xlApps) To UBound(xlApps)
        If xlApps(i).hWnd = hWnd Then
            checkHwnds = False
            GoTo xt
        End If
    Next i
    checkHwnds = True
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrMsg(ByVal err_source As String, _
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
    ErrSrc = "mWbk" & "." & sProc
End Function

Public Function Exists(ByVal ex_wbk As Variant, _
              Optional ByVal ex_wsh As Variant = Nothing, _
              Optional ByVal ex_range_name As String = vbNullString, _
              Optional ByRef ex_result_wbk As Workbook, _
              Optional ByRef ex_result_wsh As Worksheet, _
              Optional ByRef ex_result_rng As Range) As Boolean
' ----------------------------------------------------------------------------
' Universal existence check for Workbook, Worksheet, and Range Name.
' Returns TRUE when the Workbook - which may be a Workbook's name or FullName
' exists and:
' - the Worksheet (ex_wsh) and the range name (ex_range_name) = vbNullString
' - the Worksheet (ex_wsh) is provided - either by its name or its code name
'   and exists in the Workbook (ex_wbk ) which is open! and the range name
'   (ex_range_name) = vbNullString
' - the Worksheet = vbNullString and the range name (ex_range_name) exists
'   in the Workbook - regardless of the sheet
' - the Worksheet (ex_wsh) exists and the range name (ex_range_name) refers
'   to a range in it.
' Error conditions:
' - AppErr(1) when the Workbook is provided as Name '....,xl*' and is not open
' - AppErr(2) when the Workbook is not open and a Worksheet or range name is
'   provided
' ----------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim sTest       As String
    Dim wbk         As Workbook
    Dim wsh         As Worksheet
    Dim fso         As New FileSystemObject
    Dim nme         As Name
    Dim sWsName     As String
    Dim sWsCodeName As String
    
    If IsFullName(ex_wbk) Then
        Exists = fso.FileExists(ex_wbk)
    ElseIf IsName(ex_wbk) Then
        If Not mWbk.IsOpen(ex_wbk, wbk) Then
            Err.Raise AppErr(1), ErrSrc(PROC), "The existence of a Workbook provided just by its file name (" & ex_wbk & ") " & _
                                               "cannot be checked when the Workbook is closed. " & _
                                               "To check a not open Workbook's existence requires its full name!"
        End If
        Exists = True
        Set ex_result_wbk = wbk
    ElseIf mWbk.IsWbObject(ex_wbk) Then
        Set wbk = ex_wbk
        Set ex_result_wbk = wbk
        Exists = True
    End If
    If Not Exists Then GoTo xt
    
    If Not TypeName(ex_wsh) = "Nothing" Then
        If wbk Is Nothing Then
            Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook '" & ex_wbk & "' exists but is not open. " & _
                                               "The existence of a Worksheet cannot be checked for a Workbook which not open."
        End If
        If IsWsObject(ex_wsh) _
        Then sTest = ex_wsh.Name _
        Else sTest = ex_wsh

        For Each wsh In wbk.Worksheets
            Exists = wsh.Name = sTest
            If Not Exists Then Exists = wsh.CodeName = sTest
            If Exists Then
                sWsName = wsh.Name
                sWsCodeName = wsh.CodeName
                Exit For
            End If
        Next wsh
        If Not Exists Then GoTo xt
        Set ex_result_wsh = wsh
    End If
        
    If ex_range_name <> vbNullString Then
        If wsh Is Nothing Then
            '~~ Check if the range name is one in the Workbook
            For Each nme In wbk.Names
                Exists = nme.Name = ex_range_name
                If Exists Then
                    Set ex_result_rng = ex_result_wsh.Range(ex_range_name)
                    Exit For
                End If
            Next nme
        Else
            '~~ Check if the name refers to a range in the provided Worksheet
            For Each nme In wbk.Names
                Exists = nme.Name = ex_range_name
                If Exists Then Exists = nme.RefersTo Like "*" & sWsName & "*"
                If Exists Then
                    Set ex_result_rng = ex_result_wsh.Range(ex_range_name)
                    Exit For
                End If
            Next nme
        End If
    End If
            
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function GetExcelObjectFromHwnd( _
                                  ByVal hWndMain As LongPtr) As Application
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    Const PROC = "GetExcelObjectFromHwnd"

#If Win64 Then
    Dim hWndDesk As LongPtr
    Dim hWnd As LongPtr
#Else
    Dim hWndDesk As Long
    Dim hWnd As Long
#End If
    
    On Error GoTo eh
    Dim sText   As String
    Dim lRet    As Long
    Dim iid     As UUID
    Dim ob      As Object
    
    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString)

    If hWndDesk <> 0 Then
        hWnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString)

        Do While hWnd <> 0
            sText = String$(100, Chr$(0))
            lRet = CLng(GetClassName(hWnd, sText, 100))
            If Left$(sText, lRet) = "EXCEL7" Then
                Call IIDFromString(StrPtr(IID_IDispatch), iid)
                If AccessibleObjectFromWindow(hWnd, OBJID_NATIVEOM, iid, ob) = 0 Then 'S_OK
                    Set GetExcelObjectFromHwnd = ob.Application
                    GoTo xt
                End If
            End If
            hWnd = FindWindowEx(hWndDesk, hWnd, vbNullString, vbNullString)
        Loop
        
    End If
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function GetOpen(ByVal go_wbk As Variant, _
               Optional ByVal go_read_only As Boolean = False) As Workbook
' ----------------------------------------------------------------------------
' Returns an open Workbook object or raises an error. If go_wbk  is a full path-file name, the file exists but
' is not open it is opened.
' Note: A ReadOnly mode has to be set by the caller.
' ----------------------------------------------------------------------------
    Const PROC = "GetOpen"
    
    On Error GoTo eh
    Dim sWbName     As String
    Dim sWbFullName As String
    Dim wbOpen      As Workbook
    Dim fso         As New FileSystemObject
        
    Set wbOpen = Nothing
       
    If mWbk.IsWbObject(go_wbk) Then
        sWbName = go_wbk.Name
        sWbFullName = go_wbk.FullName
    ElseIf mWbk.IsFullName(go_wbk) Then
        sWbName = fso.GetFileName(go_wbk)
        sWbFullName = go_wbk
    ElseIf mWbk.IsName(go_wbk) Then
        sWbName = go_wbk
        sWbFullName = vbNullString
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter go_wbk ) is neither a Workbook object nor a string (name or fullname)!" & vbLf & _
                                           "(TypeName of argument = '" & TypeName(go_wbk) & "'!)"
    End If

    If mWbk.IsWbObject(go_wbk) Then
        Set GetOpen = go_wbk
        GoTo xt
    End If
    
    With mWbk.Opened
        If .Exists(sWbName) Then
            Set GetOpen = .Item(sWbName)
            If sWbFullName = vbNullString Then GoTo xt
            Set wbOpen = .Item(sWbName)
            If wbOpen.FullName = sWbFullName Then GoTo xt
            '~~ A Workbook with the Name is open but it has a different FullName
            If fso.FileExists(sWbFullName) Then
                '~~ When the Workbook file still exists at the provided location the one which is open is the wromg one
                Err.Raise AppErr(2), ErrSrc(PROC), "A Workbook named '" & sWbName & "' is open but its location differs." & vbLf & vbLf & _
                                                   "'" & wbOpen.FullName & "'" & vbLf & vbLf & _
                                                   "instead of the requested" & vbLf & vbLf & _
                                                   "'" & sWbFullName & "'"
            End If
            GoTo xt
        Else
            '~~ A Workbook with the provided name is yet not open
            If sWbFullName = vbNullString _
            Then Err.Raise AppErr(3), ErrSrc(PROC), "A Workbook named '" & sWbName & "' is not open - and cannot be opened because this requires the full file name!"
            If Not fso.FileExists(sWbFullName) _
            Then Err.Raise AppErr(4), ErrSrc(PROC), "A Workbook named '" & sWbFullName & "' does not exist!"
            Set GetOpen = Workbooks.Open(sWbFullName, , go_read_only)
            GoTo xt
        End If
    End With
        
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsFullName(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is a Workbook's FullName - not necessarily existing.
' ----------------------------------------------------------------------------
    With New FileSystemObject
        If VarType(v) = vbString Then
            IsFullName = mWbk.IsName(.GetFileName(v)) And (InStr(v, "\") <> 0 Or InStr(v, "/") <> 0)
        End If
    End With
End Function

Public Function IsName(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when (v) is a Workbook's Name (without path!) .
' ----------------------------------------------------------------------------
    If VarType(v) = vbString Then
        IsName = InStr(v, "\") = 0 And InStr(v, "/") = 0 And v Like "*.xl*"
    End If
End Function

Public Function IsOpen(ByVal wbk As Variant, _
              Optional ByRef wbk_result As Workbook) As Boolean
' -----------------------------------------------------------------------------------
' Returns TRUE when the Workbook (wbk) - which may be a Workbook object, a Workbook's
' name or fullname - is open in any Excel Application instance. If a fullname is
' provided and the file does not exist under this full name but a Workbook with the
' given name is open (but from another folder) the Workbook is regarded moved and
' thus is returned as open object(wbk_result).
' Because Workbooks with the same WbName may be open when they have different
' extensions a Workbook's Name including its extension is checked.
' -----------------------------------------------------------------------------------
    Const PROC = "IsOpen"
    
    On Error GoTo eh
    Dim OpenWbks As Dictionary
    Dim OpenWbk  As Workbook
    Dim fso      As New FileSystemObject
    Dim WbName As String
    
    If Not mWbk.IsWbObject(wbk) And Not mWbk.IsFullName(wbk) And Not mWbk.IsName(wbk) And Not TypeName(wbk) = "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter wbk) is neither a Workbook object nor a Workbook's name or fullname)!"
       
    Set OpenWbks = mWbk.Opened
    If mWbk.IsName(wbk) Then
        '~~ wbk is a Workbook's Name including its extension
        WbName = fso.GetFileName(wbk)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ When a Workbook's Name is provided the Workbook is only regarde open when the open
            '~~ Workbook has the same name (i.e. including its extension)
            If fso.GetFile(OpenWbk.FullName).Name <> fso.GetFileName(wbk) Then Set OpenWbk = Nothing
        End If
    ElseIf mWbk.IsFullName(wbk) Then
        WbName = fso.GetFileName(wbk)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ The provided (wbk) specifies an exist Workbook file. This Workbook is regarded open (and returned as opject)
            '~~ when a Workbook with its Name (including the extension!) is open regardless in which location
            If fso.GetFile(OpenWbk.FullName).Name <> fso.GetFileName(wbk) Then Set OpenWbk = Nothing
        End If
    ElseIf mWbk.IsWbObject(wbk) Then
        WbName = wbk.Name
        If Opened.Exists(WbName) Then
            Set OpenWbk = OpenWbks.Item(WbName)
        End If
    Else
        '~~ If wbk is a Workbook's WbName it is regarded open when one with that WbName is open
        '~~ regrdless its extension
        If OpenWbks.Exists(wbk) Then Set OpenWbk = OpenWbks.Item(wbk)
    End If
    
xt: If mWbk.IsWbObject(OpenWbk) Then
        IsOpen = True
        Set wbk_result = OpenWbk
    End If
    Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsWbObject(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is a valid Workbook object.
' ----------------------------------------------------------------------------
    If VarType(v) = vbObject Then
        IsWbObject = TypeName(v) = "Workbook"
    End If
End Function

Public Function IsWsObject(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is a valid Workbook object.
' ----------------------------------------------------------------------------
    If VarType(v) = vbObject Then
        IsWsObject = TypeName(v) = "Worksheet"
    End If
End Function

Public Function Opened() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary of all currently open Workbooks in any running excel
' application instance with the Workbook's name (including its extension!)
' as the key and the Workbook as item.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Opened"
    
    On Error GoTo eh
#If Win64 Then
    Dim hWndMain As LongPtr
#Else
    Dim hWndMain As Long
#End If
    Dim lApps   As Long
    Dim wbk     As Workbook
    Dim aApps() As Application ' Array of currently active Excel applications
    Dim app     As Variant
    Dim dct     As New Dictionary

    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString)
    lApps = 0

    '~~ Collect all runing Excel instances in the array aApps
    Do While hWndMain <> 0
        Set app = GetExcelObjectFromHwnd(hWndMain)
        If Not (app Is Nothing) Then
            If lApps = 0 Then
                lApps = 1
                ReDim aApps(1 To 1)
                Set aApps(lApps) = app
            ElseIf checkHwnds(aApps, app.hWnd) Then
                lApps = lApps + 1
                ReDim Preserve aApps(1 To lApps)
                Set aApps(lApps) = app
            End If
        End If
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString)
    Loop

    '~~ Collect all open Workbooks in a Dictionary and return it
    With dct
        .CompareMode = TextCompare
        For Each app In aApps
            For Each wbk In app.Workbooks
                If Not .Exists(wbk.Name) Then .Add wbk.Name, wbk
            Next wbk
        Next app
    End With
    Set Opened = dct

xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub WbClose(ByVal c_wbk As Variant, _
               Optional ByVal c_save_changes As Boolean = True, _
               Optional ByVal c_save_as_file As String = vbNullString, _
               Optional ByVal c_route_workbook As String = vbNullString)
' ----------------------------------------------------------------------------
' Closes the Workbook (c_wbk ) - provided either as Workbook object, as Namne
' or FullName provided it is open, i.e. without using On Error Resume Next!
' ----------------------------------------------------------------------------
    Dim wbk As Workbook
    If mWbk.IsWbObject(c_wbk) Then
        c_wbk.Close
    ElseIf mWbk.IsName(c_wbk) Or mWbk.IsFullName(c_wbk) Then
        If mWbk.IsOpen(c_wbk, wbk) Then
            wbk.Close c_save_changes, c_save_as_file, c_route_workbook
        End If
    End If
End Sub

