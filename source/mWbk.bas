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
' Uses:
' - Common Components mErH, fMsg, mMsg (in mWbkTest only!)
'
' Requires: Reference to "Microsoft Scripting Runtine"
'           Reference to "Microsoft Visual Basic for Applications Extensibility ..."
'
' W. Rauschenberger, Berlin Jan 2022
' -----------------------------------------------------------------------------------
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

' --- Error declarations
Private Const ERR_OWB01 = "A Workbook named '<>' is not open in any Excel application instance!"
Private Const ERR_GOW01 = "A Workbook with the provided name (named argument vWb) is open but it's not from requested location ('<>1' instead of '<>2')!"
Private Const ERR_GOW02 = "A Workbook named '<>' (named argument vWb) is not open. The Workbook's full name must be provided to get it opened!"
Private Const ERR_GOW03 = "A Workbook file named '<>' (named argument vWb) does not exist!"

Public Property Get Value(Optional ByVal v_ws As Worksheet, _
                          Optional ByVal v_name As String) As String
    Const PROC = "Value-Get"
    
    On Error Resume Next
    Value = v_ws.Range(v_name).Value
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Worksheet '" & v_ws.Name & "' has no range with a name '" & v_name & "'!"
    
xt: Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Sub WbClose(ByVal c_wb As Variant, _
               Optional ByVal c_save_changes As Boolean = True, _
               Optional ByVal c_save_as_file As String = vbNullString, _
               Optional ByVal c_route_workbook As String = vbNullString)
' ----------------------------------------------------------------------------
' Closes the Workbook (c_wb) - provided either as Workbook object, as Namne
' or FullName provided it is open, i.e. without using On Error Resume Next!
' ----------------------------------------------------------------------------
    Dim wb As Workbook
    If mWbk.IsWbObject(c_wb) Then
        c_wb.Close
    ElseIf mWbk.IsName(c_wb) Or mWbk.IsFullName(c_wb) Then
        If mWbk.IsOpen(c_wb, wb) Then
            wb.Close c_save_changes, c_save_as_file, c_route_workbook
        End If
    End If
End Sub

Public Property Let Value(Optional ByVal v_ws As Worksheet, _
                          Optional ByVal v_name As String, _
                                   ByVal v_value As String)
    Const PROC = "Value-Let"
    
    On Error Resume Next
    v_ws.Range(v_name).Value = v_value

    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Worksheet '" & v_ws.Name & "' has no range with a name '" & v_name & "'!"
    
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
    ErrSrc = "mWbk" & "." & sProc
End Function

Public Function Exists(ByVal ex_wb As Variant, _
              Optional ByVal ex_ws As Variant = Nothing, _
              Optional ByVal ex_range_name As String = vbNullString, _
              Optional ByRef ex_result_wbk As Workbook, _
              Optional ByRef ex_result_wsh As Worksheet, _
              Optional ByRef ex_result_rng As Range) As Boolean
' ----------------------------------------------------------------------------
' Universal existence check for Workbook, Worksheet, and Range Name.
' Returns TRUE when the Workbook - which may be a Workbook's name or FullName
' exists and:
' - the Worksheet (ex_ws) and the range name (ex_range_name) = vbNullString
' - the Worksheet (ex_ws) is provided - either by its name or its code name
'   and exists in the Workbook (ex_wb) which is open! and the range name
'   (ex_range_name) = vbNullString
' - the Worksheet = vbNullString and the range name (ex_range_name) exists
'   in the Workbook - regardless of the sheet
' - the Worksheet (ex_ws) exists and the range name (ex_range_name) refers
'   to a range in it.
' Error conditions:
' - AppErr(1) when the Workbook is provided as Name '....,xl*' and is not open
' - AppErr(2) when the Workbook is not open and a Worksheet or range name is
'   provided
' ----------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim sTest       As String
    Dim wsTest      As Worksheet
    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim fso         As New FileSystemObject
    Dim nm          As Name
    Dim sWsName     As String
    Dim sWsCodeName As String
    
    If IsFullName(ex_wb) Then
        Exists = fso.FileExists(ex_wb)
    ElseIf IsName(ex_wb) Then
        If Not mWbk.IsOpen(ex_wb, wb) Then
            Err.Raise AppErr(1), ErrSrc(PROC), "The existence of a Workbook provided just by its file name (" & ex_wb & ") " & _
                                               "cannot be checked when the Workbook is closed. " & _
                                               "To check a not open Workbook's existence requires its full name!"
        End If
        Exists = True
    ElseIf mWbk.IsWbObject(ex_wb) Then
        Set wb = ex_wb
        Exists = True
    End If
    If Not Exists Then GoTo xt
    
    If Not TypeName(ex_ws) = "Nothing" Then
        If wb Is Nothing Then
            Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook '" & ex_wb & "' exists but is not open. " & _
                                               "The existence of a Worksheet cannot be checked for a Workbook which not open."
        End If
        If IsWsObject(ex_ws) _
        Then sTest = ex_ws.Name _
        Else sTest = ex_ws

        For Each ws In wb.Worksheets
            Exists = ws.Name = sTest
            If Not Exists Then Exists = ws.CodeName = sTest
            If Exists Then
                sWsName = ws.Name
                sWsCodeName = ws.CodeName
                Exit For
            End If
        Next ws
        If Not Exists Then GoTo xt
    End If
        
    If ex_range_name <> vbNullString Then
        If ws Is Nothing Then
            '~~ Check if the range name is one in the Workbook
            For Each nm In wb.Names
                Exists = nm.Name = ex_range_name
                If Exists Then Exit For
            Next nm
        Else
            '~~ Check if the name refers to a range in the provided Worksheet
            For Each nm In wb.Names
                Exists = nm.Name = ex_range_name
                If Exists Then Exists = nm.RefersTo Like "*" & sWsName & "*"
                If Exists Then Exit For
            Next nm
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

Public Function GetOpen(ByVal vWb As Variant) As Workbook
' -------------------------------------------------------
' Returns an open Workbook object or raises an error.
' If vWb is a full path-file name, the file exists but
' is not open it is opened.
' Note: A ReadOnly mode has to be set by the caller.
' -------------------------------------------------------
    Const PROC = "GetOpen"
    
    On Error GoTo eh
    Dim sTest       As String
    Dim sWbBaseName As String
    Dim sPath       As String
    Dim wbOpen      As Workbook
    Dim fso         As New FileSystemObject
        
    Set wbOpen = Nothing
    
    If Not mWbk.IsName(vWb) And Not mWbk.IsFullName(vWb) And Not mWbk.IsWbObject(vWb) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter vWb) is neither a Workbook object nor a string (name or fullname)!"
    If mWbk.IsWbObject(vWb) _
    Then sWbBaseName = fso.GetBaseName(vWb.FullName) _
    Else sWbBaseName = fso.GetBaseName(vWb)

    If mWbk.IsWbObject(vWb) Then
        Set wbOpen = vWb
    ElseIf mWbk.IsFullName(vWb) Then
        With Opened
            If fso.FileExists(sWbBaseName) Then
                '~~ A Workbook with the same name is open
                Set wbOpen = .Item(sWbBaseName)
                If wbOpen.FullName <> vWb Then
                    '~~ The open Workook with the same name is from a different location
                    If fso.FileExists(vWb) Then
                        '~~ When the Workbook file still exists at the provided location the one which is open is the wromg one
                        Err.Raise AppErr(3), ErrSrc(PROC), Replace(Replace$(ERR_GOW01, "<>1", wbOpen.Path), "<>2", sPath)
                    Else
                        '~~ The Workbook file does not or no longer exist at the provivded location.
                        '~~ The open one is apparenty the ment Workbook just moved to the new location.
                        Set wbOpen = wbOpen
                    End If
                Else
                    '~~ The open Workook is the one indicated by the provided full name
                    Set wbOpen = wbOpen
                End If
            Else
                '~~ The Workbook is yet not open
                If fso.FileExists(vWb) Then
                    Set wbOpen = Workbooks.Open(vWb)
                Else
                    Err.Raise AppErr(4), ErrSrc(PROC), Replace(ERR_GOW03, "<>", CStr(vWb))
                End If
            End If
        End With
    ElseIf mWbk.IsName(vWb) Then
        With Opened
            If .Exists(sWbBaseName) Then
                Set wbOpen = .Item(sWbBaseName)
            Else
                Err.Raise AppErr(5), ErrSrc(PROC), "A Workbook named '" & sWbBaseName & "' is not open and it cannot be opened since only the name is provided (a full name would be required)!"
            End If
        End With
    End If
    Set GetOpen = wbOpen
    
xt: Exit Function
    
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

Public Function IsOpen(ByVal wb As Variant, _
              Optional ByRef wb_result As Workbook) As Boolean
' -----------------------------------------------------------------------------------
' Returns TRUE when the Workbook (wb) - which may be a Workbook object, a Workbook's
' name or fullname - is open in any Excel Application instance. If a fullname is
' provided and the file does not exist under this full name but a Workbook with the
' given name is open (but from another folder) the Workbook is regarded moved and
' thus is returned as open object(wb_result).
' Because Workbooks with the same WbName may be open when they have different
' extensions a Workbook's Name including its extension is checked.
' -----------------------------------------------------------------------------------
    Const PROC = "IsOpen"
    
    On Error GoTo eh
    Dim OpenWbks As Dictionary
    Dim OpenWbk  As Workbook
    Dim fso      As New FileSystemObject
    Dim WbName As String
    
    If Not mWbk.IsWbObject(wb) And Not mWbk.IsFullName(wb) And Not mWbk.IsName(wb) And Not TypeName(wb) = "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter wb) is neither a Workbook object nor a Workbook's name or fullname)!"
       
    Set OpenWbks = mWbk.Opened
    If mWbk.IsName(wb) Then
        '~~ wb is a Workbook's Name including its extension
        WbName = fso.GetFileName(wb)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ When a Workbook's Name is provided the Workbook is only regarde open when the open
            '~~ Workbook has the same name (i.e. including its extension)
            If fso.GetFile(OpenWbk.FullName).Name <> fso.GetFileName(wb) Then Set OpenWbk = Nothing
        End If
    ElseIf mWbk.IsFullName(wb) Then
        WbName = fso.GetFileName(wb)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ The provided (wb) specifies an exist Workbook file. This Workbook is regarded open (and returned as opject)
            '~~ when a Workbook with its Name (including the extension!) is open regardless in which location
            If fso.GetFile(OpenWbk.FullName).Name <> fso.GetFileName(wb) Then Set OpenWbk = Nothing
        End If
    ElseIf mWbk.IsWbObject(wb) Then
        WbName = wb.Name
        If Opened.Exists(WbName) Then
            Set OpenWbk = OpenWbks.Item(WbName)
        End If
    Else
        '~~ If wb is a Workbook's WbName it is regarded open when one with that WbName is open
        '~~ regrdless its extension
        If OpenWbks.Exists(wb) Then Set OpenWbk = OpenWbks.Item(wb)
    End If
    
xt: If mWbk.IsWbObject(OpenWbk) Then
        IsOpen = True
        Set wb_result = OpenWbk
    End If
    Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    Dim fso     As New FileSystemObject
    Dim lApps   As Long
    Dim wbk     As Workbook
    Dim aApps() As Application ' Array of currently active Excel applications
    Dim app     As Variant
    Dim dct     As New Dictionary
    Dim i       As Long

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

Private Function TestSheet(ByVal wb As Workbook, _
                           ByVal vWs As Variant) As Worksheet
' -----------------------------------------------------------
' Returns the Worksheet object (vWs) - which may be a Work-
' sheet object or a Worksheet's name - of the Workbook (wb).
' Precondition: The Worksheet exists.
' -----------------------------------------------------------
    If VarType(vWs) = vbString Then
        Set TestSheet = wb.Worksheets(vWs)
    ElseIf TypeOf vWs Is Worksheet Then
        Set TestSheet = vWs
    End If
End Function

