Attribute VB_Name = "mRng"
Option Explicit
Option Compare Text
' ------------------------------------------------------------------------------
' Standard  Module mRng: Provides services for ranges.
' ======================
'
' Public services:
' ----------------
' Exists          Returns TRUE when the range exists
' FoundInFormulas Returns TRUE when a provided string is found in any formula.
'                 When a Worksheet is provided, only in this else in all
'                 Worksheets in the provided Workbook. When a Collection is
'                 provided all cells with a formula containing the provided
'                 string are returned, else the function ends with TRUE with
'                 the first found.
' FlipStages      Changes the state (a cell's displayed character string) to
'                 the consequtive string, optionally continued with the first
'                 when the last is reached.
' FromArray       not yet implemented
' HasNames        Returns a Dictionary with all Names referring to a certain
'                 range.
' HasUrl
' Raised          Mimiks a cell as a not clicked button
' Sunken          Mimiks a cell as a clicked button
' UrlGet          Returns the hyperling of a provided range/cell
' UrlLet          Inserts a Hyperlink into a provided range/cell
'
' Uses:
' -----
' No other modules
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' W. Rauschenberger, Berlin Dec 2022
' ------------------------------------------------------------------------------
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

Private Const GITHUB_REPO_URL = "https://github.com/warbe-maker/VBA-Excel-Range"

' Begin of ShellRun declarations ---------------------------------------------
Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long
' Window Constants
Private Const WIN_NORMAL = 1         'Open Normal
' ShellRun Error Codes
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const GWL_STYLE As Long = -16
' End of ShellRun declarations ---------------------------------------------

Public Property Get Url(ByVal u_rng As Range) As String
' ------------------------------------------------------------------------------
' Returns the "Full" Hyperlink Address for the range (u_rng) even when the
' system just provides a relative address. When the range (u_rng) does not
' contain a hyperlink a vbNullString is returned.
' ------------------------------------------------------------------------------
    Const PROC = "Url"
    
    Dim FSo         As New FileSystemObject
    Dim wbk         As Workbook
    Dim sUrl        As String
    Dim sSplit      As String
    Dim aUrlHost    As Variant
    Dim aUrlPath    As Variant
    Dim i           As Long
    Dim sPath       As String
    
    Set wbk = u_rng.Parent.Parent
    sPath = wbk.Path
    FSo.GetParentFolderName (sPath)
    
    aUrlHost = Split(wbk.FullName, "\")
    
    If u_rng.Hyperlinks.Count = 1 Then
        sUrl = u_rng.Hyperlinks(1).Address
        If InStr(sUrl, "/") <> 0 Then sSplit = "/" Else sSplit = "\"
        aUrlPath = Split(sUrl, sSplit)
        If UBound(aUrlPath) = 0 Then
            sUrl = aUrlHost(LBound(aUrlHost))
            For i = LBound(aUrlHost) + 1 To UBound(aUrlHost) - 1
                sUrl = sUrl & sSplit & aUrlHost(i)
            Next i
            Url = sUrl & sSplit & aUrlPath(0)
        ElseIf aUrlPath(0) = ".." Then
            '~~ Handle ..
            sUrl = vbNullString
            For i = UBound(aUrlPath) To LBound(aUrlPath) Step -1
                If aUrlPath(i) = ".." Then
                    sUrl = sUrl & sSplit & FSo.GetParentFolderName(sPath)
                    sPath = FSo.GetParentFolderName(sPath)
                End If
            Next i
            Debug.Print ErrSrc(PROC) & ": " & sUrl
            
            For i = LBound(aUrlPath) + 1 To UBound(aUrlPath)
                If aUrlPath(i) <> ".." Then
                    sUrl = sUrl & sSplit & aUrlPath(i)
                End If
            Next i
            Url = sUrl
        Else
            If u_rng.Hyperlinks(1).SubAddress <> vbNullString Then
                Url = sUrl & "#" & u_rng.Hyperlinks(1).SubAddress
            Else
                Url = sUrl
            End If
        End If
    End If
    If Left(Url, 1) = sSplit Then Url = Right(Url, Len(Url) - 1)
    
End Property

Private Function AllWbksOpen() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary of all currently open Workbooks in any running excel
' application instance with the Workbook's name (including its extension!)
' as the key and the Workbook as item.
' ----------------------------------------------------------------------------
    Const PROC  As String = "AllWbksOpen"
    
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
    Set AllWbksOpen = dct

xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    Const README_URL = "/blob/master/README.md"
    
    If r_bookmark = vbNullString _
    Then ShellRun GITHUB_REPO_URL & README_URL _
    Else ShellRun GITHUB_REPO_URL & README_URL & "#" & r_bookmark
        
End Sub

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

Private Function ShellRun(ByVal sr_string As String, _
                 Optional ByVal sr_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Public Sub BoP(ByVal b_proc As String, _
      Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.BoP b_proc, b_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

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

Public Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRng" & "." & sProc
End Function

Public Function Exists(ByVal e_wbk As Variant, _
                       ByVal e_wsh As Variant, _
                       ByVal e_rng As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the provided Range (e_rng) - which may be range object or a
' range's name - exists in the provided Worksheet (e_wsh) - which may be a
' sheet's name or a Worksheet object - of the provided Workbook (e_wbk) - which
' may be a Workbbok's Name or Full-Name or a Workbook object.
' ------------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim sTest       As String
    Dim wsh         As Worksheet
    Dim wbk         As Workbook
    Dim rng         As Range
    Dim wsResult    As Worksheet

    BoP ErrSrc(PROC)
    
    Exists = False
    Set wbk = e_wbk   ' raises an error when not open
    
    If TypeName(e_rng) = "Nothing" _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Range (parameter e_rng) for the Range's existence check is ""Nothing""!"
    
    If TypeOf e_rng Is Range Then
        Set rng = e_rng
        On Error Resume Next
        sTest = rng.Address
        Exists = Err.Number = 0
    ElseIf VarType(e_rng) = vbString Then
        If Not SheetExists(wbk, e_wsh, wsResult) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The Worksheet (parameter vWs) for the Range's existence check does not exist in Workbook (e_wbk)!"
        Set wsh = wsResult
        On Error Resume Next
        sTest = wsh.Range(e_rng).Address
        Exists = Err.Number = 0
    End If
            
xt: EoP ErrSrc(PROC)
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

'Public Sub FromArray(ByVal r As Range, _
'                     ByVal vArr As Variant)
'' -----------------------------------------
'' Returns the range (r) filled with the
'' content of array (vArr).
'' -----------------------------------------
'
'End Sub

Public Sub FlipStages(ByVal fs_cell As Range, _
                      ByVal fs_round_robin As Boolean, _
                      ByVal fs_select_next As Variant, _
                 ParamArray fs_stages() As Variant)
' ------------------------------------------------------------------------------
' Changes the character of a cell (the first item in the array) to the one next
' to the current character, when the last is reached to the first one.
'
' Error handling is with the caller!
' ------------------------------------------------------------------------------
    Const PROC = "FlipStages"
    
    Dim i As Long
    Dim sCurrent    As String
    Dim sNext       As String
    Dim wsh         As Worksheet
    Dim bEmptyOk    As Boolean
    Dim celLast     As Range
    Dim bEvents     As Boolean
    
    Application.EnableEvents = False
    If fs_cell.Cells.Count <> 1 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided range is not a single cell!"
    Set wsh = fs_cell.Parent
    
    If UBound(fs_stages) < 1 _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided argument does not have at least two 'flip-items'!"
        
    If Not fs_select_next Is Nothing Then
        If Not TypeOf fs_select_next Is Range _
        Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided 'fs_select_next' argument is neither Nothing nor a range object!"
    End If
    
    For i = 0 To UBound(fs_stages)
        If fs_stages(i) = fs_cell.Value Then
            If fs_stages(i) = vbNullString Then bEmptyOk = True
            sCurrent = fs_stages(i)
            If i = UBound(fs_stages) Then
                If fs_round_robin _
                Then sNext = fs_stages(0) _
                Else sNext = sCurrent
            Else:    sNext = fs_stages(i + 1)
            End If
            Exit For
        End If
    Next i
    If Not bEmptyOk And sCurrent = vbNullString _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The cells current value is invalid! (not a known 'stage-value' provided by the fs_stage argument)"
    
    If wsh.ProtectContents Then
        '~~ When the Worksheet is protected allowing to select locked cells
        wsh.Unprotect
        fs_cell.Value = sNext
        wsh.Protect
    Else
        fs_cell.Value = sNext
    End If
    
    '~~ De-Select in order to allow a subsequent stage flip
    If fs_select_next Is Nothing Then
        Set celLast = wsh.Cells.SpecialCells(xlCellTypeLastCell)
        celLast.Offset(1, 1).Select
    ElseIf TypeOf fs_select_next Is Range Then
        fs_select_next.Select
    End If
    
    Application.EnableEvents = bEvents
    
End Sub

Public Function FoundInFormulas(ByVal fif_str As String, _
                                ByVal fif_wbk As Workbook, _
                       Optional ByVal fif_wsh As Worksheet = Nothing, _
                       Optional ByRef fif_cll As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the provided string (fit_str) is found in a formula. When a
' Worksheet (fif_wsh) is provided, only in this else in all Worksheets in the
' provided Workbook (fif_wbk). When a Collection (fif_cll) is provided all cells
' with a formula containing the string (fif_str) are returned, else the function
' ends with TRUE with the first found.
' ------------------------------------------------------------------------------
    Const PROC = "FoundInFormulas"
    
    Dim cel As Range
    Dim cll As New Collection
    Dim wsh As Worksheet
    Dim rng As Range
    
    BoP PROC
    For Each wsh In fif_wbk.Worksheets
        If Not fif_wsh Is Nothing Then
            If Not wsh Is fif_wsh Then GoTo nw
        End If
        
        On Error Resume Next
        Set rng = wsh.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number <> 0 Then GoTo nw
        For Each cel In rng
            If InStr(1, cel.Formula, fif_str) > 0 Then
                FoundInFormulas = True
                If IsMissing(fif_cll) Then
                    GoTo xt
                Else
                    cll.Add cel
                End If
            End If
        Next cel
        
nw: Next wsh

xt: Set fif_cll = cll
    Set cll = Nothing
    EoP PROC
    Exit Function

End Function

Private Function GetExcelObjectFromHwnd(ByVal hWndMain As LongPtr) As Application
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

                        
Public Function HasName(ByVal hn_rng As Range) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the range (hn_rng) has a Name.
' ------------------------------------------------------------------------------
    Dim nme As Name
    Dim s   As String
    
    On Error Resume Next
    Set nme = Nothing ' only including in case you are planning to use this approach in a loop.
    Set nme = hn_rng.Name
    On Error GoTo xt
    If Not nme Is Nothing Then
        s = nme.Name
        HasName = True
    End If
    
xt: Exit Function

End Function

Public Function HasNames(ByVal hn_rng As Range) As Dictionary
' ------------------------------------------------------------------------
' Returns a Dictionary with all Names for the range (hn_rng).
' ------------------------------------------------------------------------
    Dim nme As Name
    Dim dct As New Dictionary
    Dim wbk As Workbook
    
    Set wbk = hn_rng.Parent.Parent
    For Each nme In wbk.Names
        If nme.RefersTo = hn_rng.Name Then
            dct.Add nme.Name, nme.RefersTo
        End If
    Next nme
    
    Set HasNames = dct
    Set dct = Nothing

End Function

Public Function HasUrl(ByVal r As Range) As Boolean
    HasUrl = r.Hyperlinks.Count > 0
End Function

Public Sub IndicateByBackColor()
    Dim rgbInput            As XlRgbColor
    Dim rgbLockedNamed      As XlRgbColor
    Dim rgbLockedUnNamed    As XlRgbColor
    
    rgbInput = RGB(242, 242, 242)
    rgbLockedNamed = RGB(255, 255, 204)
    
    rgbLockedUnNamed = rgbGray
    
End Sub

Private Function IsWbkFullName(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is a Workbook's FullName - not necessarily existing.
' ----------------------------------------------------------------------------
    With New FileSystemObject
        If VarType(v) = vbString Then
            IsWbkFullName = IsWbkName(.GetFileName(v)) And (InStr(v, "\") <> 0 Or InStr(v, "/") <> 0)
        End If
    End With
End Function

Private Function IsWbkName(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when (v) is a Workbook's Name (without path!) .
' ----------------------------------------------------------------------------
    If VarType(v) = vbString Then
        IsWbkName = InStr(v, "\") = 0 And InStr(v, "/") = 0 And v Like "*.xl*"
    End If
End Function

Private Function IsWbkObject(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is a valid Workbook object.
' ----------------------------------------------------------------------------
    If VarType(v) = vbObject Then
        IsWbkObject = TypeName(v) = "Workbook"
    End If
End Function

Private Function IsWbObject(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is a valid Workbook object.
' ----------------------------------------------------------------------------
    If VarType(v) = vbObject Then
        IsWbObject = TypeName(v) = "Workbook"
    End If
End Function

Public Sub Raised(ByVal r As Range)
    With r
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.399945066682943
            .PatternTintAndShade = 0
        End With
    End With
End Sub

Public Sub SetBackColor(ByVal sbc_rng As Range, _
                        ByVal sbc_rbg As XlRgbColor)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    With sbc_rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .Color = sbc_rbg
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    
End Sub

Private Function SheetExists(ByVal vWb As Variant, _
                             ByVal vWs As Variant, _
                    Optional ByRef wsResult As Worksheet) As Boolean
' ------------------------------------------------------------
' Returns TRUE when the Worksheet (vWs) - which may be a
' Worksheet object or a Worksheet's name - exists in the
' Workbook (vWb).
' ------------------------------------------------------------
    Const PROC  As String = "SheetExists"    ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim wsTest  As Worksheet
    Dim wb      As Workbook
    Dim ws      As Worksheet

    SheetExists = False
    
    If TypeName(vWb) = "Nothing" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter vWb) is ""Nothing""!"
    
    If Not WbkIsOpen(vWb, wb) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Workbook (parameter vWb) is not open!"

    If TypeName(vWs) = "Nothing" _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The Worksheet (parameter vWs) for the Worksheet's existence check is ""Nothing""!"
    
    If Not TypeOf vWs Is Worksheet And VarType(vWs) <> vbString _
    Then Err.Raise AppErr(4), ErrSrc(PROC), "The Worksheet (parameter vWs) for the Worksheet's existence check is neither a Worksheet object nor a Worksheet's name or modulename!"
    
    If TypeOf vWs Is Worksheet Then
        Set ws = vWs
        For Each wsTest In wb.Worksheets
            If wsTest Is ws Then
                SheetExists = True
                Set wsResult = wsTest
                GoTo xt
            End If
        Next wsTest
        GoTo xt
    ElseIf VarType(vWs) = vbString Then
        For Each wsTest In wb.Worksheets
            If wsTest.Name = vWs Then
                SheetExists = True
                Set wsResult = wsTest
                GoTo xt
            End If
        Next wsTest
    End If
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Sunken(ByVal r As Range)
    With r
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.399945066682943
            .PatternTintAndShade = 0
        End With
    End With
End Sub

Function URLExists(ByVal ue_url As String) As Boolean
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    Dim oRequest    As Object
    Dim var         As Variant

    On Error GoTo xt
    Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    With oRequest
      .Open "GET", ue_url, False
      .sEnd
      var = .StatusText
    End With
    URLExists = var = "OK"

xt: Set oRequest = Nothing
    Exit Function

End Function

Public Sub UrlLet(ByVal u_wsh As Worksheet, _
                  ByVal u_rng As Range, _
                  ByVal u_address As String, _
         Optional ByVal u_text_to_display As String = vbNullString)
' --------------------------------------------------------------
'
' --------------------------------------------------------------
    u_rng.Hyperlinks.Delete
    u_wsh.Hyperlinks.Add Anchor:=u_rng, Address:=u_address, TextToDisplay:=u_text_to_display
End Sub

Public Function WbkGetOpen(ByVal go_wb As Variant, _
                  Optional ByVal go_read_only As Boolean = False) As Workbook
' ----------------------------------------------------------------------------
' Returns an open Workbook object or raises an error. If go_wb is a full
' path-file name, the file exists but is not open it is opened.
' Note: A ReadOnly mode has to be set by the caller.
' ----------------------------------------------------------------------------
    Const PROC = "GetOpen"
    
    On Error GoTo eh
    Dim sWbName     As String
    Dim sWbFullName As String
    Dim wbOpen      As Workbook
    Dim FSo         As New FileSystemObject
        
    Set wbOpen = Nothing
       
    If IsWbObject(go_wb) Then
        sWbName = go_wb.Name
        sWbFullName = go_wb.FullName
    ElseIf IsWbkFullName(go_wb) Then
        sWbName = FSo.GetFileName(go_wb)
        sWbFullName = go_wb
    ElseIf IsWbkName(go_wb) Then
        sWbName = go_wb
        sWbFullName = vbNullString
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter go_wb) is neither a Workbook object nor a string (name or fullname)!" & vbLf & _
                                           "(TypeName of argument = '" & TypeName(go_wb) & "'!)"
    End If

    If IsWbkObject(go_wb) Then
        Set WbkGetOpen = go_wb
        GoTo xt
    End If
    
    With AllWbksOpen
        If .Exists(sWbName) Then
            Set WbkGetOpen = .Item(sWbName)
            If sWbFullName = vbNullString Then GoTo xt
            Set wbOpen = .Item(sWbName)
            If wbOpen.FullName = sWbFullName Then GoTo xt
            '~~ A Workbook with the Name is open but it has a different FullName
            If FSo.FileExists(sWbFullName) Then
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
            If Not FSo.FileExists(sWbFullName) _
            Then Err.Raise AppErr(4), ErrSrc(PROC), "A Workbook named '" & sWbFullName & "' does not exist!"
            Set WbkGetOpen = Workbooks.Open(sWbFullName, , go_read_only)
            GoTo xt
        End If
    End With
        
xt: Set FSo = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsOpen(ByVal wb As Variant, _
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
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim OpenWbks As Dictionary
    Dim OpenWbk  As Workbook
    Dim FSo      As New FileSystemObject
    Dim WbName As String
    
    If Not IsWbkObject(wb) And Not IsWbkFullName(wb) And Not IsWbkName(wb) And Not TypeName(wb) = "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter wb) is neither a Workbook object nor a Workbook's name or fullname)!"
       
    Set OpenWbks = mWbk.Opened
    If IsWbkName(wb) Then
        '~~ wb is a Workbook's Name including its extension
        WbName = FSo.GetFileName(wb)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ When a Workbook's Name is provided the Workbook is only regarde open when the open
            '~~ Workbook has the same name (i.e. including its extension)
            If FSo.GetFile(OpenWbk.FullName).Name <> FSo.GetFileName(wb) Then Set OpenWbk = Nothing
        End If
    ElseIf IsWbkFullName(wb) Then
        WbName = FSo.GetFileName(wb)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ The provided (wb) specifies an exist Workbook file. This Workbook is regarded open (and returned as opject)
            '~~ when a Workbook with its Name (including the extension!) is open regardless in which location
            If FSo.GetFile(OpenWbk.FullName).Name <> FSo.GetFileName(wb) Then Set OpenWbk = Nothing
        End If
    ElseIf IsWbkObject(wb) Then
        WbName = wb.Name
        If Opened.Exists(WbName) Then
            Set OpenWbk = OpenWbks.Item(WbName)
        End If
    Else
        '~~ If wb is a Workbook's WbName it is regarded open when one with that WbName is open
        '~~ regrdless its extension
        If OpenWbks.Exists(wb) Then Set OpenWbk = OpenWbks.Item(wb)
    End If
    
xt: If IsWbkObject(OpenWbk) Then
        WbkIsOpen = True
        Set wb_result = OpenWbk
    End If
    Set FSo = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

