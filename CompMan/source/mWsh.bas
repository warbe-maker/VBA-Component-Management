Attribute VB_Name = "mWsh"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mWsh: Common Worksheet services.
' =====================
'
' Public services:
' ----------------
' - ChangeCodeName  Renames the provided VB-Component identified by its current
'                   CodeName to the new provided name and returns it as
'                   VB-Component object.
' - Delete          Provides a 'clean' deletion of a Worksheet by removing all
'                   relevant Name objects beforehand in order th prevent invalid
'                   Name object reulting from the deletion.
' - Exists          Returns TRUE when in the Workbook a given Worksheet exists
'                   whereby the Worksheet argument may be its name or code-name.
' - HasUrl          Returns TRUE when a range has an url.
' - Url             Adds or modifies a Worksheet ranges url with formatting
'                   options.
' - Value Let/Get   Returns from/write to a Worksheet's range a variant value
'                   whereby the range argument may be a range name or a range
'                   object.
' - README          Displays the component's README in the corresponding
'                   public GitHub repository.
'
' W. Rauschenberger, Berlin Dec 2024
' See: https://github.com/warbe-maker/VBA-Excel-Worksheet
' ------------------------------------------------------------------------------
Private Const GITHUB_REPO_URL = "GITHUB_REPO_URL"

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

Public Enum enExistenceQuality
    enNameOrCodeName
    enNameAndCodeName
End Enum

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    Const README_URL = "/blob/master/README.md"
    
    If r_bookmark = vbNullString _
    Then ShellRun GITHUB_REPO_URL & README_URL _
    Else: ShellRun GITHUB_REPO_URL & README_URL & "#" & r_bookmark
        
End Sub

Public Function ShellRun(ByVal sr_string As String, _
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

Public Sub ChangeCodeName(ByVal ccn_wbk As Workbook, _
                          ByVal ccn_old As String, _
                          ByVal ccn_new As String, _
                 Optional ByRef ccn_vbc_renamed As VBComponent)
' ------------------------------------------------------------------------------
' Renames the VBComponent identified by the name (ccn_old) to (ccn_new) and
' returns it (ccn_vbc_renamed).
' ------------------------------------------------------------------------------
    Const PROC = "ChangeCodeName"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    
    For Each vbc In ccn_wbk.VBProject.VBComponents
        If vbc.Name = ccn_old Then
            vbc.Name = ccn_new
            Set ccn_vbc_renamed = vbc
            Exit For
        End If
    Next vbc
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Delete(ByVal d_wsh As Worksheet)
' ------------------------------------------------------------------------------
' Provides a 'clean deletion' of a Worksheet by removing all relevant Name
' objects beforehand in order to prevent invalid Name objects reulting from the
' deletion
' ------------------------------------------------------------------------------
    Const PROC = "Delete"
    
    On Error GoTo eh
    RemoveNames d_wsh
    d_wsh.Delete
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub RemoveNames(ByVal rn_wsh As Worksheet)
' ----------------------------------------------------------------------------
' Removes all Name objects which either refer to a range of the Worksheet or
' are scoped to the Worksheet.
' ----------------------------------------------------------------------------
    Const PROC = "RemoveNames"
    
    On Error GoTo eh
    Dim wbk As Workbook
    Dim nme As Name
    
    Set wbk = rn_wsh.Parent
    For Each nme In wbk.Names
        If InStr(nme.RefersTo, "=" & rn_wsh.Name & "!") <> 0 _
        Or InStr(nme.Name, rn_wsh.Name & "!") <> 0 Then
            nme.Delete
        End If
    Next nme

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Property Let Url(Optional ByVal u_wsh As Worksheet, _
                        Optional ByVal u_rng As Range, _
                        Optional ByVal u_underline As XlUnderlineStyle = xlUnderlineStyleSingle, _
                        Optional ByVal u_font_size As Long = 11, _
                                 ByVal u_url As String)
' ----------------------------------------------------------------------------
' Adds or modifies in a Worksheet (u_wsh) a ranges (u_rng) url (s_url) with
' fomt options (u_underline, u_font).
' ----------------------------------------------------------------------------
    Const PROC = "Url"
    
    On Error GoTo eh
    Dim sAddress    As String
    Dim sSubAddress As String
    Dim bProtected  As Boolean
    
    Application.ScreenUpdating = False
    bProtected = u_wsh.ProtectContents
    If bProtected Then u_wsh.Unprotect
    
    sAddress = Split(u_url, "#")(0)
    sSubAddress = Split(u_url, "#")(1)
    
    If Not HasUrl(u_rng) Then
        ActiveSheet.Hyperlinks.Add Anchor:=u_rng _
                                 , Address:=sAddress _
                                 , SubAddress:=sSubAddress
    Else
        With u_rng.Hyperlinks(1)
            .Address = sAddress
            .SubAddress = sSubAddress
        End With
    End If
    
    With u_rng.Font
        .Name = "Calibri"
        .Size = u_font_size
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = u_underline
        .ThemeColor = xlThemeColorHyperlink
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    If bProtected Then u_wsh.Protect
    
xt: Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

' ----------------------------------------------------------------------------
' Returns/writes a value from/to a Worksheet (v_wsh) identified by (v_name)
' which may be a ranges name or a range object.
' ----------------------------------------------------------------------------
Public Property Get Value(Optional ByVal v_wsh As Worksheet, _
                          Optional ByVal v_name As Variant) As Variant
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
                                   ByVal v_value As Variant)
    Const PROC = "Value-Let"
    
    Dim rng         As Range
    Dim bProtected  As Boolean
    
    On Error Resume Next
    Select Case TypeName(v_name)
        Case "String": Set rng = v_wsh.Range(v_name)
        Case "Range":  Set rng = v_name
    End Select
    If Err.Number <> 0 _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Worksheet '" & v_wsh.Name & "' has no range with a name '" & v_name & "'!"
    
    bProtected = v_wsh.ProtectContents
    
    If bProtected And rng.Locked Then
        '~~ Unprotect is required only when the range is locked and the sheet is protected
        On Error Resume Next
        v_wsh.Unprotect
        If Err.Number <> 0 _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "The Worksheet '" & v_wsh.Name & "' is apparently password protected which is not supported by this component's Value service!"
        rng.Value = v_value
        If bProtected Then v_wsh.Protect
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

Private Sub BoP(ByVal b_proc As String, _
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

Private Sub EoP(ByVal e_proc As String, _
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
    ErrSrc = "mWsh" & "." & sProc
End Function

Private Function IsOpen(ByVal i_wbk As Workbook) As Boolean
    On Error Resume Next
    IsOpen = i_wbk.Name <> vbNullString
End Function

Public Function Exists(ByVal x_wbk As Workbook, _
              Optional ByVal x_wsh As Worksheet = Nothing, _
              Optional ByVal x_wsh_name As String = vbNullString, _
              Optional ByVal x_wsh_code_name As String = vbNullString, _
              Optional ByVal x_quality As enExistenceQuality = enNameOrCodeName, _
              Optional ByRef x_wsh_result As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when in the Workbook (x_wbk) depending on what is provided:
' - a Worksheet named and/or code-named like the provided Worksheet (x_wsh)
'   exists whereby and/or depends on the requested quality (x_quality) which
'   defaults to Name or CodeName
' - a Worksheet named (x_wsh_name) exists
' - a Worksheet code-named (x_wsh_code-name) exists.
' When the Workbook is not open the function returns FALSE without notice.
' ------------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim wsh         As Worksheet
    Dim bByName     As Boolean
    Dim bByCodeName As Boolean
    
    If Not IsOpen(x_wbk) Then GoTo xt
    If Not x_wsh Is Nothing Then
        x_wsh_name = x_wsh.Name
        x_wsh_code_name = x_wsh.CodeName
    End If
    bByName = x_wsh_name <> vbNullString
    bByCodeName = x_wsh_code_name <> vbNullString
    
    For Each wsh In x_wbk.Worksheets
        Select Case True
            Case bByName And bByCodeName
                If x_quality = enNameAndCodeName Then
                    Exists = wsh.Name = x_wsh_name And wsh.CodeName = x_wsh_code_name
                Else
                    Exists = wsh.Name = x_wsh_name Or wsh.CodeName = x_wsh_code_name
                End If
                Set x_wsh_result = wsh
                If Exists Then Exit For
            Case bByCodeName
                Exists = wsh.CodeName = x_wsh_code_name
                Set x_wsh_result = wsh
                If Exists Then Exit For
            Case bByName
                Exists = wsh.Name = x_wsh_name
                Set x_wsh_result = wsh
                If Exists Then Exit For
        End Select
    Next wsh
        
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function HasUrl(ByVal h_rng As Range) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when a range (h_rng) has an url.
' ------------------------------------------------------------------------------
    HasUrl = h_rng.Hyperlinks.Count <> 0
End Function

