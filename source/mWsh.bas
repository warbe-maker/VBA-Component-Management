Attribute VB_Name = "mWsh"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mWsh: Common Worksheet services
'
' Public services:
' - ChangeCodeName
' - Delete          Provides a 'clean deletion of a Worksheet by removing all
'                   relevant Name objects beforehand in order th prevent invalid
'                   Name object reulting from the deletion.
' - Exists
' - HasUrl
' - Url
' - Value Let/Get   ..
'
' W. Rauschenberger, Berlin Dec 2022
' ------------------------------------------------------------------------------
Public Enum enExistenceQuality
    enNameOrCodeName
    enNameAndCodeName
End Enum

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

Public Property Let Url(Optional ByVal su_wsh As Worksheet, _
                        Optional ByVal su_rng As Range, _
                        Optional ByVal url_underline As XlUnderlineStyle = xlUnderlineStyleSingle, _
                        Optional ByVal url_font_size As Long = 11, _
                                 ByVal su_url As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Url"
    
    On Error GoTo eh
    Dim sAddress    As String
    Dim sSubAddress As String
    Dim bProtected  As Boolean
    Dim rng         As Range
    
    Application.ScreenUpdating = False
    bProtected = su_wsh.ProtectContents
    If bProtected Then su_wsh.Unprotect
    
    sAddress = Split(su_url, "#")(0)
    sSubAddress = Split(su_url, "#")(1)
    
    If Not HasUrl(su_rng) Then
        ActiveSheet.Hyperlinks.Add Anchor:=su_rng _
                                 , Address:=sAddress _
                                 , SubAddress:=sSubAddress
    Else
        With su_rng.Hyperlinks(1)
            .Address = sAddress
            .SubAddress = sSubAddress
        End With
    End If
    
    With su_rng.Font
        .Name = "Calibri"
        .Size = url_font_size
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = url_underline
        .ThemeColor = xlThemeColorHyperlink
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    If bProtected Then su_wsh.Protect
    
xt: Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Value(Optional ByVal v_wsh As Worksheet, _
                          Optional ByVal v_name As Variant) As Variant
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
                                   ByVal v_value As Variant)
' ----------------------------------------------------------------------------
' Saves a variant value (v_value) to a Worksheet (v_wsh) whereby the target
' range (v_name) may be a range name or a Range object.
' ----------------------------------------------------------------------------
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
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

' -----------------------------------------------------------------------------------
' Standard  Module mWsh Checks the existence of a Worksheet.
'
' Methods:
' - Exists  Returns TRUE when the object exists
'
' Uses:     Standard Module mErrHndlr
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' W. Rauschenberger, Berlin August 2019
' -----------------------------------------------------------------------------------
Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' Basic service:
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Displays the error message by means of VBA.MsgBox when neither of the
'   following is installed
'
' Extendend service when other Common Components are installed and indicated via
' Conditional Compile Arguments:
' - Invokes mErH.ErrMsg when the Conditional Compile Argument ErHComp = 1
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = 1 (and
'   the mErH module is not installed / MsgComp not set)
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to turn
'          them into negative and in the error message back into a positive
'          number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See: https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
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
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
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

Public Function HasUrl(ByVal hu_rng As Range) As Boolean
    HasUrl = hu_rng.Hyperlinks.Count <> 0
End Function

