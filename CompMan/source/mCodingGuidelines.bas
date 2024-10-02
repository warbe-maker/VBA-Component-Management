Attribute VB_Name = "mCodingGuidelines"
Option Explicit
' ------------------------------------------------------------------------
' Standard-Module mCodingGuidelines:
' ===================================
'
' W. Rauschenberger, Berlin Jan 2024
' See also
' https://github.com/warbe-maker/VBA-Component-Management/blob/master/CODING_RULES.md
' ------------------------------------------------------------------------

Private Sub Arguments(ByVal a_any1 As String, _
                      ByVal a_any2 As Variant)
' ------------------------------------------------------------------------
' Arguments have a prefix followed by an underline. Because this notation
' is not used anywhere else it clearly indicates where an argument is used.
' Additionally this notation avoids any conflicts with system arguments of
' which the case may change from upper- to lower-case or vice versa.
' Because VBA is case insensitive, constants must not use a single
' character prefix
' ------------------------------------------------------------------------
End Sub

Private Sub Constants()
' ------------------------------------------------------------------------
' Constants are
' - written in upper case letters
' - use _ (underlines) last but not least for better readability
' - Begin with more than one character in order to clearly differ them
'   from arguments.
' ------------------------------------------------------------------------
    Const ERR_ANY_1 = "any1"
    Const ERR_ANY_2 = "any2"
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
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
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

Private Sub ErrorHandling()
' ------------------------------------------------------------------------
' This is a common Sub template. It uses:
' - ErrSrc to indicate the component
' - A common ErrMsg function which supports the optional use of the Common
'   Components mErH, mMsg/fMsg, and mTrc/clsTrc of which the availability
'   and usage is indicated by corresponding Conditional Compile Arguments.
' ------------------------------------------------------------------------
    Const PROC = "ErrorHandling"
    
    On Error GoTo eh
    
    ' any code including declarations
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub DebugPrint()
' ------------------------------------------------------------------------
' Any Debug.Print at first displays the source, i.e. the procedure in
' which it is used. Which means that any component in which a Debug.Print
' statement is used must have a ErrSrc Function which prefixes the
' procedure name with the component's name.
' ------------------------------------------------------------------------
    Const PROC = "DebugPrint"
    
    '~~ Example:
    Debug.Print ErrSrc(PROC) & ": " '& ......
    
End Sub

Private Sub PathToTheError()
' ------------------------------------------------------------------------
' In order to support the display of a "path to the error" procedures
' may use BoP/EoP (Beginn/End of Procedure) statements which are uesd
' to maintain a call stack which is used with the error message display
' to show the path to the error.
' Note: This procedure is a common template for all those which
'       potentially should be included in a "path-to-the-error" when an
'       error message is displayed.
' ------------------------------------------------------------------------
    Const PROC = "PathToTheError"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    ' any code including declarations
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCodingGuidelines." & sProc
End Function

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

Private Sub BoP(ByVal b_proc As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

