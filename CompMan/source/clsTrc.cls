VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' Class Module clsTrc: Common VBA Execution Trace Service. Traces the elapsed
' ==================== execution of procedures and code snippets with the
' highest possible precision. The trace log is written to a file which ensures
' at least a partial trace in case the execution terminates by exception. When
' this module is installed the availability of the sevice is triggered/
' activated by the Conditional Compile Argument 'clsTrc = 1'. When the
' Cond. Comp. Arg. is 0 all services are disabled even when the module is
' installed, avoiding any negative effect on the performance though this
' effect very little anyway.
' Please note that only the main component using this component will declare
' it: `Public Trc As clsTrc`.
'
' Public Methods:
' ----------------
' BoC      Indicates the (B)egin (o)f the execution trace of a (C)ode
'          snippet.
' BoP      Indicates the (B)egin (o)f the execution trace of a (P)rocedure.
' BoP_ErH  Exclusively used by the mErH module.
' Continue Commands the execution trace to continue taking the execution
'          time when it had been paused. Pause and Continue is used by the
'          mErH module for example to avoid useless execution time counting
'          while waiting for the users reply.
' Dsply    Displays the content of the trace-log-file (FileFullName).
' EoC      Indicates the (E)nd (o)f the execution trace of a (C)ode snippet.
' EoP      Indicates the (E)nd (o)f the execution trace of a (P)rocedure.
' LogInfo  Explicitly writes an entry to the trace lof file by considering the
'          current nesting level.
' Pause    Stops the execution traces time taking, e.g. while an error message
'          is displayed.
' README   Displays the README in the corresponding public GitHub rebo.
'
' Public Properties:
' ------------------
' FileFullName r/w Specifies the full name of the trace-log-file, defaults to
'                  Path & "\" & FileName when not specified.
' FileName         Specifies the trace-log-file's name, defaults to
'                  "ExecTrace.log" when not specified.
' KeepLogs     w   Specifies the number of days a trace-log-file is kept until
'                  it is deleted and re-created.
' Path         r/w Specifies the path to the trace-log-file, defaults to
'                  ThisWorkbook.Path when not specified.
' Title        w   Specifies a trace-log title
'
' Uses (for test purpose only!):
' ------------------------------
' mMsg/fMsg Supports a more comprehensive and well designed error message.
'           See https://github.com/warbe-maker/VBA-Message
'           for how to install and use it in any module.
'
' mErH      Privides an error message with additional information and options.
'           See https://github.com/warbe-maker/VBA-Error
'           for how to install an use in any module.
'
' Requires:
' ---------
' Reference to 'Microsoft Scripting Runtime'
'
' W. Rauschenberger, Berlin, Aug 2024
' See: https://github.com/warbe-maker/VBA-Trace
' ----------------------------------------------------------------------------
Private Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-Trace"
Private FSo                     As New FileSystemObject

#If Not mMsg = 1 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Cond. Comp. Arg. 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Cond. Comp. Arg. mMsg = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Enum enTraceInfo
    enItmId = 1
    enItmDir
    enItmLvl
    enItmTcks
    enItmArgs
End Enum

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

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cySysFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const DIR_BEGIN_ID      As String = ">"     ' Begin procedure or code trace indicator
Private Const DIR_END_ID        As String = "<"     ' End procedure or code trace indicator
Private Const TRC_LOG_SEC_FRMT  As String = "00.0000 "

Private bNewFile            As Boolean          ' default to False
Private cyTcksAtStart       As Currency         ' Trace log to file
Private LastNtry            As Collection       '
Private cySysFrequency      As Currency         ' Execution Trace SysFrequency (initialized with init)
Private cyTcksOvrhdItm      As Currency         ' Execution Trace time accumulated by caused by the time tracking itself
Private cyTcksOvrhdTrc      As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksOvrhdTrcStrt  As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksPaused        As Currency         ' Accumulated with procedure Continue
Private cyTcksPauseStart    As Currency         ' Set with procedure Pause
Private dtTraceBegin        As Date             ' Initialized at start of execution trace
Private iTrcLvl             As Long             ' Increased with each begin entry and decreased with each end entry
Private sFileFullName       As String
Private sFirstTraceItem     As String
Private sTitle              As String
Private TraceStack          As Collection       ' Trace stack for the trace log written to a file
Private bLogToFileSuspended As Boolean
Private lKeepLogs           As Long

Private Sub Class_Initialize()

    Set LastNtry = Nothing
    dtTraceBegin = Now()
    cyTcksOvrhdItm = 0
    iTrcLvl = 0
    cySysFrequency = 0
    sFirstTraceItem = vbNullString
    
End Sub

Private Sub Class_Terminate()
    
    Set FSo = Nothing
    If lKeepLogs <> 0 Then Reorg sFileFullName, lKeepLogs
    
End Sub

' ----------------------------------------------------------------------------
' Universal read/write array procedure.
' Read:  Returns Null when a given array (c_arr) is not allocated or a
'        provided index is beyond/outside current number of items.
' Write: - Adds an item (c_var) to an array (c_arr) when no index is provided
'          or adds it with the provided index
'        - When an index is provided, the item is inserted/updated at the
'          given index, even when the array yet doesn't exist or yet is not
'          allocated.
' ----------------------------------------------------------------------------
Private Property Get Arry(Optional ByRef c_arr As Variant, _
                          Optional ByVal c_index As Long = -1) As Variant
    Dim i As Long
    
    If IsArray(c_arr) Then
        On Error Resume Next
        i = LBound(c_arr)
        If Err.Number = 0 Then
            If c_index >= LBound(c_arr) And c_index <= UBound(c_arr) _
            Then Arry = c_arr(c_index)
        End If
    End If
    
End Property

Private Property Let Arry(Optional ByRef c_arr As Variant, _
                          Optional ByVal c_index As Long = -99, _
                                   ByVal c_var As Variant)
    Const PROC = "Arry-Let"
    
    Dim bIsAllocated As Boolean
    Dim s            As String
    
    If IsArray(c_arr) Then
        On Error GoTo -1
        On Error Resume Next
        bIsAllocated = UBound(c_arr) >= LBound(c_arr)
        On Error GoTo eh
    ElseIf VarType(c_arr) <> 0 Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Not a Variant type!"
    End If
    
    If bIsAllocated = True Then
        '~~ The array has at least one item
        If c_index = -99 Then
            '~~ When for an allocated array no index is provided, the item is added
            ReDim Preserve c_arr(UBound(c_arr) + 1)
            c_arr(UBound(c_arr)) = c_var
        ElseIf c_index >= 0 And c_index <= UBound(c_arr) Then
            '~~ Replace an existing item
            c_arr(c_index) = c_var
        ElseIf c_index > UBound(c_arr) Then
            '~~ New item beyond current UBound
            ReDim Preserve c_arr(c_index)
            c_arr(c_index) = c_var
        ElseIf c_index < LBound(c_arr) Then
            Err.Raise AppErr(2), ErrSrc(PROC), "Index is less than LBound of array!"
        End If
        
    ElseIf bIsAllocated = False Then
        '~~ The array does yet not exist
        If c_index = -99 Then
            '~~ When no index is provided the item is the first of a new array
            c_arr = Array(c_var)
        ElseIf c_index >= 0 Then
            ReDim c_arr(c_index)
            c_arr(c_index) = c_var
        Else
            Err.Raise AppErr(3), ErrSrc(PROC), "the provided index is less than 0!"
        End If
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get DefaultFileName() As String
    DefaultFileName = ThisWorkbook.Path & "\" & "ExecTrace.log"
End Property

Private Property Get DIR_BEGIN_CODE() As String:            DIR_BEGIN_CODE = DIR_BEGIN_ID:                  End Property

Private Property Get DIR_BEGIN_PROC() As String:            DIR_BEGIN_PROC = VBA.String$(2, DIR_BEGIN_ID):  End Property

Private Property Get DIR_END_CODE() As String:              DIR_END_CODE = DIR_END_ID:                      End Property

Private Property Get DIR_END_PROC() As String:              DIR_END_PROC = VBA.String$(2, DIR_END_ID):      End Property

Friend Property Get FileFullName() As String
    If sFileFullName = vbNullString Then
        sFileFullName = DefaultFileName
    End If
    FileFullName = sFileFullName
End Property

Friend Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Specifies the trace-log-file's name and location, by the way maintains the
' Path and the FileName property.
' ----------------------------------------------------------------------------
    Dim lDaysAge As Long
    
    With FSo
        sFileFullName = s
        If .FileExists(sFileFullName) Then
            '~~ In case the file already existed it may have passed the KeepLogs limit
            lDaysAge = VBA.DateDiff("d", .GetFile(sFileFullName).DateLastAccessed, Now())
            If lDaysAge > lKeepLogs Then
                .DeleteFile sFileFullName
            End If
        End If
    End With

End Property

Private Property Get ItmArgs(Optional ByRef t_entry As Collection) As Variant
    ItmArgs = t_entry("I")(enItmArgs)
End Property

Private Property Get ItmDir(Optional ByRef t_entry As Collection) As String
    ItmDir = t_entry("I")(enItmDir)
End Property

Private Property Get ItmId(Optional ByRef t_entry As Collection) As String
    ItmId = t_entry("I")(enItmId)
End Property

Private Property Get ItmLvl(Optional ByRef t_entry As Collection) As Long
    ItmLvl = t_entry("I")(enItmLvl)
End Property

Private Property Get ItmTcks(Optional ByRef t_entry As Collection) As Currency
    ItmTcks = t_entry("I")(enItmTcks)
End Property

Friend Property Let KeepLogs(ByVal l As Long):  lKeepLogs = l:  End Property

Private Property Let Log(ByVal l_line As String)
' ----------------------------------------------------------------------------
' Writes the string (l_line) to the FileFullName.
' ----------------------------------------------------------------------------
                
    Open sFileFullName For Append As #1
    Print #1, l_line
    Close #1

End Property

Friend Property Let LogInfo(ByVal l_info As String)
' ----------------------------------------------------------------------------
' Write an info line (l_info) to the trace-log-file (sFileFullName)
' ----------------------------------------------------------------------------
    Dim s As String
    s = LogNow _
      & String(Len(TRC_LOG_SEC_FRMT) * 2, " ") _
      & RepeatStrng("|  ", LogInfoLvl) _
      & "|  " _
      & l_info
    Log = s
    
End Property

Friend Property Get LogSuspended() As Boolean:              LogSuspended = bLogToFileSuspended:             End Property

Friend Property Let LogSuspended(ByVal b As Boolean):       bLogToFileSuspended = b:                        End Property

Private Property Let NtryItm(Optional ByVal t_entry As Collection, ByVal v As Variant)
    t_entry.Add v, "I"
End Property

Private Property Get NtryTcksOvrhdNtry(Optional ByRef t_entry As Collection) As Currency
    On Error Resume Next
    NtryTcksOvrhdNtry = t_entry("TON")
    If Err.Number <> 0 Then NtryTcksOvrhdNtry = 0
End Property

Private Property Let NtryTcksOvrhdNtry(Optional ByRef t_entry As Collection, ByRef cy As Currency)
    If t_entry Is Nothing Then Set t_entry = New Collection
    t_entry.Add cy, "TON"
End Property

Private Property Get SplitStr(ByRef s As String)
' ----------------------------------------------------------------------------
' Returns the split string in string (s) used by VBA.Split() to turn the
' string into an array.
' ----------------------------------------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
End Property

Private Property Get SysCrrntTcks() As Currency
    getTickCount SysCrrntTcks
End Property

Private Property Get SysFrequency() As Currency
    If cySysFrequency = 0 Then
        getFrequency cySysFrequency
    End If
    SysFrequency = cySysFrequency
End Property

Friend Property Get Title() As String:                      Title = sTitle:                                 End Property

Friend Property Let Title(ByVal s As String):               sTitle = s:                                     End Property

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

Private Function ArryIsAllocated(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Retunrs TRUE when the array (a_arr) is allocated, i.e. has at least one item.
' ----------------------------------------------------------------------------
    On Error Resume Next
    ArryIsAllocated = UBound(a_arr) >= LBound(a_arr)
    On Error GoTo -1
    
End Function

Private Function ArrayAsFile(ByVal a_array As Variant, _
                    Optional ByRef a_file As Variant = vbNullString, _
                    Optional ByVal a_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes all items of an array (a_arry) to a file (a_file) which might be a
' file object, a file's full name. When no file (a_file) is provided a
' temporary file is returned, else the provided file (a_file) as object.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
      
    If Not ArryIsAllocated(a_array) Then Exit Function
    
    Select Case True
        Case a_file = vbNullString:     a_file = TempFile
        Case TypeName(a_file) = "File": a_file = a_file.Path
    End Select
    
    If a_file_append _
    Then Open a_file For Append As #1 _
    Else Open a_file For Output As #1
    Print #1, Join(a_array, vbCrLf)
    Close #1
    Set ArrayAsFile = FSo.GetFile(a_file)
    
End Function

Private Function ArryErase(ByRef c_arr As Variant)
    If IsArray(c_arr) Then Erase c_arr
End Function

Public Sub BoC(ByVal b_id As String, _
      Optional ByVal b_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Begin of code sequence trace.
' ----------------------------------------------------------------------------
    Dim cll             As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    TrcBgn t_id:=b_id, t_dir:=DIR_BEGIN_CODE, t_args:=b_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub BoP(ByVal b_id As String, _
      Optional ByVal b_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Begin of procedure trace.
' ----------------------------------------------------------------------------
    Dim cll           As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If sFirstTraceItem = vbNullString _
    Then sFirstTraceItem = b_id
    TrcBgn t_id:=b_id, t_dir:=DIR_BEGIN_PROC, t_args:=b_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub BoP_ErH(ByVal b_id As String, _
          Optional ByVal b_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Begin of procedure trace, specifically for being used by the mErH module.
' ----------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If sFirstTraceItem = vbNullString _
    Then sFirstTraceItem = b_id
    TrcBgn t_id:=b_id, t_dir:=DIR_BEGIN_PROC, t_args:=b_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub Continue()
' ----------------------------------------------------------------------------
' Continues counting the execution time - which had been paused for a user
' intervention (wait for a reply button click).
' ----------------------------------------------------------------------------
    cyTcksPaused = cyTcksPaused + (SysCrrntTcks - cyTcksPauseStart)
End Sub

Public Sub Dsply()
' ----------------------------------------------------------------------------
' Display service using ShellRun to open the log-file by means of the
' application associated with the log-file's file-extenstion.
' ----------------------------------------------------------------------------
    If FileFullName <> vbNullString Then
        ShellRun FileFullName, WIN_NORMAL
    Else
        ShellRun FileFullName, WIN_NORMAL
    End If
End Sub

Private Function DsplyArgName(ByVal s As String) As Boolean

    If Right(s, 1) = ":" _
    Or Right(s, 1) = "=" _
    Or Right(s, 2) = ": " _
    Or Right(s, 2) = " :" _
    Or Right(s, 2) = "= " _
    Or Right(s, 2) = " =" _
    Or Right(s, 3) = " : " _
    Or Right(s, 3) = " = " _
    Then DsplyArgName = True
    
End Function

Private Function DsplyArgs(ByVal t_entry As Collection) As String
' ------------------------------------------------------------------------------
' Returns a string with the collection of the traced arguments. Any entry ending
' with a ":" or "=" is an arguments name with its value in the subsequent item.
' ------------------------------------------------------------------------------
    Dim va()    As Variant
    Dim i       As Long
    Dim sL      As String
    Dim sR      As String
    
    On Error Resume Next
    va = ItmArgs(t_entry)
    If Err.Number <> 0 Then Exit Function
    i = LBound(va)
    If Err.Number <> 0 Then Exit Function
    
    For i = i To UBound(va)
        If DsplyArgs = vbNullString Then
            ' This is the very first argument
            If DsplyArgName(va(i)) Then
                ' The element is the name of an argument followed by a subsequent value
                DsplyArgs = "|  " & va(i) & CStr(va(i + 1))
                i = i + 1
            Else
                sL = ">": sR = "<"
                DsplyArgs = "|  Argument values: " & sL & va(i) & sR
            End If
        Else
            If DsplyArgName(va(i)) Then
                ' The element is the name of an argument followed by a subsequent value
                DsplyArgs = DsplyArgs & ", " & va(i) & CStr(va(i + 1))
                i = i + 1
            Else
                sL = ">": sR = "<"
                DsplyArgs = DsplyArgs & "  " & sL & va(i) & sR
            End If
        End If
    Next i
End Function

Public Sub EoC(ByVal eoc_id As String, _
      Optional ByVal eoc_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' End of the trace of a code sequence.
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If StckIsEmpty(TraceStack) Then Exit Sub
    TrcEnd t_id:=eoc_id, t_dir:=DIR_END_CODE, t_args:=eoc_inf, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub EoP(ByVal e_id As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' End of the trace of a procedure.
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If StckIsEmpty(TraceStack) Then Exit Sub        ' Nothing to trace any longer. Stack has been emptied after an error to finish the trace
    
    TrcEnd t_id:=e_id, t_dir:=DIR_END_PROC, t_args:=e_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the end-of-trace entry

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
    ErrSrc = "mTrc." & sProc
End Function

Private Function FileAsArray(ByVal f_file As Variant, _
                    Optional ByVal f_empty_excluded = False, _
                    Optional ByVal f_trim As Variant = False) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsArray"
    
    On Error GoTo eh
    Dim arr     As Variant
    Dim sSplit  As String
    Dim s       As String
    
    If TypeName(f_file) = "String" Then f_file = FSo.GetFile(f_file)
    s = FileAsString(f_file, sSplit, f_empty_excluded)
    FileAsArray = StringAsArray(s, sSplit, f_trim)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FileAsString(ByVal f_file As Variant, _
                    Optional ByRef f_split As String, _
                    Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns a file's (f_file) - provided as full name or object - records/lines
' as a single string with the records/lines delimited (f_split).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsString"
    
    On Error GoTo eh
    Dim s       As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    '~~ An error is passed on to the caller
    If Not FSo.FileExists(f_file) Then Err.Raise AppErr(1), ErrSrc(PROC), _
                                       "The file, provided by a full name, does not exist!"
    
    Open f_file For Input As #1
    s = Input$(lOf(1), 1)
    Close #1
    
    f_split = SplitIndctr(s) ' may be vbCrLf or vbLf (when file is a download)
    
    '~~ Eliminate any trailing split string
    Do While Right(s, Len(f_split)) = f_split
        s = Left(s, Len(s) - Len(f_split))
        If Len(s) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        s = FileAsStringEmptyExcluded(s)
    End If
    FileAsString = s

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsStringEmptyExcluded(ByVal f_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (f_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Do While InStr(f_s, vbCrLf & vbCrLf) <> 0
        f_s = Replace(f_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = f_s
    
End Function

Private Function IsDelimiterLine(ByVal i_line As String) As Boolean
' -----------------------------------------------------------------------------------
' Returns TRUE when a line (i_line) is a delimiter line.
' -----------------------------------------------------------------------------------
    IsDelimiterLine = i_line Like "*==========*"
End Function

Private Function Itm(ByVal i_dir As String, _
                     ByVal i_id As String, _
                     ByVal i_lvl As Long, _
                     ByVal i_tcks As Currency, _
                     ByVal i_args As Variant) As Variant()
' ------------------------------------------------------------------------------
' Returns an array with the arguments ordered by their enumerated position.
' ------------------------------------------------------------------------------
    Dim av(1 To 6) As Variant
    
    av(enItmId) = i_id
    av(enItmDir) = i_dir
    av(enItmLvl) = i_lvl
    av(enItmTcks) = i_tcks
    av(enItmArgs) = i_args
    Itm = av
    
End Function

Private Sub LogBgn(ByVal l_ntry As Collection, _
          Optional ByVal l_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Writes a begin trace line to the trace-log-file (sFileFullName).
' ----------------------------------------------------------------------------
    Const PROC = "LogBgn"
    
    Dim ElapsedSecsTotal    As String
    Dim ElapsedSecs         As String
    Dim TopNtry             As Collection
    Dim s                   As String
    
    If sFileFullName = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Trace log not possible when no log file name has been provided!"
    
    Set TopNtry = StckTop(TraceStack)
    If TopNtry Is Nothing _
    Then ElapsedSecsTotal = vbNullString _
    Else ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(l_ntry))
    
    StckPush TraceStack, l_ntry
    
    If TraceStack.Count = 1 Then
        '~~ Provide separator line (when file exists) and header
        cyTcksAtStart = ItmTcks(l_ntry)
    
        '~~ Trace-Log header
        s = LogText(l_elpsd_total:="Elapsed " _
                    , l_elpsd_secs:="Length  " _
                    , l_strng:="Execution trace by 'Common VBA Execution Trace Service (mTrc)' (" & GITHUB_REPO_URL & ")")
        
        '~~ Separator line
        If FSo.FileExists(FileFullName) Then Log = String(Len(s), "=")
        Log = s
        
        '~~ Trace-Log title
        If sTitle = vbNullString Then
            Log = LogText(l_elpsd_total:="seconds " _
                        , l_elpsd_secs:="seconds " _
                        , l_strng:=ItmDir(l_ntry) & " Begin execution trace ")
        Else
            Log = LogText(l_elpsd_total:="seconds " _
                        , l_elpsd_secs:="seconds " _
                        , l_strng:=ItmDir(l_ntry) & " " & sTitle)
        End If
        '~~ Keep the ticks at start for the calculation of the elepased ticks with each entry
    End If
    
    ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(l_ntry))
    Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                , l_elpsd_secs:=ElapsedSecs _
                , l_strng:=RepeatStrng("|  ", ItmLvl(l_ntry)) & ItmDir(l_ntry) & " " & ItmId(l_ntry) _
                , l_args:=l_args)
    
End Sub

Private Function LogElapsedSecs(ByVal et_ticks_end As Currency, _
                                     ByVal et_ticks_start As Currency) As String
    LogElapsedSecs = Format(CDec(et_ticks_end - et_ticks_start) / CDec(SysFrequency), TRC_LOG_SEC_FRMT)
End Function

Private Function LogElapsedSecsTotal(ByVal et_ticks As Currency) As String
    LogElapsedSecsTotal = Format(CDec(et_ticks - cyTcksAtStart) / CDec(SysFrequency), TRC_LOG_SEC_FRMT)
End Function

Private Sub LogEnd(ByVal l_ntry As Collection)
' ------------------------------------------------------------------------------
' Write an end trace line to the trace-log-file (sFileFullName) - provided one
' had been specified - with the execution time calculated in seconds. When the
' TraceStack is empty write an additional End trace footer line.
' ------------------------------------------------------------------------------
    Const PROC = "LogEnd"
    
    On Error GoTo eh
    Dim BgnNtry             As Collection
    Dim ElapsedSecs         As String
    Dim ElapsedSecsTotal    As String
    Dim s                   As String
    
    StckPop TraceStack, l_ntry, BgnNtry
    ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(l_ntry))
    ElapsedSecs = LogElapsedSecs(et_ticks_end:=ItmTcks(l_ntry), et_ticks_start:=ItmTcks(BgnNtry))
    
    s = LogText(l_elpsd_total:=ElapsedSecsTotal _
              , l_elpsd_secs:=ElapsedSecs _
              , l_strng:=RepeatStrng("|  ", ItmLvl(l_ntry)) _
                      & ItmDir(l_ntry) _
                      & " " _
                      & ItmId(l_ntry) _
             , l_args:=ItmArgs(l_ntry))
    Log = s
    
    If TraceStack.Count = 1 Then
        
        '~~ Trace bottom title
        If sTitle = vbNullString Then
            Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                        , l_elpsd_secs:=ElapsedSecs _
                        , l_strng:=ItmDir(l_ntry) & " " & "End execution trace ")
        Else
            Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                        , l_elpsd_secs:=ElapsedSecs _
                        , l_strng:=ItmDir(l_ntry) & " " & sTitle)
        End If
        
        '~~ Trace footer and summary
        Log = LogText(l_strng:="Execution trace by 'Common VBA Execution Trace Service (mTrc)' (" & GITHUB_REPO_URL & ")")
        Log = LogText(l_strng:="Impact on the overall performance (caused by the trace itself): " & Format(LogSecsOverhead * 1000, "#0.0") & " milliseconds!")
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogEntry(ByVal t_id As String, _
                     ByVal t_tcks As Currency, _
                     ByVal t_dir As String, _
                     ByVal t_lvl As Long, _
            Optional ByVal t_args As String = vbNullString, _
            Optional ByRef t_ntry As Collection)
' ----------------------------------------------------------------------------
' Writes an entry to the trace-log-file by meand of LogBgn or LogEnd.
' ----------------------------------------------------------------------------
    Const PROC = "LogEntry"
    
    Static sLastDrctv   As String
    Static sLastId      As String
    Static lLastLvl     As String
    
    Dim bAlreadyAdded   As Boolean
    
    If Not LastNtry Is Nothing Then
        '~~ When this is not the first entry added the overhead ticks caused by the previous entry is saved.
        '~~ Saving it with the next entry avoids a wrong overhead when saved with the entry itself because.
        '~~ Its maybe nitpicking but worth the try to get execution time figures as correct/exact as possible.
        If sLastId = t_id And lLastLvl = t_lvl And sLastDrctv = t_dir Then bAlreadyAdded = True
        If Not bAlreadyAdded Then
            NtryTcksOvrhdNtry(LastNtry) = cyTcksOvrhdTrc
        Else
            Debug.Print ErrSrc(PROC) & ": " & ItmId(LastNtry) & " already added"
        End If
    End If
    
    If Not bAlreadyAdded Then
        Set t_ntry = Ntry(n_id:=t_id, n_dir:=t_dir, n_lvl:=t_lvl, n_tcks:=t_tcks, n_args:=t_args)
        
        If t_dir Like DIR_BEGIN_CODE & "*" _
        Then LogBgn t_ntry, t_args _
        Else LogEnd t_ntry
        Set LastNtry = t_ntry
        sLastDrctv = t_dir
        sLastId = t_id
        lLastLvl = t_lvl
    Else
        Debug.Print ErrSrc(PROC) & ": " & ItmId(LastNtry) & " already added"
    End If

End Sub

Private Function LogInfoLvl() As Long
    
    If ItmDir(LastNtry) Like DIR_END_ID & "*" _
    Then LogInfoLvl = ItmLvl(LastNtry) - 1 _
    Else LogInfoLvl = ItmLvl(LastNtry)

End Function

Private Function LogNow() As String
    LogNow = Format(Now(), "YY-MM-DD hh:mm:ss ")
End Function

Private Function LogSecsOverhead()
    LogSecsOverhead = Format(CDec(cyTcksOvrhdTrc / CDec(SysFrequency)), TRC_LOG_SEC_FRMT)
End Function

Private Function LogText(Optional ByVal l_elpsd_total As String = vbNullString, _
                         Optional ByVal l_elpsd_secs As String = vbNullString, _
                         Optional ByVal l_strng As String = vbNullString, _
                         Optional ByVal l_args As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the uniformed assemled log text.
' ----------------------------------------------------------------------------
    LogText = LogNow
    
    If l_elpsd_total = vbNullString Then l_elpsd_total = String((Len(TRC_LOG_SEC_FRMT)), " ")
    LogText = LogText & l_elpsd_total
    
    If l_elpsd_secs = vbNullString Then l_elpsd_secs = String((Len(TRC_LOG_SEC_FRMT)), " ")
    LogText = LogText & l_elpsd_secs
    
    LogText = LogText & l_strng
    
    If l_args <> vbNullString Then
        If InStr(l_args, "!!") <> 0 _
        Then LogText = LogText & l_args _
        Else LogText = LogText & " (" & l_args & ")"
    End If
    
End Function

Private Function Max(ParamArray va() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ------------------------------------------------------------------------------
    Dim v   As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Public Sub NewFile(Optional ByVal n_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Specifies a new file's full name and deletes an existing one. When none
' is specified and no FileFullName had been specified either, the file's
' full name becomes a temp file
' ----------------------------------------------------------------------------
    If n_file = vbNullString Then n_file = FileFullName
    With FSo
        If .FileExists(n_file) Then .DeleteFile n_file, True
    End With
    sFileFullName = n_file
    
End Sub

Private Function Ntry(ByVal n_id As String, _
                      ByVal n_dir As String, _
                      ByVal n_lvl As Long, _
                      ByVal n_tcks As Currency, _
                      ByVal n_args As String) As Collection
' ----------------------------------------------------------------------------
' Return the arguments as elements in an array as an item in a collection.
' ----------------------------------------------------------------------------
    Const PROC = "Ntry"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim VarItm  As Variant
    
    VarItm = Itm(i_id:=n_id, i_dir:=n_dir, i_lvl:=n_lvl, i_tcks:=n_tcks, i_args:=n_args)
    NtryItm(cll) = VarItm
    Set Ntry = cll
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function NtryIsCode(ByVal cll As Collection) As Boolean
    
    Select Case ItmDir(cll)
        Case DIR_BEGIN_CODE, DIR_END_CODE: NtryIsCode = True
    End Select

End Function

Public Sub Pause()
    cyTcksPauseStart = SysCrrntTcks
End Sub

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const GITHUB_REPO_URL As String = "https://github.com/warbe-maker/VBA-Trace"
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        ShellRun GITHUB_REPO_URL & Replace("#" & r_bookmark, "##", "#")
    End If

End Sub

Private Sub Reorg(ByVal r_file_full_name As String, _
         Optional ByVal r_keep As Long = 1)
' ----------------------------------------------------------------------------
' Returns a log file (r_file_full_name) with any logs exceeding a number
' (r_keep) unstripped.
' Note: When only one log is kept the log file still may have two logs. The
'       one kept and the one added after the Reorg. When the Reorg is done
'       with the termination 1 will be 1.
' ----------------------------------------------------------------------------
    Const PROC = "Reorg"
    
    On Error GoTo eh
    Dim aFile   As Variant
    Dim aLog    As Variant
    Dim bAppend As Boolean
    Dim cllLog  As New Collection
    Dim cllLogs As New Collection
    Dim i       As Long
    Dim lLen    As Long
    Dim s       As String
    Dim v       As Variant
    Dim bTop    As Boolean: bTop = True
    
    If Not FSo.FileExists(r_file_full_name) Then GoTo xt
    If r_keep = 0 Then GoTo xt
    
    '~~ Collect the logs (i.e. lines delimited by a log delimiter line)
    '~~ into a Collection with each series of log entries/lines as a Collection
    '~~ with the array and the max line length as items
    aFile = FileAsArray(r_file_full_name)
    ArryErase aLog
    For i = LBound(aFile) To UBound(aFile)
        If IsDelimiterLine(aFile(i)) Then
            '~~ A delimiter line indicates the start of a new series of log entries
            cllLog.Add aLog
            cllLog.Add lLen
            cllLogs.Add cllLog
            ArryErase aLog
            Set cllLog = Nothing
            Set cllLog = New Collection
            lLen = 0
        End If
        lLen = Max(lLen, Len(aFile(i)))
        Arry(aLog) = aFile(i)
    Next i
    '~~ Collect the last series of log entries
    cllLog.Add aLog
    cllLog.Add lLen
    cllLogs.Add cllLog
    ArryErase aLog
    Set cllLog = Nothing
    Set cllLog = New Collection
    
    '~~ Remove all excessive logs
    If cllLogs.Count < r_keep Then GoTo xt
    Do While cllLogs.Count > r_keep
        cllLogs.Remove 1
    Loop
    
    '~~ Rewrite the remaining logs to the log file
    ArryErase aFile
    For Each v In cllLogs
        Set cllLog = v
        aLog = cllLog(1)
        lLen = cllLog(2)
        i = LBound(aLog)
        If IsDelimiterLine(aLog(i)) And bTop Then i = i + 1
        bTop = False
        For i = i To UBound(aLog)
            Arry(aFile) = aLog(i)
        Next i
        '~~ Re-write the log
    Next v
    ArrayAsFile aFile, sFileFullName
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function RepeatStrng(ByVal rs_s As String, _
                             ByVal rs_n As Long) As String
' ----------------------------------------------------------------------------
' Returns the string (s) concatenated (n) times. VBA.String in not appropriate
' because it does not support leading and trailing spaces.
' ----------------------------------------------------------------------------
    Dim i   As Long
    For i = 1 To rs_n: RepeatStrng = RepeatStrng & rs_s:  Next i
End Function

Private Function ShellRun(ByVal oue_string As String, _
                 Optional ByVal oue_show_how As Long = WIN_NORMAL) As String
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
    lRet = apiShellExecute(hWndAccessApp, vbNullString, oue_string, vbNullString, vbNullString, oue_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & oue_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Function SplitIndctr(ByVal s_strng As String, _
                    Optional ByRef s_indctr As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the split indicator of a string (s_strng) as string and as argument
' (s_indctr) provided no split indicator (s_indctr) is already provided.
' The dedection of a split indicator is bypassed in case one has already been
' provided.
' ----------------------------------------------------------------------------
    If s_indctr = vbNullString Then
        Select Case True
            Case InStr(s_strng, vbCrLf) <> 0: s_indctr = vbCrLf
            Case InStr(s_strng, vbLf) <> 0:   s_indctr = vbLf      ' e.g. in case of a downloaded file's_strng complete string
            Case InStr(s_strng, "|&|") <> 0:  s_indctr = "|&|"
            Case InStr(s_strng, ", ") <> 0:   s_indctr = ", "
            Case InStr(s_strng, "; ") <> 0:   s_indctr = "; "
            Case InStr(s_strng, ",") <> 0:    s_indctr = ","
            Case InStr(s_strng, ";") <> 0:    s_indctr = ";"
        End Select
    End If
    SplitIndctr = s_indctr

End Function

Private Sub StckAdjust(ByVal t_id As String)
    Dim cllNtry As Collection
    Dim i       As Long
    
    For i = TraceStack.Count To 1 Step -1
        Set cllNtry = TraceStack(i)
        If ItmId(cllNtry) = t_id Then
            Exit For
        Else
            TraceStack.Remove (TraceStack.Count)
            iTrcLvl = iTrcLvl - 1
        End If
    Next i

End Sub

Private Function StckEd(ByVal stck_id As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when last item pushed to the stack is identical with the item
' (stck_id) and level (stck_lvl).
' ----------------------------------------------------------------------------
    Const PROC = "StckEd"
    
    On Error GoTo eh
    Dim cllNtry As Collection
    Dim i       As Long
    
    For i = TraceStack.Count To 1 Step -1
        Set cllNtry = TraceStack(i)
        If ItmId(cllNtry) = stck_id Then ' And ItmLvl(cllNtry) = stck_lvl Then
            StckEd = True
            Exit Function
        End If
    Next i

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function StckIsEmpty(ByVal stck As Collection) As Boolean
    StckIsEmpty = stck Is Nothing
    If Not StckIsEmpty Then StckIsEmpty = stck.Count = 0
End Function

Private Sub StckPop(ByRef stck As Collection, _
                    ByVal stck_item As Variant, _
           Optional ByRef stck_ppd As Collection)
' ----------------------------------------------------------------------------
' Pops the item (stck_item) from the stack (stck) when it is the top item.
' When the top item is not identical with the provided item (stck_item) the
' pop is skipped.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    Dim cllTop  As Collection: Set cllTop = StckTop(stck)
    Dim cll     As Collection: Set cll = stck_item
    
    While ItmId(cll) <> ItmId(cllTop) And Not StckIsEmpty(TraceStack)
        '~~ Finish any unfinished code trace still on the stack which needs to be finished first
        If NtryIsCode(cllTop) Then
            EoC eoc_id:=ItmId(cllTop), eoc_inf:="ended by stack!!"
        Else
            EoP e_id:=ItmId(cllTop), e_args:="ended by stack!!"
        End If
        If Not StckIsEmpty(TraceStack) Then Set cllTop = StckTop(stck)
    Wend
    
    If StckIsEmpty(TraceStack) Then GoTo xt
    
    If ItmId(cll) = ItmId(cllTop) Then
        Set stck_ppd = cllTop
        TraceStack.Remove TraceStack.Count
        Set cllTop = StckTop(TraceStack)
    Else
        '~~ There is nothing to pop because the top item is not the one requested to pop
        Debug.Print ErrSrc(PROC) & ": " & "Stack Pop ='" & ItmId(cll) _
                  & "', Stack Top = '" & ItmId(cllTop) _
                  & "', Stack Dir = '" & ItmDir(cllTop) _
                  & "', Stack Lvl = '" & ItmLvl(cllTop) _
                  & "', Stack Cnt = '" & TraceStack.Count
        Stop
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub StckPush(ByRef stck As Collection, _
                     ByVal stck_item As Variant)
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_item
End Sub

Private Function StckTop(ByVal stck As Collection) As Collection
    If Not StckIsEmpty(stck) _
    Then Set StckTop = stck(stck.Count)
End Function

Private Function StringAsArray(ByVal s_strng As String, _
                      Optional ByVal s_split As String = vbNullString, _
                      Optional ByVal s_trim As Variant = True) As Variant
' ----------------------------------------------------------------------------
' Returns a string (s_strng) split into an array of strings. When no split
' indicator (s_split) is provided it one is found by examination of the
' string (s_strng). When the option (s_trim) is TRUE (the default), "R", or
' "L" the items in the array are returned trimmed accordingly.
' Example 1: arr = StringAsArray("this is a string", " ") is returned as an
'            array with 3 items: "this", "is", "a", "string".
' Example 2: arr = StringAsArray(FileAsString(FileName),sSplit,False) is
'            returned as any array with records/lines of the provided file,
'            whereby the lines are not trimmed, i.e. leading spaces are
'            preserved.
'            Note: The not provided split indicator has the advantage that it
'                  is provided by the SplitIndctr service, which in that case
'                  returns either vbCrLf or vbLf, the latter when the file is
'                  a download.
' Example 3: arr = FileAsArray(<file>) return the same as example 2.
' Note: Split indicators dedected by examination are: vbCrLf, vbLf, "|&|",
'       ", ", "; ", "," or ";". When neither is dedected vbCrLf is returned.
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim i   As Long
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_strng)
    arr = Split(s_strng, SplitIndctr(s_strng, s_split))
    If Not s_trim = False Then
        For i = LBound(arr) To UBound(arr)
            Select Case s_trim
                Case True:  arr(i) = VBA.Trim$(arr(i))
                Case "R":   arr(i) = VBA.RTrim$(arr(i))
                Case "L":   arr(i) = VBA.Trim$(arr(i))
            End Select
        Next i
    End If
    StringAsArray = arr

End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt", _
                          Optional ByVal f_create_as_textstream As Boolean = True) As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' The returned temporary file is registered in cllTestFiles for being removed
' either explicitely with TestFilesRemove or implicitly when the class
' terminates.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    If f_create_as_textstream Then FSo.CreateTextFile sTemp

End Function

Public Sub Terminate()
' ----------------------------------------------------------------------------
' Should be called by any error handling when a new execution trace is about
' to begin with the very first procedure's execution.
' ----------------------------------------------------------------------------
    Set TraceStack = Nothing
    cyTcksPaused = 0
End Sub

Private Sub TrcBgn(ByVal t_id As String, _
                   ByVal t_dir As String, _
          Optional ByVal t_args As String = vbNullString, _
          Optional ByRef t_cll As Collection)
' ----------------------------------------------------------------------------
' Collect a trace begin entry with the current ticks count for the procedure
' or code (item).
' ----------------------------------------------------------------------------
    Dim cy  As Currency
    
    cy = SysCrrntTcks - cyTcksPaused
           
    iTrcLvl = iTrcLvl + 1
    LogEntry t_id:=t_id _
           , t_dir:=t_dir _
           , t_tcks:=cy _
           , t_lvl:=iTrcLvl _
           , t_args:=t_args _
           , t_ntry:=t_cll
    StckPush TraceStack, t_cll

End Sub

Private Sub TrcEnd(ByVal t_id As String, _
                   ByVal t_dir As String, _
          Optional ByVal t_args As String = vbNullString, _
          Optional ByRef t_cll As Collection)
' ----------------------------------------------------------------------------
' Collect an end trace entry with the current ticks count for the procedure or
' code (item).
' ----------------------------------------------------------------------------
    Const PROC = "TrcEnd"
    
    On Error GoTo eh
    Dim cy  As Currency:    cy = SysCrrntTcks - cyTcksPaused
    Dim Top As Collection:  Set Top = StckTop(TraceStack)
    Dim Itm As Collection
        
    '~~ Any end trace for an item not on the stack is ignored. On the other hand,
    '~~ if on the stack but not the last item the stack is adjusted because this
    '~~ indicates a begin without a corresponding end trace statement.
    If Not StckEd(stck_id:=t_id) Then
        Exit Sub
    Else
        StckAdjust t_id
    End If
    
    If ItmId(Top) <> t_id And ItmLvl(Top) = iTrcLvl Then
        StckPop TraceStack, Top
    End If

    LogEntry t_id:=t_id _
           , t_tcks:=cy _
           , t_dir:=t_dir _
           , t_lvl:=iTrcLvl _
           , t_args:=t_args _
           , t_ntry:=t_cll
         
    StckPop stck:=TraceStack, stck_item:=t_cll, stck_ppd:=Itm
    iTrcLvl = iTrcLvl - 1

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

