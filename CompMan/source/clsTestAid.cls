VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestAid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
' ----------------------------------------------------------------------------
' Class Module clsTestAid: Common services which support testing including
' ======================== regression testing, first and foremost providing
' a log file for performed tests and their result (Passed or Failed).
' - The class module by default initializes a "Test" folder in the
'   application's parent folder.
' - Each individual test is identified by a unique test-number
' - The result (of the tested procedure) and the expected result may be
'   provided in any form (Array, Collection, Dictionary, File, String, or
'   Boolean). When both had been provided for a test they are compared which
'   results in a failed or passed status.
' - Result/ResultExpected comparisons are numbered which means that an
'   individual test may have several result verifications
' - A regression-mode allows to specify "asserted errors" beforehand in order
'   to not interrupt a test sequence. Furthermore, all results are collected
'   and may finally be displayed (ResultSummaryLog). In contrast, i.e. in
'   non-regression-mode - which is the default) any failed test is immediately
'   displayed (ResultFailedLog)
'
' Public services:
' ----------------
' AssertedErrors              Error numbers regarded asserted for a specific
'                             test. Effective only when the Standard Module
'                             mErH is installed and indicated by the Cond.
'                             Comp. Argument `mErH = 1`.
' ExcludeFromComparison P   w Specifies a like string which is excluded from
'                             the evaluation/comparision of a result with an
'                             expected result.
' FolderUnZip                 Unzips a zipped folder to a target folder.
' FolderZip                   Zips a folder into a zip file.
' Guide                       Non modal message popup providing testing
'                             guidance.
' Instructions          M     Displays instructions for a required interaction
'                             by the "tester". When the _Common VBA Message
'                             Component_ `mMsg`/`fMsg` is installed (indicated
'                             by the Cond. Comp. Arg. `mMsg = 1`) the
'                             mMsg.Dsply service is used to display the
'                             instructions non-modal, else the instructions
'                             are provided in the Application.StatusBar.
' InstructionsDone      M     Unload the displayed instructions if any.
'                             Note: The form window is identified via the
'                             displayed title which is the key to the
'                             displayed instance of the `fMsg` UserForm.
' InstructionsPosLeft   P r/w Popup message position
' InstructionsPosTop    P r/w Popup message position
' ModeRegression        P r/w Boolean expression, indicates regression mode.
'                             When True, a failed test result is not displayed
'                             but displayed as summary (ResultSummaryLog).
' ResultSummaryLog      M     Displays a RegressionTest.log file provided
'                             ModeRegression = True.
' Result                P r/w Variant expression registers/returns the result
'                             of a test, may be a file object, any other
'                             object or any other data type.
' Result                P r   Returns Passed or Failed
' ResultExpected        P r/w The expected result of a test, a file object,
'                             any other object or any other data type
' CleanUp               M     Removes all files provided either by the
'                             TempTestItem or the TempFileFullName service.
' TestedComp            P   w Name of the tested VBComponent in a test
'                             procedure.
' TestedProc            P   w The name of the tested procedure (a Sub,
'                             Function, Property, or Method).
' TestedProcType            P   w Tested procedure type (Sub, Function, ...)
' TestFolder            P r/w String expression, specifies the test folder
' TestFile              M     Returns the file object with the
'                             TestFileFullName, raises an erro when not
'                             existing.
' TestFileFullName            Returns a test file's full name based on the
'                             current `TestId`
' TestId            P   w     Essential, obligatory, distinct identifcation of
'                             a test, provided as the very first Property for
'                             each ResultExpected/Result pair
'                             respectively.
' TempFileFullName      M
' TempTestItem          P r/w Saves the name of a temporary test file in a
'                             Collection for being removed when the class
'                             terminates or explicitly by the CleanUp service.
' TimerStart                  Initiates a maximum precision timer
' TimerElapsedTicks     M     Returns the elapsed ticks since start
' TimerEnd                    Returns the elapsed time since start in
'                             milliseconds whereby the timer overhead is
'                             deducted in order to return a result of 0
'                             milliseconds when nothing is measured between
'                             TimerStart and TimerEnd.
' TimerOverhead         P r   Returns the milliseconds the TimerStart/TimerEnd
'                             itself consumes when no taking any actual
'                             measurement.
' Title                 P   w Title of a test procedure, valid for all tests
'                             and verifications in it.
' Verification          P   w A string describing what specifically is verified
'                             by the tests ResultExpected/Result.
' BoTP/EoTP             M     Effective only when the execution trace module
'                             (mTrc or clsTrc) is installed, indicates the
'                             begin/end of a procedure.
'
' Requires:  Reference to "Microsoft Scripting Runtime"
'
' Uses:
' clsLog              To write and display failed test details
' mTrc/clsTrc         Only when an execution trace is desired, indicated by
'                     Cond. Comp. Arg. `mTrc = 1` or `clsTrc = 1`
'
' Note 1: In order to work completely autonomous the module integrates code
'         copies of services from the Common Components "mAlign" and clsLog".
' Note 2: This Common Component is "hosted" in a dedicated Workbook
'         "TestAid.xlsb" which provides regression testing. However, it may
'         be modified, ammendet, etc. in any VB-Project use of it provided
'         the Workbook is serviced by CompMan which cares for making the
'         modified code public again.
'
' W. Rauschenberger, Berlin Dec 2024
' ----------------------------------------------------------------------------
Public FSo                  As New FileSystemObject

Private Const RESULT_PASSED As String = " Passed "
Private Const RESULT_FAILED As String = " F a i l e d ! "

#If Win64 Or VBA7 Then
    Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As LongPtr
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Const WIN_NORMAL = 1         'Open Normal
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

Private aLogColAligns               As Variant
Private aLogColFills                As Variant
Private aLogColsArrangedHeaders     As Variant
Private aLogColWidths               As Variant
Private bModeRegression             As Boolean
Private cExecTimeMsecs              As Currency
Private cFrequency                  As Currency
Private cllTempTestItems            As Collection
Private cTicksCurrent               As Currency
Private cTicksStart                 As Currency
Private cTimerOverhead              As Currency
Private dctResult                   As New Dictionary
Private dctResultExpected           As New Dictionary
Private dctTestSummary              As New Dictionary
Private lCols                       As Long
Private lVerificationNumber         As Long
Private sDelim                      As String
Private sLogColsDelim               As String
Private sLogFile                    As String
Private sTestResultSummaryLogFile   As String
Private sNameTestFile               As String
Private sNameTestResultExpectedFile As String
Private sNameTestResultFile         As String
Private sResultEvaluated            As String
Private sExcludeFromComparison      As String
Private sTestedComp                 As String
Private sTestedProc                 As String
Private sTestedProcType             As String
Private sTestFileExtension          As String
Private sTestFolder                 As String
Private sTitle                      As String
Private sTitleRegression            As String
Private sTestId                     As String
Private sVerification               As String
Private v                           As Variant
Private vTempTestItem               As Variant
Private siInstructionsPosLeft       As Single
Private siInstructionsPosTop        As Single
Private siInstructionsTitle         As String

#If mMsg = 0 Then
    ' -------------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed which must
    ' be indicated by the Conditional Compile Argument mMsg = 1.
    ' See https://github.com/warbe-maker/Common-VBA-Message-Service
    ' -------------------------------------------------------------------------------
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#Else
    Private frmInstructions  As fMsg
#End If

Private Sub Class_Initialize()
    
    Const lLoop As Long = 1000
    
    Dim i               As Long
    Dim cTicksAtStart   As Currency
    
    Set cllTempTestItems = New Collection
    TestFolder = ThisWorkbook.Path & "\Test"                             ' default test folder
    TestResultSummaryLogFile = sTestFolder & "\Test-Result-Summary.log"  ' default regression test summary log file's full name
    lVerificationNumber = 0
    
    ' ------------------------------------------------------------------------
    ' Provide the milliseconds the timer itself consumes when it not does nay
    ' actual measurement. This overhead is computed when the class is
    ' initiated and deducted from any measurement result. This is the way the
    ' timer service provided (TimerStart, TimerEnd) will return a figure with
    ' an unavoidable precision failure which is a fraction of a milliseconds.
    ' ------------------------------------------------------------------------
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cTicksAtStart
    For i = 1 To lLoop
        TimerStart
        ' no time to be taken
        TimerEnd
    Next i
    QueryPerformanceCounter cTicksCurrent
    cTimerOverhead = (((cTicksCurrent - cTicksAtStart) / lLoop) / cFrequency) * 1000
    sTestFileExtension = "txt"
    cExecTimeMsecs = 0
    
    siInstructionsPosTop = 10
    siInstructionsPosLeft = 10
    
End Sub

Private Sub Class_Terminate()

    Dim v As Variant

    Set dctTestSummary = Nothing
    Set dctResult = Nothing
    Set dctResultExpected = Nothing
    CleanUp
    Set cllTempTestItems = Nothing

#If mMsg = 1 Then
    If Not mMsg.Instances Is Nothing Then
        For Each v In mMsg.Instances
            On Error Resume Next
            Unload mMsg.Instances(v)
        Next v
    End If
#End If
    
End Sub

Private Property Get Arry(Optional ByRef c_arr As Variant, _
                          Optional ByVal c_index As Long = -1) As Variant
' ----------------------------------------------------------------------------
' Universal array read procedure. Returns Null when a given array (c_arr) is
' not allocated or a provided index is beyond/outside current dimensions.
' ----------------------------------------------------------------------------
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
' ----------------------------------------------------------------------------
' Universal array add/update procedure, avoiding any prior checks whether
' allocated, empty not yet existing, etc.
' - Adds an item (c_var) to an array (c_arr) when no index is provided or when
'   the index lies beyond UBound
' - When an index is provided, the item is inserted/updated at the given
'   index - even when the array yet doesn't exist or is not yet allocated.
' ----------------------------------------------------------------------------
    Const PROC = "Arry-Let"
    
    Dim bIsAllocated As Boolean
    
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

Private Property Let Dict(Optional ByRef d_dct As Dictionary, _
                          Optional ByVal d_key As Variant, _
                                   ByVal d_item As Variant)
    If d_dct Is Nothing Then Set d_dct = New Dictionary
    If d_dct.Exists(d_key) Then d_dct.Remove d_key
    d_dct.Add d_key, d_item
    
End Property

Friend Property Let ExcludeFromComparison(ByVal s As String):   sExcludeFromComparison = s:                     End Property

Private Property Get Failed() As Boolean:                       Failed = sResultEvaluated = RESULT_FAILED:      End Property

Friend Property Get InstructionsPosLeft() As Single:            InstructionsPosLeft = siInstructionsPosLeft:    End Property

Friend Property Let InstructionsPosLeft(si As Single):          siInstructionsPosLeft = si:                     End Property

Friend Property Get InstructionsPosTop() As Single:             InstructionsPosTop = siInstructionsPosTop:      End Property

Friend Property Let InstructionsPosTop(si As Single):           siInstructionsPosTop = si:                      End Property

Friend Property Get InstructionsTitle() As String:          InstructionsTitle = siInstructionsTitle:                    End Property

Private Property Get LogColAlign(Optional ByVal c_col As Long) As String
' ----------------------------------------------------------------------------
' Return a column's width, whereby an available explicit specification
' supercedes a width iplicitely derived from a string.
' ----------------------------------------------------------------------------
    LogColAlign = Arry(aLogColAligns, c_col)
End Property

Private Property Let LogColAlign(Optional ByVal c_col As Long, _
                                       ByVal c_align As String)
    If Arry(aLogColAligns, c_col) = vbNullString Then
        Arry(aLogColAligns, c_col) = c_align
    End If
End Property

Private Property Get LogColFill(Optional ByVal c_col As Long) As String
        LogColFill = Arry(aLogColFills, c_col)
End Property

Private Property Let LogColFill(Optional ByVal c_col As Long, _
                                      ByVal c_fill As String)
    If Arry(aLogColFills, c_col) = vbNullString Then
        Arry(aLogColFills, c_col) = c_fill
    End If
End Property

Private Property Get LogCols() As Long
    LogCols = Max(lCols, ArrayNoOfElements(aLogColAligns), ArrayNoOfElements(aLogColWidths))
End Property

Private Property Let LogCols(ByVal c_cols As Long):    lCols = c_cols: End Property

Private Property Get LogColsDelim() As String:                  LogColsDelim = sLogColsDelim:           End Property

Private Property Let LogColsDelim(ByVal s As String):           sLogColsDelim = s:                      End Property

Private Property Get LogColsHeader(Optional ByVal c_line As Long) As Variant
    LogColsHeader = Arry(aLogColsArrangedHeaders, c_line)
End Property

Private Property Let LogColsHeader(Optional ByVal c_line As Long, _
                                            ByVal c_cols_header As Variant)
' -----------------------------------------------------------------------------------
' Saves provided column headers (c_cols_header), provided as array, collection, or as
' a string delimited by vbCrLf, ", " or "|&|", as string with the column headers
' arranged in columns, considering a maximized column width (in order to adapt any
' already specified width to the effective header's width), to the aColHeaders array.
' -----------------------------------------------------------------------------------
    Dim s           As String
    Dim i           As Long
    Dim iCol        As Long
    Dim sColHeader  As String
    Dim sColFill    As String
    Dim sItem       As String
    Dim aHeaders    As Variant
    Dim sDelim      As String
    
    c_line = c_line
    aHeaders = VarAsArray(c_cols_header)
    sColFill = " "

    For i = LBound(aHeaders) To UBound(aHeaders)
        iCol = iCol + 1
        sColHeader = Trim$(aHeaders(i)) ' ensure centered regardless any leading or trailing spaces
        LogColWidth(iCol) = Len(sColHeader)
        sItem = AlignString(sColHeader, "C", LogColWidth(iCol), sColFill, LogColsMargin, iCol, True)
        s = ConcatenateStrings(s, sItem, LogColsDelim, sDelim)
    Next i
    Arry(aLogColsArrangedHeaders) = s ' add s to the array
    
End Property

Private Property Get LogColsMargin() As String
    If sLogColsDelim <> " " Then LogColsMargin = " " Else LogColsMargin = vbNullString
End Property

Private Property Let LogColsSpecs(ByVal c_specs As String)

    Dim aSpecs  As Variant
    Dim iCol    As Long
    Dim v       As Variant
    
    aSpecs = StringAsArray(c_specs)
    
    For Each v In aSpecs
        iCol = iCol + 1
        LogColSpec iCol, v
    Next v
    
End Property

Private Property Get LogColWidth(Optional ByVal c_col As Long) As Long
' ----------------------------------------------------------------------------
' The (maximized) column width is used for:
' 1. The Title line (when one had been specified)
' 2. The column header lines and the header underline
' 3. The final re-assembling of the very first column arranged Entry line
'    (which triggered the above two)
' ----------------------------------------------------------------------------
    LogColWidth = Arry(aLogColWidths, c_col)

End Property

Private Property Let LogColWidth(Optional ByVal c_col As Long, _
                                          ByVal c_width As Long)
' ----------------------------------------------------------------------------
' Initially a column's width may or may not be specified explicitly by
' 'LogColsWidth'. Provided the truncate option is False the width is
' subsequently maximized by the aligned item which may include a possible
' (implicitly specified) filling. The final column width is the one specified
' explicitly or derived implicitly plus the length of the specified filling.
' I.e. an item specified "xxxx .:" will finally end up in a column width of 10.
' ----------------------------------------------------------------------------
    Arry(aLogColWidths, c_col) = Max(Arry(aLogColWidths, c_col), c_width)
    
End Property

Friend Property Get ModeRegression() As Boolean:            ModeRegression = bModeRegression:                           End Property

Friend Property Let ModeRegression(ByVal b As Boolean)
    
    bModeRegression = b
#If mErH Then
    mErH.Regression = b
#End If

End Property

Friend Property Get NameTestFile() As String:               NameTestFile = sNameTestFile:                               End Property

Friend Property Get NameTestResultExpectedFile() As String: NameTestResultExpectedFile = sNameTestResultExpectedFile:   End Property

Friend Property Get NameTestResultFile() As String:         NameTestResultFile = sNameTestResultFile:                   End Property

Private Property Get Passed() As Boolean:                   Passed = sResultEvaluated = RESULT_PASSED:                  End Property

Friend Property Get Result() As Variant
' ------------------------------------------------------------------------------
' Returns the most current registered test result.
' ------------------------------------------------------------------------------
    
    If VBA.IsObject(dctResult(VerificationId)) _
    Then Set Result = dctResult(VerificationId) _
    Else Result = dctResult(VerificationId)

End Property

Friend Property Let Result(ByVal r_var As Variant)
' ------------------------------------------------------------------------------
' Register in dctResult a new test result. When the corresponding expected
' result has also already been registered, provide the final test result.
' ------------------------------------------------------------------------------
    Const PROC = "Result-Let"
        
    Select Case True
        Case sVerification = vbNullString And sTitle = vbNullString
            Err.Raise AppErr(1), ErrSrc(PROC), ErrorMsgMissingVerification
        Case sVerification = vbNullString And sTitle <> vbNullString
            sVerification = sTitle
    End Select
    With dctResult
        If .Exists(VerificationId) Then .Remove VerificationId
        .Add VerificationId, r_var
    End With
    
    If dctResultExpected.Exists(VerificationId) Then ResultEvaluation
    
End Property

Friend Property Get ResultExpected() As Variant
' ------------------------------------------------------------------------------
' Returns the most current registered expected test result.
' ------------------------------------------------------------------------------
    
    If VBA.IsObject(dctResultExpected(VerificationId)) _
    Then Set ResultExpected = dctResultExpected(VerificationId) _
    Else ResultExpected = dctResultExpected(VerificationId)

End Property

Friend Property Let ResultExpected(ByVal r_var As Variant)
' ------------------------------------------------------------------------------
' Register in dctResultExpected a new expected test result. When the
' corresponding test result has also already been registered, the final test
' result is provided.
' ------------------------------------------------------------------------------
    Const PROC = "ResultExpected-Let"
    
    Select Case True
        Case sVerification = vbNullString And sTitle = vbNullString
            Err.Raise AppErr(1), ErrSrc(PROC), ErrorMsgMissingVerification
        Case sVerification = vbNullString And sTitle <> vbNullString
            sVerification = sTitle
    End Select
    
    With dctResultExpected
        If .Exists(VerificationId) Then .Remove VerificationId
        .Add VerificationId, r_var
    End With
    
    If dctResult.Exists(VerificationId) Then ResultEvaluation
    
End Property

Friend Property Let TempTestItem(Optional ByVal t_test_no As String = vbNullString, _
                                          ByVal t_item As Variant)
' ------------------------------------------------------------------------------
' - When no test-number (t_test_no) is provided and the provided test item
'   (t_item) exists (may be a file, a folder, or a Workbook), a record is kept
'   in a file (sTestItems) for being removed (when a file or a folder)
'   or closed (when a Workbook) with the CleanUp service.
' - When a test-number is provided the full default name of an existing test
'   file or folder becomes the current test item.
' ------------------------------------------------------------------------------
    Const PROC = "TempTestItem-Let"
        
    If t_test_no = vbNullString Then
        '~~ Keep a record of the test item
        Select Case True
            Case TypeName(t_item) = "String"
                cllTempTestItems.Add t_item
                vTempTestItem = t_item
            Case VBA.IsObject(t_item)
                cllTempTestItems.Add t_item
                Set vTempTestItem = t_item
            Case Else:                          Err.Raise AppErr(1), ErrSrc(PROC), "Unexpected kind of test item (record not kept)!"
        End Select
    Else
        '~~ Provide the test item
        vTempTestItem = TestFileFullName
    End If
    
End Property

Friend Property Let TestedComp(ByVal s As String):                  sTestedComp = s:                                            End Property

Friend Property Let TestedProc(ByVal s As String):                  sTestedProc = s:                                            End Property

Friend Property Let TestedProcType(ByVal s As String):              sTestedProcType = s:                                        End Property

Friend Property Get TestFileExtension() As String:                  TestFileExtension = sTestFileExtension:                     End Property

Friend Property Let TestFileExtension(ByVal s As String):           sTestFileExtension = Replace("." & s, "..", "."):           End Property

Friend Property Get TestFolder() As String:                         TestFolder = sTestFolder:                                   End Property

Friend Property Let TestFolder(ByVal s As String)
    
    If sTestFolder <> vbNullString Then
        If sTestFolder <> s Then
            If FSo.FolderExists(s) Then FSo.DeleteFolder s
        End If
    End If
    sTestFolder = s
    If Not FSo.FolderExists(sTestFolder) _
    Then FSo.CreateFolder sTestFolder
    ChDir sTestFolder

End Property

Private Property Get TestFullName() As String
    
    Dim s As String
    
    Select Case Left(sTestedProcType, 1)
        Case "F":   s = "Function"
        Case "M":   s = "Method  "
        Case "P":   s = "Property"
        Case "S":   s = "Sub     "
    End Select
    TestFullName = sTestId & "-" & lVerificationNumber & " " & s & " " & sTestedComp & "." & sTestedProc & ": " & sVerification

End Property

Friend Property Get TestId() As String:                             TestId = sTestId:                                           End Property

Friend Property Let TestId(ByVal s As String)
' ------------------------------------------------------------------------------
' Essential, obligatory distinct identification of a test, provided as the very
' first Property for each test. Resets a previous test to prepare for a new one.
' ------------------------------------------------------------------------------
    
    If sTestId <> s Then
        sTestId = s
        dctResult.RemoveAll
        dctResultExpected.RemoveAll
        vTempTestItem = vbNullString
        lVerificationNumber = 0
        sVerification = vbNullString
    End If
    
End Property

Friend Property Get TestResult() As String:                         TestResult = sResultEvaluated:                              End Property

Private Property Let TestResultSummaryLogFile(ByVal s As String)
    
    If sTestResultSummaryLogFile <> vbNullString Then
        If FSo.FileExists(sTestResultSummaryLogFile) Then
            FSo.DeleteFile sTestResultSummaryLogFile
        End If
    End If
    sTestResultSummaryLogFile = s
    
End Property

Friend Property Get TimerExecTimeMsecs() As Currency:               TimerExecTimeMsecs = cExecTimeMsecs:                        End Property

Friend Property Get TimerOverhead() As Currency:                    TimerOverhead = cTimerOverhead:                             End Property

Friend Property Get Title() As String:                              Title = sTitle:                                             End Property

Friend Property Let Title(ByVal s As String):                       sTitle = s:                                                 End Property

Friend Property Let TitleRegression(ByVal s As String):             sTitleRegression = s:                                       End Property

Friend Property Let Verification(ByVal s As String)
' ----------------------------------------------------------------------------
' Obligatory prior any .Result/.ResultExpected statements. However, when a
' test identified by a u nique .TestId has only one .Result/.ResultExpected
' for a missing Verification the Title is used alternatively - when it
' had been provided. When neither has been provided an error is raised. For
' any addidtional verification (.Result/.ResultExpected) the provision of a
' description (this property) is obligatory.
' Note: Verifications are indicated by a .number attached to the TestId.
' ----------------------------------------------------------------------------
    
    sVerification = s
    lVerificationNumber = lVerificationNumber + 1
    sNameTestFile = sTestFolder & "\Test-" & sTestId & "-" & lVerificationNumber & sTestFileExtension
    sNameTestResultFile = sTestFolder & "\Test-" & sTestId & "-Result-" & lVerificationNumber & sTestFileExtension
    sNameTestResultExpectedFile = sTestFolder & "\Test-" & sTestId & "-Result-Expected-" & lVerificationNumber & sTestFileExtension

End Property

Private Function AlignCntr(ByVal a_string As String, _
                           ByRef a_length As Long, _
                  Optional ByVal a_fill As String = " ", _
                  Optional ByVal a_margin As String = vbNullString, _
                  Optional ByVal a_col As Long = 0, _
                  Optional ByVal a_col_arranged = False) As String
' ----------------------------------------------------------------------------
' Returns a string (a_string) centered within a guven width (a_length). When
' truncate (a_truncate) is False (the default), the returned with (a_with)
' may be extended to the width of the string (a_steing). The result may not be
' centered exactly.
' ----------------------------------------------------------------------------
    Const PROC = "AlignCntr"
    
    On Error GoTo eh
    Dim lLoop       As Long
    Dim s           As String
    Dim sLeft       As String
    Dim sRight      As String
    Dim sFill       As String
    Dim lWidth      As Long

    If a_length = 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A zero length is nor supported by alignment centered!"

    Select Case a_fill
        Case " ":        sFill = " ": sLeft = " ":          sRight = " "
        Case "-":        sFill = "-": sLeft = vbNullString: sRight = vbNullString
        Case " -", "- ": sFill = "-": sLeft = " ":          sRight = " "
        Case "=":        sFill = "=": sLeft = vbNullString: sRight = vbNullString
        Case " =", "= ": sFill = "=": sLeft = " ":          sRight = " "
    End Select
    
    s = TruncateIfApplicable(a_string, a_length, a_col, LogCols, "C")
    lWidth = AlignFinalFullColWidthExMargin(a_fill, a_length, "C", a_col_arranged)
    
    Do While Len(s) < lWidth
        lLoop = lLoop + 1
        Select Case True
            Case Len(s) = lWidth
                Exit Do
            Case Len(s & sRight) <= lWidth
                s = s & sRight
                sRight = sFill
                If Len(s) < lWidth Then
                    s = sLeft & s ' add fill left
                    sLeft = sFill
                End If
        End Select
        If lLoop > lWidth Then
            Stop
        End If
    Loop
    AlignCntr = a_margin & s & a_margin

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AlignFinalFullColWidthExMargin(ByVal a_fill As String, _
                                                ByVal a_length As Long, _
                                                ByVal a_align As String, _
                                       Optional ByVal a_col_arranged As Boolean = False) As Long
' -----------------------------------------------------------------------------------
' When another fill but a single space The alignment always includes at least one
' fill. Thus, in this case a final column width is returned which includes the lenght
' of the fill.
' -----------------------------------------------------------------------------------
    Dim l As Long
    
    l = a_length ' default
    If a_fill <> " " And a_col_arranged Then
        l = l + Len(a_fill)   ' one in any case
        If a_align = "C" Then l = l + Len(a_fill) ' one fill at the left and another one at the right
    End If
    AlignFinalFullColWidthExMargin = l
  
End Function

Private Function AlignLeft(ByRef a_string As String, _
                           ByRef a_length As Long, _
                  Optional ByVal a_fill As String = " ", _
                  Optional ByVal a_margin As String = vbNullString, _
                  Optional ByVal a_col As Long = 0, _
                  Optional ByVal a_col_arranged As Boolean = False) As String
' -----------------------------------------------------------------------------------
' Left aligns a string (a_string) by:
' - adding fill characters to the right up to the specified length (a_length),
'   whereby the filling may start with a single space
' - enclose the result in left and right margins - which default to a vbNullString
' Specifics:
' - Spaces provided with a string alrteady (a_string) are preserved
' - The final length is the specified length plus a left and a right marging - which
'   may be a vbNullString
' - It is considered that the string may be an column item which may look ate the end
'   "| xxxx" when a column delimiter  | is used, else just "xxxx" when the column
'   delimiter is already a single space. In that case the margin (a_margin) would be
'   a vbNullString.
' Precondition: The string (a_string) not/no longer contains any implicit alignment
'               specs.
' -----------------------------------------------------------------------------------

    Dim s           As String
    Dim sFillStart  As String
    Dim sFill       As String
    Dim sFillEnd    As String
    Dim lLoop       As Long
    Dim lWidth      As Long
    
    Select Case a_fill
        Case " ":   sFillStart = vbNullString: sFill = " ": sFillEnd = " "
        Case ".":   sFillStart = vbNullString: sFill = ".": sFillEnd = vbNullString
        Case " .":  sFillStart = " ":          sFill = ".": sFillEnd = vbNullString
        Case ".:":  sFillStart = vbNullString: sFill = ".": sFillEnd = ":"
        Case " .:": sFillStart = " ":          sFill = ".": sFillEnd = ":"
    End Select
        
    s = TruncateIfApplicable(a_string, a_length, a_col, LogCols, "L")
    lWidth = AlignFinalFullColWidthExMargin(a_fill, a_length, "L", a_col_arranged)
    
    Do
        lLoop = lLoop + 1
        Select Case True
            Case Len(s) >= lWidth:                                Exit Do
            Case Len(s & sFillEnd) = lWidth:                      s = s & sFillEnd
            Case Len(s & sFillStart) = lWidth:                    s = s & sFillStart:         sFillStart = vbNullString
            Case Len(s & sFillStart & sFill & sFillEnd) < lWidth: s = s & sFillStart & sFill: sFillStart = vbNullString
            Case Else:                                            s = s & sFill
        End Select
        If lLoop > lWidth Then
            Stop
        End If
    Loop
    AlignLeft = a_margin & s & a_margin
    
End Function

Private Function AlignRght(ByVal a_string As String, _
                  Optional ByRef a_length As Long = 0, _
                  Optional ByVal a_fill As String = " ", _
                  Optional ByVal a_margin As String = vbNullString, _
                  Optional ByVal a_col As Long = 0, _
                  Optional ByVal a_col_arranged As Boolean = False) As String
' -----------------------------------------------------------------------------------
' Returns a string (a_string) aligned right with at the left filled (a_fill) in a
' given length (a_lenght), enclosed in specified margins (a_margin).
' -----------------------------------------------------------------------------------
    Const PROC = "AlignRght"
    
    On Error GoTo eh
    Dim lLoop       As Long
    Dim lWidth      As Long
    Dim s           As String
    Dim sFill       As String
    Dim sFillEnd    As String
    Dim sFillStart  As String
    
    If a_length = 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A zero length is nor supported by alignment right!"
    
    If a_fill = vbNullString Then a_fill = " "
    Select Case a_fill
        Case " ":           sFillStart = vbNullString: sFill = " ": sFillEnd = " "
        Case ".":           sFillStart = vbNullString: sFill = ".": sFillEnd = vbNullString
        Case " .", ". ":    sFillStart = " ":          sFill = ".": sFillEnd = vbNullString
        Case "-":           sFillStart = vbNullString: sFill = "-": sFillEnd = vbNullString
        Case " -", "- ":    sFillStart = " ":          sFill = "-": sFillEnd = vbNullString
    End Select
    
    s = TruncateIfApplicable(a_string, a_length, a_col, LogCols, "R")
    lWidth = AlignFinalFullColWidthExMargin(a_fill, a_length, "R", a_col_arranged)
    
    Do
        lLoop = lLoop + 1
        Select Case True
            Case Len(s) >= lWidth:                                Exit Do
            Case Len(sFillStart & s) = lWidth:                    s = sFillStart & s:           sFillStart = vbNullString
            Case Len(sFillEnd & s) = lWidth:                      s = sFillEnd & s              ' last fill add
            Case Len(sFillEnd & sFill & sFillStart & s) < lWidth: s = sFill & sFillStart & s:   sFillStart = vbNullString
            Case Else:                                            s = sFill & s
        End Select
        If lLoop > lWidth Then
            Stop
        End If
    Loop
    AlignRght = a_margin & s & a_margin
                   
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AlignString(ByVal a_string As String, _
                   Optional ByVal a_align As String = "L", _
                   Optional ByVal a_length As Long = 0, _
                   Optional ByVal a_fill As String = " ", _
                   Optional ByVal a_margin As String = vbNullString, _
                   Optional ByVal a_col As Long = 0, _
                   Optional ByVal a_col_arranged As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns a string (a_string)
' - in the length (a_length) enclosed in margins (a_margin) whereby a non
'   provided width defaults to the string's (a_string) width
' - aligned (a_align), when not provided, defaults to the alignment implicitly
'   specified
' - filled with (a_fill), defaults to a single space when not provided.
' Specifics:
' - When a margin is provided, the final length will be the specified length
'   plus the length of a left and a right margin. A margin is typically used
'   when the string is one item of serveral organized in columns when the
'   column delimiter is a vbNullString. When the column delimiter is a |
'   usually a marign of a single space is used
' - The string (a_strn) may contain already leading or trailing spaces of
'   the left are preserved when the string is left aligned and the right
'   spaces are preserved when the string is right aligned.
'   spaces.
' - The function is also used to align items arragend in columns.
'
' Note: The argument a_col_arranged is only used for a string which is aligned
'       within a column. See LogColsArrangedHeader and LogColsArrangedItems.
'
' W. Rauschenberger, Berlin Jun 2024
' ----------------------------------------------------------------------------
    Const PROC = "AlignString"
    
    On Error GoTo eh
    Dim lLength As Long
    
    If a_string = vbNullString Then
        '~~ A non provided string results in one filled with a_fill
        '~~ enclosed in margins (a_margin)
        AlignString = a_margin & String$(a_length, a_fill) & a_margin
        GoTo xt
    End If
        
    If a_align = vbNullString Then
        '~~ When no alignment is indicated it default to left with a single space fill.
        a_align = "L"
        a_fill = " "
    End If
        
    If a_length = 0 _
    Then lLength = Len(a_string) _
    Else lLength = a_length

    If a_fill = vbNullString Then a_fill = " "
    Select Case a_align
        Case "L":   AlignString = AlignLeft(a_string, lLength, a_fill, a_margin, a_col, a_col_arranged)
        Case "R":   AlignString = AlignRght(a_string, lLength, a_fill, a_margin, a_col, a_col_arranged)
        Case "C":   AlignString = AlignCntr(a_string, lLength, a_fill, a_margin, a_col, a_col_arranged)
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

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

Private Function ArrayAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In v_items
            .Add v
        Next v
    End With
    Set ArrayAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function ArrayAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In v_items
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set ArrayAsDictionary = dct
    Set dct = Nothing
    
End Function

Private Function ArrayAsFile(ByVal v_items As Variant, _
                    Optional ByRef v_file_name As String = vbNullString, _
                    Optional ByVal v_file_append As Boolean = True) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
      
    If v_file_name = vbNullString Then v_file_name = TempFileFullName
    StringAsFile ArrayAsString(v_items), v_file_name, v_file_append
    Set ArrayAsFile = FSo.GetFile(v_file_name)
    
End Function

Private Function ArrayAsString(ByVal c_array As Variant, _
                      Optional ByVal c_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (c_array) as string with the items delimited (c_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArrayAsString = Join(c_array, c_delim)
End Function

Private Function ArrayErase(ByRef c_arr As Variant)
    If IsArray(c_arr) Then Erase c_arr
End Function

Private Function ArrayHasElements(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the array (a_arr) has at least one element.
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    On Error Resume Next
    v = a_arr(LBound(a_arr))
    ArrayHasElements = Err.Number = 0
    Err.Clear

End Function

Private Function ArrayIsAllocated(ByVal Arry As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = UBound(Arry) >= LBound(Arry)
    On Error GoTo -1
    
End Function

Private Function ArrayNoOfElements(ByVal a_arr As Variant) As Long

    If ArrayHasElements(a_arr) Then
        If LBound(a_arr) = 0 _
        Then ArrayNoOfElements = UBound(a_arr) - LBound(a_arr) _
        Else ArrayNoOfElements = UBound(a_arr) - LBound(a_arr) + 1
    End If
    
End Function

Public Sub AssertedErrors(ParamArray c_app_errs() As Variant)
' ----------------------------------------------------------------------------
' When mErH is not installed (Cond. Comp. Arg. `mErh=1`) this is ignored.
' ----------------------------------------------------------------------------
    Const PROC = "AssertedErrors"

#If mErH Then
    On Error GoTo xt
    If ArrayIsAllocated(c_app_errs) Then
        Select Case UBound(c_app_errs)
            Case 0: mErH.Asserted c_app_errs(0)
            Case 1: mErH.Asserted c_app_errs(0), c_app_errs(1)
            Case 2: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2)
            Case 3: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3)
            Case 4: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3), c_app_errs(4)
            Case 5: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3), c_app_errs(4), c_app_errs(5)
            Case 6: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3), c_app_errs(4), c_app_errs(5), c_app_errs(6)
            Case 7: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3), c_app_errs(4), c_app_errs(5), c_app_errs(6), c_app_errs(7)
            Case 8: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3), c_app_errs(4), c_app_errs(5), c_app_errs(6), c_app_errs(7), c_app_errs(8)
            Case 9: mErH.Asserted c_app_errs(0), c_app_errs(1), c_app_errs(2), c_app_errs(3), c_app_errs(4), c_app_errs(6), c_app_errs(6), c_app_errs(7), c_app_errs(8), c_app_errs(9)
            Case Else
                Err.Raise AppErr(1), ErrSrc(PROC), "Max number of assertable errors (10) exceeded!"
        End Select
    End If

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
#Else
    Err.Raise AppErr(1), ErrSrc(PROC), "The AssertedErrors service is only available when the mErH " & _
                                       "component is installed and indicated by the Cond. Comp. Arg. `mErH = 1`"
#End If
End Sub

Private Function BooleanAsString(ByVal b As Boolean) As String
    If b Then BooleanAsString = "TRUE" Else BooleanAsString = "FALSE"
End Function

Public Sub BoTP()
' ------------------------------------------------------------------------------
' Begin of Testesd Procedure: Interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP TestFullName
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP TestFullName
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP TestFullName
#End If
End Sub

Private Function checkHwnds(ByRef xlApps() As Application, hWnd As LongPtr) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "checkHwnds"
    
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

Public Sub CleanUp(ParamArray t_names() As Variant)
' ------------------------------------------------------------------------------
' Removes/Deletes any item collected in cllTempTestItems.
' ------------------------------------------------------------------------------
    Const PROC = "CleanUp"
    
    On Error GoTo eh
    Dim i   As Long
    Dim fle As File
    Dim v   As Variant
    
    With FSo
        For Each fle In .GetFolder(sTestFolder).Files
            '~~ Remove any temporary test file
            If fle.Name Like "rad*" Then
                .DeleteFile fle.Path
            End If
            '~~ Remove when explicitly provided named
            If ArrayIsAllocated(t_names) Then
                For Each v In t_names
                    Select Case True
                        Case fle.Name Like v, _
                             fle.Name Like Replace(v, v & "**", v & "*")
                            On Error Resume Next
                            If .FileExists(fle.Path) Then
                                If Err.Number = 0 Then
                                    .DeleteFile fle.Path
                                End If
                            End If
                    End Select
                Next v
            End If
        Next fle
          
        If cllTempTestItems.Count = 0 Then GoTo xt
        
        For i = cllTempTestItems.Count To 1 Step -1
            If cllTempTestItems(i) Like "*.xl*" Then
                On Error Resume Next
                Application.Workbooks(cllTempTestItems(i)).Close False ' close Workbook without save (keep its preconditions)
                If Err.Number = 0 Then cllTempTestItems.Remove i
            End If
        Next i
        
        If cllTempTestItems.Count = 0 Then GoTo xt
        For i = cllTempTestItems.Count To 1 Step -1
            Select Case True
                Case .FileExists(cllTempTestItems(i))
                    .DeleteFile cllTempTestItems(i)
                Case .FolderExists(cllTempTestItems(i))
                    If Right(cllTempTestItems(i), 1) = "\" Then v = Left(cllTempTestItems(i), Len(cllTempTestItems(i)) - 1)
                    On Error Resume Next
                    .DeleteFolder cllTempTestItems(i)
            End Select
            cllTempTestItems.Remove i
        Next i
    End With


xt: Exit Sub

eh: Err.Raise Err.Number, ErrSrc(PROC), Err.Description
End Sub

Private Function CollectionAsArray(ByVal v_items As Collection) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    
    For Each v In v_items
        If IsObject(v, s) _
        Then Arry(arr) = s _
        Else Arry(arr) = v
    Next v
    CollectionAsArray = arr
    
End Function

Private Function CollectionAsDictionary(ByVal v_items As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In v_items
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set CollectionAsDictionary = KeySort(dct)
    Set dct = Nothing

End Function

Private Function CollectionAsFile(ByVal c_items As Collection, _
                        Optional ByRef c_file As String = vbNullString, _
                        Optional ByVal c_file_append As Boolean = True) As File
' ----------------------------------------------------------------------------
' Transfers the items of a Collection (c_coll) as records/lines to a
' file (c_file_name), optionally appended (c_file_append).
' ----------------------------------------------------------------------------

    If c_file = vbNullString Then c_file = TempFileFullName
    StringAsFile CollectionAsString(c_items), c_file, c_file_append
    Set CollectionAsFile = FSo.GetFile(c_file)

End Function

Private Function CollectionAsString(ByVal c_coll As Collection, _
                           Optional ByRef c_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an item is an object its
' Name property is used (an error is raised when the object has no Name
' property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sName   As String
    Dim sSplit  As String
    Dim v       As Variant
    Dim v2      As Variant
    
    If c_split = vbNullString Then c_split = vbCrLf
    For Each v In c_coll
        Select Case True
            Case IsObject(v, sName)
                s = s & sSplit & sName
                sSplit = c_split
            Case TypeName(v) Like "*()"
                For Each v2 In v
                    s = s & sSplit & CStr(v2)
                    sSplit = c_split
                Next v2
            Case Else
                s = s & sSplit & v
                sSplit = c_split
        End Select
    Next v
    CollectionAsString = s

End Function

Private Function ConcatenateStrings(ByVal c_string As String, _
                                    ByVal c_item As String, _
                                    ByVal c_conc As String, _
                                    ByRef c_delim As String) As String
    ConcatenateStrings = c_string & c_delim & c_item
    c_delim = c_conc
End Function

Private Function DictionaryAsArray(ByVal v_items As Dictionary) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    
    For Each v In v_items
        If IsObject(v, s) _
        Then Arry(arr) = s _
        Else Arry(arr) = v
    Next v
    DictionaryAsArray = arr
    
End Function

Private Function DictionaryAsCollection(ByVal v_items As Dictionary) As Collection
' ----------------------------------------------------------------------------
' Returns the Dictionary (v_items) with the keys in ascending order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In v_items
            .Add v
        Next v
    End With
    Set DictionaryAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function DictionaryAsDictionary(ByVal v_items As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns the Dictionary (v_items) with the keys in ascending order.
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    With dct
        For Each v In v_items
            .Add v, vbNullString
        Next v
    End With
    Set DictionaryAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function DictionaryAsFile(ByVal v_items As Dictionary, _
                         Optional ByRef v_file_name As String = vbNullString, _
                         Optional ByVal v_file_append As Boolean = True) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    If v_file_name = vbNullString Then v_file_name = TempFileFullName
    StringAsFile DictionaryAsString(v_items), v_file_name, v_file_append
    Set DictionaryAsFile = FSo.GetFile(v_file_name)

End Function

Private Function DictionaryAsString(ByVal d_items As Dictionary, _
                           Optional ByRef d_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a Dictionary's keys as string with each key delimited by a vbCrLf.
' Keys are converted into a string, if an item is an object its Name property
' is used. An error is raised when the object has no Name property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sName   As String
    Dim sSplit  As String
    
    If d_split = vbNullString Then d_split = vbCrLf
    For Each v In d_items
        If IsObject(v, sName) _
        Then s = s & sSplit & sName _
        Else: s = s & sSplit & v
        sSplit = d_split
    Next v
    DictionaryAsString = s

End Function

Public Sub DsplyFile(ByVal d_file_name As String)
    ShellRun d_file_name
End Sub

Public Sub EoTP()
' ------------------------------------------------------------------------------
' End of Tested Procedure: Interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP TestFullName, ">" & TestResult
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP TestFullName, ">" & TestResult
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP TestFullName, ">" & TestResult
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
            Else: ErrType = "VB Runtime Error "
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

Private Function ErrorMsgMissingVerification() As String
    ErrorMsgMissingVerification = "For the test " & sTestId & " neither a Verification nor a Title had been specified!" & "||" & _
                                  "Within/for an individual test number a dedicated "".Verification"", " & _
                                  "specifying the quality/extent of the expected test result, is mandatory for each pair of .ResultExpected/.Result " & _
                                  "statements. However, when no Verification had been specified a specified Title is used instead."
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsTestAid." & sProc
End Function

Private Function Exclude(ByVal e_strng As String) As String
' ------------------------------------------------------------------------------
' Returns a string (e_string) with the "exclude from result evaluation" string
' excluded. Whereby the string (e_string) contains substrings delimited by a
' vbCrLf.
' ------------------------------------------------------------------------------
                         
    Dim a       As Variant
    Dim b       As Boolean
    Dim i       As Long
    Dim s       As String
    Dim lPos    As Long
    Dim lLen    As Long
    Dim bLike   As Boolean
    
    If sExcludeFromComparison = vbNullString Then
        Exclude = e_strng
        GoTo xt
    End If
    
    bLike = InStr(e_strng, "*") <> 0
    a = StringAsArray(e_strng, vbCrLf)
    For i = LBound(a) To UBound(a)
        s = a(i)
        If s Like Replace("*" & sExcludeFromComparison & "*", "**", "*") Then
            lPos = ExcludePosition(s, sExcludeFromComparison, bLike)
            If lPos <> 0 Then
                If bLike Then
                    lLen = Len(sExcludeFromComparison)
                    a(i) = ExcludeSubstring(s, lPos, lLen)
                Else
                    a(i) = Right(s, Len(s) - lPos)
                End If
                b = True
            End If
        End If
    Next i
    
    If b _
    Then Exclude = Join(a, vbCrLf) _
    Else Exclude = e_strng
    
xt: Exit Function

End Function

Private Function ExcludePosition(ByVal e_main As String, _
                                 ByVal e_pattern As String, _
                        Optional ByVal e_like As Boolean = True) As Integer
' ------------------------------------------------------------------------------
' Returns the position of a "Like" pattern (e_pattern) in a string (e_main) when
' like = True (e_like). When like = False then the position of the pattern in
' the string (e_main).
' Returns 0 when e_pattern is not contained in the string (e_main).
' ------------------------------------------------------------------------------
    Dim i        As Integer
    Dim sPattern As String
    
    ExcludePosition = 0
    For i = 1 To Len(e_main)
        sPattern = Mid(e_main, i, Len(e_pattern))
        If sPattern Like e_pattern Then
            If e_like _
            Then ExcludePosition = i _
            Else ExcludePosition = i + Len(e_pattern) - 1
            GoTo xt
        End If
    Next i

xt: Exit Function
End Function

Private Function ExcludeSubstring(ByVal r_main As String, _
                                  ByVal r_start As Long, _
                                  ByVal r_length As Long) As String
' ------------------------------------------------------------------------------
' Returns a string (r_main) with a substring beginning (r_start) and a length
' (r_length) excluded.
' ------------------------------------------------------------------------------
    
    Dim sLeft   As String
    Dim sRight  As String
    
    sLeft = Left(r_main, r_start - 1)
    sRight = Mid(r_main, r_start + r_length)
    ExcludeSubstring = sLeft & sRight
    
End Function

Private Function FileAsArray(ByVal f_file As Variant, _
                    Optional ByVal f_empty_excluded) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim sSplit  As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    For Each v In Split(FileAsString(f_file, sSplit), sSplit)
        If f_empty_excluded Then
            If Trim$(v) <> vbNullString Then Arry(arr) = v
        Else
            Arry(arr) = v
        End If
    Next v
    FileAsArray = arr
    
End Function

Private Function FileAsCollection(ByVal v_items As File) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "FileAsCollection"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim sSplit  As String
    
    With cll
        For Each v In Split(FileAsString(v_items.Path, sSplit), sSplit)
            .Add v
        Next v
    End With
    Set FileAsCollection = cll
    Set cll = Nothing

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsDictionary(ByVal v_items As File) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the file's records (vulgo text lines) are returned as
'            Directory keys, lines will be unified. I e. line will exist only
'            once. To make this restriction productive, the number of
'            occurrences of each line is returned as item.
' ----------------------------------------------------------------------------
    Dim dct     As New Dictionary
    Dim l       As Long
    Dim s       As String
    Dim sSplit  As String
    
    With dct
        For Each v In Split(FileAsString(v_items.Path, sSplit), sSplit)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set FileAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function FileAsFile(ByVal f_file_in As File, _
                            ByVal f_file_out As String, _
                   Optional ByVal f_append As Boolean = False, _
                   Optional ByVal f_rename As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a file (f_file_in) as a file with another full name.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsFile"
    
    On Error GoTo eh
    Dim sSplit As String
    
    With FSo
        Select Case True
            Case f_rename And Not f_append:     If f_file_in.Path = .GetParentFolderName(f_file_out) _
                                                Then f_file_in.Name = .GetFileName(f_file_out) _
                                                Else Err.Raise AppErr(1), ErrSrc(PROC), "File cannot be renamed when the provided file and the new file's name " & _
                                                                                        "do not point to the same location!"
            Case Not f_rename And Not f_append: .CopyFile f_file_in.Path, f_file_out
            Case Not f_rename And f_append:     StringAsFile FileAsString(f_file_in, sSplit), f_file_out, True
            Case Else:                          Err.Raise AppErr(2), ErrSrc(PROC), "Rename  a n d  append is not supported!"
        End Select
    
        Set FileAsFile = .GetFile(f_file_out)
    End With
    
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsString(Optional ByVal f_file_full_name As String, _
                              Optional ByRef f_split As String, _
                              Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    On Error Resume Next
    Dim s As String
    
    Open f_file_full_name For Input As #1
    If Err.Number <> 0 Then GoTo xt
    s = Input$(lOf(1), 1)
    Close #1
    
    Select Case True
        Case InStr(s, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(s, vbCr) <> 0:   f_split = vbCr
        Case InStr(s, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate a trailing eof if any
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    Do While Right(s, Len(f_split)) = f_split
        s = Left(s, Len(s) - Len(f_split))
        If Len(s) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        s = FileAsStringEmptyExcluded(s)
    End If
    FileAsString = s
    
xt:
End Function

Private Function FileAsStringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    
    s_s = FileAsStringTrimmed(s_s) ' leading and trailing empty already excluded
    Do While InStr(s_s, vbCrLf & vbCrLf) <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = s_s
    
End Function

Private Function FileAsStringTrimmed(ByVal s_s As String, _
                            Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns a file as string (s_s) with any leading and trailing empty items,
' i.e. record, lines, excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    
    s = s_s
    '~~ Eliminate any leading empty items
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate a trailing eof if any
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    '~~ Eliminate any trailing empty items
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    FileAsStringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Public Sub FldrUnZip(ByVal f_source As Variant, _
                       ByVal f_target As Variant)
' ----------------------------------------------------------------------------
' Unzips a zip file (f_source) to a target folder (f_target).
' ----------------------------------------------------------------------------
    Const PROC = "FldrUnZip"
    
    On Error GoTo eh
    Dim ShellApp As Object

    'Copy the files & folders from the zip into a folder
    Set ShellApp = CreateObject("Shell.Application")
    With ShellApp
        .Namespace(f_target).CopyHere .Namespace(f_source).Items
    End With
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub FldrZip(ByVal f_source As Variant, _
                     ByVal f_target As Variant)
' ----------------------------------------------------------------------------
' Zips a folder (f_source) into a zip file (f_target).
' ----------------------------------------------------------------------------

    Dim ShellApp As Object

    '~~ Create an empty zip file
    Open f_target For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1

    '~~ Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    With ShellApp
        .Namespace(f_target).CopyHere ShellApp.Namespace(f_source).Items

        '~~ Zipping the files may take a while, create loop to pause the macro until zipping has finished.
        On Error Resume Next
        Do Until .Namespace(f_target).Items.Count = .Namespace(f_source).Items.Count
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
    End With
    On Error GoTo 0

End Sub

Public Sub FolderUnZip(ByVal t_zip_file_name As String, _
              Optional ByRef t_folder As String)
' ----------------------------------------------------------------------------
' Unzips a test folder identified by a test-number (t_no) which is a
' sub-folder in the default test folder. I.e. the name syntax for a test
' folder is "Test_" followed by the provoded number ending with an extension
' ".zip".
' ----------------------------------------------------------------------------
    Const PROC = "FolderUnZip"
    
    On Error GoTo eh
    Dim sSource As String
    Dim sTarget As String
    
    sSource = sTestFolder & "\" & t_zip_file_name
    sTarget = sTestFolder & "\" & Replace(t_zip_file_name, ".zip", vbNullString)
    If FSo.FolderExists(sTarget) Then FSo.DeleteFolder sTarget
    FSo.CreateFolder sTarget
    FldrUnZip sSource, sTarget
    t_folder = sTarget

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub FolderZip(ByVal t_no As String)
' ----------------------------------------------------------------------------
' Zips the test folder with the number (t_no).
' ----------------------------------------------------------------------------
    Const PROC = "FolderZip"
    
    Dim sSource As String
    Dim sTarget As String
    
    sTarget = sTestFolder & "\Test_" & t_no & ".zip"
    sSource = sTestFolder & "\Test_" & t_no
    
    If Not FSo.FolderExists(sSource) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A folder to zip """ & sSource & """ does not exist!"
    
    FldrZip sSource, sTarget

End Sub

Public Sub Instructions(ByVal i_instruction As String)
' ------------------------------------------------------------------------------
' Displays instructions for a required interaction by the "tester". When the
' Common VBA Message Component (`mMsg`/`fMsg`( is installed (indicated by the
' Cond. Comp. Arg. `mMsg = 1`) the mMsg.Dsply service is used to display the
' instructions non-modal, else the instruction is provided in the
' Application.StatusBar.
' ------------------------------------------------------------------------------
    Const INST_TITLE = "Required manual interaction for test "
    
    siInstructionsTitle = INST_TITLE & sTestId

#If mMsg = 1 Then
    Dim Msg     As mMsg.udtMsg
    Dim i       As Long
    
    If Not mMsg.Instances Is Nothing Then
        If mMsg.Instances.Count <> 0 Then
            '~~ Unload any previously dsiplayed instructions
            For i = mMsg.Instances.Count - 1 To 0 Step -1
                If InStr(mMsg.Instances.Keys(i), INST_TITLE) <> 0 Then
                    If i = mMsg.Instances.Count - 1 Then
                        '~~ Obtain the position on screen from the previous/last displayed instructions
                        On Error Resume Next
                        siInstructionsPosTop = mMsg.Instance(mMsg.Instances.Keys(i)).Top
                        siInstructionsPosLeft = mMsg.Instance(mMsg.Instances.Keys(i)).Left
                    End If
                    Unload mMsg.Instance(mMsg.Instances.Keys(i))
                End If
            Next i
        End If
    End If
    Msg.Section(1).Text.Text = i_instruction

    mMsg.Dsply siInstructionsTitle, Msg, "", vbNullString, , , , True, , , , , siInstructionsPosTop & ";" & siInstructionsPosLeft
#Else
    If Application.StatusBar = vbNullString _
    Then Application.StatusBar = i_instruction _
    Else Application.StatusBar = Application.StatusBar & "  (" & i_instruction & ")"
#End If
End Sub

Public Sub InstructionsDone()

#If mMsg = 1 Then
    On Error Resume Next
    siInstructionsPosTop = frmInstructions.Top
    siInstructionsPosLeft = frmInstructions.Left
    Unload mMsg.Instance(siInstructionsTitle)
#Else
    Application.StatusBar = vbNullString
#End If
    
End Sub

Private Function IsObject(ByVal i_var As Variant, _
                          ByRef i_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the object's (i_var) name (i_name) when a variant (i_var)
' is an object. When the object does not have a Name property an error is
' raised.
' ----------------------------------------------------------------------------
    Const PROC = "IsObject"
    
    If Not VBA.IsObject(i_var) Then Exit Function
    IsObject = True
    On Error Resume Next
    i_name = i_var.Name
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "VarTrans tries to use the Name property of an object when it is to be " & _
         "transferred into a string which is the case when String, Array, or File " & _
         "is the target format. However, the current item is an object which does " & _
         "not have a Name property!"
    
End Function

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add Key:=vKey, Item:=k_dct.Item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function LogColsArrangedItems(ParamArray c_items() As Variant) As String
' -----------------------------------------------------------------------------------
' Returns a string with the provided items (c_items) arranged in columns, whereby:
' - column widths provided beforehand by means of LogColsWidth are considered,
'   expanded to the actual required width when the truncate option (c_truncate) is
'   False and truncated to the specified width otherwise. For a column no width has
'   been specified beforehand the column width of the first aligned column item is
'   considered even when the truncate option (c_truncate) is True.
' - a column delimiter may be a single space or a | (vertical bar). In the first
'   case the margin for each individual item becomes a vbNullString, in the second
'   case the margin will default to a single space.
'
' Precondition: Column specs had been provided (LogColsSpecs).
' -----------------------------------------------------------------------------------
    Const PROC = "LogColsArrangedItems"
    
    On Error GoTo eh
    Dim iCol       As Long
    Dim s          As String
    Dim sColDelim  As String
    Dim sColItem   As String
    Dim sColString As String
    Dim sItemAlgnd As String
    Dim v          As Variant
        
    If Not ArrayIsAllocated(c_items) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No items to be aligned in columns had been provided by the caller!"
    
    For Each v In c_items
        iCol = iCol + 1
        sColItem = v
        Select Case LogColAlign(iCol)
            Case "L":   sColString = RTrim$(sColItem)
            Case "C":   sColString = Trim$(sColItem)
            Case "R":   sColString = LTrim$(sColItem)
        End Select
        
        '~~ With the exception of the rightmost columns - which has an unlimited length -
        '~~ each column requires a specified width.
        If LogColWidth(iCol) = 0 _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "Missing width spec for column " & iCol & " !"
        
        sItemAlgnd = AlignString(sColString, LogColAlign(iCol), LogColWidth(iCol), LogColFill(iCol), LogColsMargin, iCol, True)
        s = ConcatenateStrings(s, sItemAlgnd, LogColsDelim, sColDelim)
    Next v
    
xt: LogColsArrangedItems = s
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub LogColsHeaderWrite(ByVal l_file As String, _
                      Optional ByRef l_title As String = vbNullString, _
                      Optional ByVal l_with_header_underline As Boolean = False, _
                      Optional ByVal l_width As Long)
' ------------------------------------------------------------------------------
' Write a title (l_title) when one is still specified and write column headers
' when one are still specified - both are reset once written to a file (l_file).
'
' Preconditions:
' - Columns had been specified (LogColsSpecs) whereby the column width is
'   mandatory while the alignment defaults to Left when ommited
' - A column delimiter has been specified (LogColsDelim)
' - One or more header lines had been specified (LogColsHeader).
' ------------------------------------------------------------------------------
    Const PROC = "LogColsHeaderWrite"
    
    On Error GoTo eh
    Dim lDelim      As Long
    Dim sColDelimUl As String
    Dim sHeaderUl   As String
    Dim v           As Variant
    
    If l_title <> vbNullString Then
        StringAsFile AlignCntr(l_title, l_width, " ="), l_file, True
        l_title = vbNullString
    End If
        
    If ArrayIsAllocated(aLogColsArrangedHeaders) Then
        If l_with_header_underline Then
            '~~ Prepare cols header top/bottom line
            lDelim = Len(LogColsDelim) + Len(LogColsMargin)
            For Each v In aLogColWidths
                sHeaderUl = ConcatenateStrings(sHeaderUl, String(v + lDelim, "-"), "+", sColDelimUl)
            Next v
            StringAsFile sHeaderUl, l_file, True
        End If
        For Each v In aLogColsArrangedHeaders
            StringAsFile CStr(v), l_file, True
        Next v
        If l_with_header_underline _
        Then StringAsFile sHeaderUl, l_file, True
        ArrayErase aLogColsArrangedHeaders
    End If
        
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogColSpec(ByVal c_col As Long, _
                       ByVal c_strng As String)
' ----------------------------------------------------------------------------
' Returns a column's (c_col) specification of the width (c_width), alignment
' (sAlign), and fill (sFill).
' ----------------------------------------------------------------------------
    Const PROC = "ColSpec"
    
    On Error GoTo eh
    Dim s       As String
    Dim sAlign  As String
    Dim sFill   As String
    
    s = c_strng
    '~~ Get alignment (defaults to Left when none is specified
    Select Case True
        Case LogColSpecTry(s, "Left"):      sAlign = "L"
        Case LogColSpecTry(s, "Center"):    sAlign = "C"
        Case LogColSpecTry(s, "Right"):     sAlign = "R"
        Case LogColSpecTry(s, "Lft"):       sAlign = "L"
        Case LogColSpecTry(s, "Cntr"):      sAlign = "C"
        Case LogColSpecTry(s, "Rght"):      sAlign = "R"
        Case LogColSpecTry(s, "L"):         sAlign = "L"
        Case LogColSpecTry(s, "C"):         sAlign = "C"
        Case LogColSpecTry(s, "R"):         sAlign = "R"
        Case Else:                          sAlign = "L" ' the default
    End Select
    LogColAlign(c_col) = sAlign
    
    '~~ Get fill (defaults to a single space when none is specified
    Select Case True
        Case LogColSpecTry(s, " .:", sFill)
        Case LogColSpecTry(s, ".:", sFill)
        Case LogColSpecTry(s, " -", sFill):    LogColSpecTry s, "- ", sFill   ' if alignment is Centered
        Case LogColSpecTry(s, " =", sFill):    LogColSpecTry s, "= ", sFill   ' if alignment is Centered
        Case LogColSpecTry(s, ".", sFill)
        Case LogColSpecTry(s, "-", sFill):     LogColSpecTry s, "-", sFill    ' if alignment is Centered
        Case LogColSpecTry(s, " ", sFill)
        Case Else:                          sFill = " "  ' the default
    End Select
    LogColFill(c_col) = sFill
    
    '~~ Get column width
    '~~ When alignment and fill specs had been extracted (and eliminated), the
    '~~ remaining (or the only spec) is to be the column width.
    If Not IsNumeric(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "The column specification does not orovide a numeric value indicating the columns width." & "||" & _
         "The column width is mandatoryIt may be the only specification in case the fill is a single " & _
         "space (the default) and the alignment is Left (the default)."
    
    LogColWidth(c_col) = CLng(s)
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function LogColSpecTry(ByRef c_strng As String, _
                               ByVal c_try As String, _
                      Optional ByRef c_return As String) As Boolean
                   
    If InStr(c_strng, c_try) <> 0 Then
        c_return = c_try
        LogColSpecTry = True
        c_strng = Replace(c_strng, c_try, vbNullString)
    End If

End Function

Private Function Max(ParamArray va() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Private Function ObjectAsString(ByVal o As Object) As String
    
    On Error Resume Next
    ObjectAsString = o.Name
    If Err.Number <> 0 Then ObjectAsString = "Object_" & o.HashCode

End Function

Private Sub ResultEvaluated(Optional ByVal r_result As Collection = Nothing, _
                            Optional ByVal r_result_expected As Collection = Nothing, _
                            Optional ByVal r_test_id As String = vbNullString)
' ------------------------------------------------------------------------------
' When test collections (r_result and r_result_expected are provided these
' collections are compared else the Result and the ResultExpected identified
' by the test-number in the dctResult and dctResultExpected Dictionaries.
' The first choice allows basic without using any function tested elsewhere.
' ------------------------------------------------------------------------------
    Const PROC = "ResultEvaluated"
    
    On Error GoTo eh
    Dim i               As Long
    Dim sResult         As String
    Dim sResultExpected As String
    
    sResultEvaluated = RESULT_PASSED ' the default
    If r_result Is Nothing And r_result_expected Is Nothing Then
        sResult = Exclude(VarAsString(dctResult(r_test_id)))
        sResultExpected = Exclude(VarAsString(dctResultExpected(r_test_id)))
        If StrComp(sResult _
                 , sResultExpected _
                 , vbBinaryCompare) <> 0 _
        Then sResultEvaluated = RESULT_FAILED
    Else
        If r_result.Count <> r_result_expected.Count Then
            sResultEvaluated = RESULT_FAILED
            GoTo xt
        Else
            For i = 1 To r_result.Count
                If StrComp(CStr(r_result(i)), CStr(r_result_expected(i)), vbBinaryCompare) <> 0 Then
                    sResultEvaluated = RESULT_FAILED
                    GoTo xt
                End If
            Next i
        End If
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ResultEvaluation(Optional ByVal r_result As Collection = Nothing, _
                             Optional ByVal r_result_expected As Collection = Nothing)
' ----------------------------------------------------------------------------
' When a result Collection (r_result) and a result expected Collection
' (r_result_expected) is provided the two peovided Collections are compared,
' else the Result and the ResultExpected identified by the test-number in the
' dctResult and dctResultExpected Dictionaries.
' In regression mode the result is provided as summary only, else a detailed
' result is only displayed when the test has failed.
' ----------------------------------------------------------------------------
    Const PROC = "ResultEvaluation"
    
    On Error GoTo eh
    Dim cll As Collection
        
    ResultEvaluated r_result, r_result_expected, VerificationId
    If bModeRegression Then
        '~~ In regression mode all results are collected and presented as summary
        If Not r_result Is Nothing And Not r_result_expected Is Nothing Then
            '~~ The Result and the ResultExpected has been provided explicitly.
            '~~ For a correct test-id, both are saved to the dctResult and dctResultExpected
            If dctResult.Exists(VerificationId) Then dctResult.Remove VerificationId
            Dict(dctResult, VerificationId) = r_result
            If dctResultExpected.Exists(VerificationId) Then dctResultExpected.Remove VerificationId
            Dict(dctResultExpected, VerificationId) = r_result_expected
        End If
        '~~ In regression mode the result is saved for being displayed as summary.
        Set cll = New Collection
        With cll
            .Add sTitle
            .Add VerificationId
            .Add sTestedProcType
            .Add sTestedComp
            .Add sTestedProc
            .Add sVerification
            .Add sResultEvaluated
            If cExecTimeMsecs = 0 _
            Then .Add "n/a" _
            Else .Add TimerExecTimeFormatted(cExecTimeMsecs)
            cExecTimeMsecs = 0
        End With
        Dict(dctTestSummary, VerificationId) = cll
'        Debug.Print ErrSrc(PROC) & ": " & r_test_id & " added to dctTestSummary"
        Set cll = Nothing
    Else
        '~~ In non-regression mode a failed test result is presented immediately
        If sResultEvaluated = RESULT_FAILED Then ResultFailedLog r_result, r_result_expected, VerificationId
    End If
    
xt: sVerification = vbNullString ' Another pair of .Result/.ResultExpected requires a new specified Verification.
    sTitle = vbNullString
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ResultFailedLog(Optional ByVal r_result As Collection = Nothing, _
                            Optional ByVal r_result_expected As Collection = Nothing, _
                            Optional ByVal r_test_id As String = vbNullString)
' ------------------------------------------------------------------------------
' Displays the result of a test versus its expected result. The result is
' presented in three columns: Passed/Failed Result Result Expected.
' Called by ResultEvaluation.
' ------------------------------------------------------------------------------
    Const PROC = "ResultFailedLog"
    
    On Error GoTo eh
    Dim cllResult           As New Collection
    Dim cllResultExpected   As New Collection
    Dim i                   As Long
    Dim lMaxResult          As Long
    Dim lMaxResultExpected  As Long
    Dim v                   As Variant
    Dim sLine               As String
    Dim sPassedFailed       As String
    
    If Not r_result Is Nothing And Not r_result_expected Is Nothing Then
        Set cllResult = r_result
        Set cllResultExpected = r_result_expected
    Else
        Set cllResult = VarAsCollection(dctResult(r_test_id))
        Set cllResultExpected = VarAsCollection(dctResultExpected(r_test_id))
    End If
    
    For Each v In cllResult
        lMaxResult = Max(lMaxResult, Len(v) + 2)
    Next v
    For Each v In cllResultExpected
        lMaxResultExpected = Max(lMaxResultExpected, Len(v) + 2)
    Next v

    sLogFile = NameTestResultFile
    If FSo.FileExists(sLogFile) Then FSo.DeleteFile sLogFile
    
    LogColsSpecs = "L" & Len(sVerification) & ", C8, L" & lMaxResult & ", L" & lMaxResultExpected
    LogColsDelim = "|"
    LogColsHeader = "Verification, P/F, Result, Result Expected"
                  
    For i = 1 To Min(cllResult.Count, cllResultExpected.Count)
        If StrComp(CStr(cllResult(i)), CStr(cllResultExpected(i)), vbBinaryCompare) = 0 _
        Then sPassedFailed = "Passed" _
        Else sPassedFailed = "Failed"
        sLine = LogColsArrangedItems(sVerification, sPassedFailed, """" & Trim(cllResult(i)) & """", """" & Trim(cllResultExpected(i)) & """")
        LogColsHeaderWrite l_file:=sLogFile _
                         , l_title:=sTitle _
                         , l_with_header_underline:=True _
                         , l_width:=Len(sLine)
        StringAsFile sLine, sLogFile, True
    Next i
    If cllResult.Count > cllResultExpected.Count Then
        For i = i To cllResult.Count
            StringAsFile LogColsArrangedItems(sVerification, "Failed", """" & Trim(CStr(cllResult(i))) & """", """" & vbNullString & """"), sLogFile, True
        Next i
    ElseIf cllResultExpected.Count > cllResult.Count Then
        For i = i To cllResultExpected.Count
            StringAsFile LogColsArrangedItems(sVerification, "Failed", """" & vbNullString & """", """" & Trim(CStr(cllResultExpected(i)))) & """", sLogFile, True
        Next i
    End If
    ShellRun sLogFile ' display log

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ResultSummaryLog()
' ----------------------------------------------------------------------------
' Dsiplays the test result summary.
' ----------------------------------------------------------------------------
    Const PROC = "ResultSummaryLog"
    
    On Error GoTo eh
    Dim bMore   As Boolean
    Dim cll     As Collection
    Dim lComp   As Long
    Dim lDesc   As Long
    Dim lTtId   As Long
    Dim lProc   As Long
    Dim lRslt   As Long
    Dim lTtle   As Long
    Dim lType   As Long
    Dim lMsec   As Long
    Dim sComp   As String
    Dim sSpec   As String
    Dim sTtId   As String
    Dim sLine   As String
    Dim sMsec   As String
    Dim sProc   As String
    Dim sRslt   As String
    Dim sTtle   As String
    Dim sType   As String
    Dim v       As Variant
    
    If Not bModeRegression Then Exit Sub ' is bypassed except ModeRegression = True
    If FSo.FileExists(sTestResultSummaryLogFile) Then FSo.DeleteFile sTestResultSummaryLogFile
    
    '~~ Get max lengths
    For Each v In dctTestSummary
        Set cll = dctTestSummary(v)
        If sComp <> vbNullString And sComp <> cll(3) Then
            bMore = True
        ElseIf sComp = vbNullString Then sComp = cll(3)
        End If
        lTtle = Max(lTtle, Len(cll(1)))
        lTtId = Max(lTtId, Len(cll(2)))
        lType = Max(lType, Len(cll(3)))
        lComp = Max(lComp, Len(cll(4)))
        lProc = Max(lProc, Len(cll(5)))
        lDesc = Max(lDesc, Len(cll(6)))
        lRslt = Max(lRslt, Len(cll(7)))
        lMsec = Max(lMsec, Len(cll(8)))
    Next v
        
    '~~ Column specs and headers
    LogColsSpecs = lTtle & ", " & lTtId & ", " & lComp & ", " & lType & ", " & lProc & ", " & lDesc & ", C" & lRslt & ", R" & lMsec
    LogColsDelim = "|"
    '~~ Specifiy headers (are all alligned centered)
    LogColsHeader = "Test,  Test, Component, Component, Procedure, Verification, Test, Exec time"
    LogColsHeader = "Title, Id  , Name,      Type,      Name, (expected/tested result), Result, (msec)"
    sTtle = "Regression Test Summary"
    
    For Each v In dctTestSummary
        Set cll = dctTestSummary(v)
        If sTtle <> cll(1) Then sTtle = cll(1) Else sTtle = vbNullString
        sTtId = cll(2)
        sType = cll(3)
        sComp = cll(4)
        sProc = cll(5)
        sSpec = cll(6)
        sRslt = cll(7)
        sMsec = cll(8)
        sLine = LogColsArrangedItems(sTtle, sTtId, sComp, sType, sProc, sSpec, sRslt, sMsec)
        LogColsHeaderWrite l_file:=sTestResultSummaryLogFile _
                         , l_title:=sTitleRegression _
                         , l_with_header_underline:=True _
                         , l_width:=Len(sLine)

        StringAsFile sLine, sTestResultSummaryLogFile, True
    Next v
    ShellRun sTestResultSummaryLogFile
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

Private Function StringAsCollection(ByVal s_items As String, _
                           Optional ByVal s_split As String = vbNullString) As Collection
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    With cll
        For Each v In Split(s_items, s_split)
            .Add v
        Next v
    End With
    Set StringAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function StringAsDictionary(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    
    With dct
        For Each v In Split(s_items, s_split)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                l = dct(v) + 1
                .Remove v
                .Add v, l
            End If
        Next v
    End With
    Set StringAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function StringAsFile(ByVal s_strng As String, _
                     Optional ByRef s_file As Variant = vbNullString, _
                     Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned.
' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
'         is written a records/lines.
' Note 2: When the string has the alternate split indicator "|&|" this one is
'         replaced by vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFileFullName
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = FSo.GetFile(s_file)
    
End Function

Private Function StringAsString(ByVal s_items As String, _
                       Optional ByVal s_delimiter As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns a string (v_item) with any delimiter delimited by vbCrLf by default.
' ----------------------------------------------------------------------------
    Const PROC = "StringAsString"
    
    On Error GoTo eh
    Dim sSplit As String
    sSplit = SplitIndctr(s_items)

    StringAsString = Replace(s_items, sSplit, s_delimiter)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function TempFileFullName(Optional ByVal f_path As String = vbNullString, _
                                 Optional ByVal f_ext As String = "txt", _
                                 Optional ByVal f_create_as_textstream As Boolean = False) As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' The returned temporary file is registered in cllTempTestItems for being removed
' either explicitly with CleanUp or implicitly when the class
' terminates.
' ------------------------------------------------------------------------------
    Dim s As String
        
    If VBA.Left$(f_ext, 1) <> "." Then f_ext = "." & f_ext
    
    s = Replace(FSo.GetTempName, ".tmp", f_ext)
    If f_path = vbNullString Then f_path = CurDir
    s = VBA.Replace(f_path & "\" & s, "\\", "\")
    If f_create_as_textstream Then FSo.CreateTextFile s
    TempFileFullName = s
    TempTestItem = s ' keep a record of it for being removed finally

End Function

Public Function TestFile(Optional ByVal t_name As String = vbNullString, _
                         Optional ByVal t_ext As String = vbNullString) As File
' ------------------------------------------------------------------------------
' Returns the file object with the TestFileFullName, raises an erro when not
' existing.
' ------------------------------------------------------------------------------
    Const PROC = "TestFile"
    
    If FSo.FileExists(TestFileFullName(t_name, t_ext)) _
    Then Set TestFile = FSo.GetFile(TestFileFullName) _
    Else Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & TestFileFullName & """ does not exist!"

End Function

Public Function TestFileFullName(Optional ByVal t_name As String = vbNullString, _
                                 Optional ByVal t_ext As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a TestId and Verification(number) specific default test file name
' with an optional name (t_name) and an optional extension (t_ext).
' When the name is omitted it defaults to a vbNullString.
' When the extension is ommited it defaults to the property TestFileExtension
' which defaults to "txt".
' ----------------------------------------------------------------------------
    Dim sName As String
    Dim sExt  As String
    Dim s     As String
    
    If t_name <> vbNullString Then sName = "_" & t_name
    If t_ext = vbNullString Then sExt = TestFileExtension Else sExt = t_ext
    
    s = sTestFolder & "\Test-" & TestId & sName & "-" & lVerificationNumber & "." & sExt
    s = Replace(s, "..", ".")
    s = Replace(s, "__", "_")
    s = Replace(s, "-_", "_")
    TestFileFullName = s

End Function

Public Function TestResultExpectedFile() As File
' ----------------------------------------------------------------------------
' Returns a file named <test-folder>/.
' ----------------------------------------------------------------------------
    Const PROC = "TestResultExpectedFile"
    
    With FSo
        If .FileExists(sNameTestResultExpectedFile) _
        Then Set TestResultExpectedFile = .GetFile(sNameTestResultExpectedFile) _
        Else Err.Raise AppErr(1), ErrSrc(PROC), "A test result file named """ & sNameTestResultExpectedFile & """ does not exist!"
    End With
End Function

Public Function TestResultFile() As File
' ------------------------------------------------------------------------------
' Returns the default name of a test result file in the form:
' <tes-folder>\TestResult-<test-number><file-extention>
' when a file with the returned name already exists, it is deleted
' ------------------------------------------------------------------------------
    Const PROC = "TestResultFile"
    
    With FSo
        If .FileExists(sNameTestResultFile) _
        Then Set TestResultFile = .GetFile(sNameTestResultFile) _
        Else Err.Raise AppErr(1), ErrSrc(PROC), "A test result file named """ & sNameTestResultFile & """ does not exist!"
    End With
    
End Function

Public Function TimerElapsedTicks() As Double
' ----------------------------------------------------------------------------
' Returns the elapsed number of ticks since start.
' ----------------------------------------------------------------------------
    QueryPerformanceCounter cTicksCurrent
    TimerElapsedTicks = (cTicksCurrent - cTicksStart) / cFrequency
End Function

Public Function TimerEnd(Optional t_exec_time As Currency) As Currency
' ----------------------------------------------------------------------------
' Returns the elapsed time (t_elpsd) since start in milliseconds whereby the
' timer overhead is deducted in order to rturn a result of 0 milliseconds
' when nothing is measured between the TimerStart and TimerElapsedMsecs
' service. The average timer overhead is calculated when the class is
' initialized.
' ----------------------------------------------------------------------------
    t_exec_time = (TimerElapsedTicks() * 1000) - cTimerOverhead
    If t_exec_time < 0 Then t_exec_time = 0.0001
    cExecTimeMsecs = t_exec_time
    TimerEnd = t_exec_time
End Function

Private Function TimerExecTimeFormatted(ByVal t_exec_time As Currency) As String
' ------------------------------------------------------------------------------
' Returns the Currency (t_exec_time) formatted with a length of 11.
' ------------------------------------------------------------------------------
    Dim s As String
    
    s = Split(FormatCurrency(t_exec_time, 4, vbTrue, vbUseDefault, vbFalse), " ")(0)
    s = String(11 - Len(s), " ") & s
    TimerExecTimeFormatted = s
    
End Function

Public Sub TimerStart():    QueryPerformanceCounter cTicksStart:    End Sub

Private Function TruncateIfApplicable(ByRef t_string As String, _
                                      ByVal t_length As Long, _
                                      ByVal t_col As Long, _
                                      ByVal t_cols As Long, _
                                      ByVal t_align As String) As String
' -----------------------------------------------------------------------------------
' Unless the column (t_col) is not the rightmost (t_cols) the string (t_string) is
' truncated when it exceeds the provided length (t_length).
' -----------------------------------------------------------------------------------
    Dim s As String
    
    TruncateIfApplicable = t_string ' in case truncation is not applicable
    If t_col <> 0 And t_cols <> 0 And t_col = t_cols Then
        '~~ Adjust the - optional - width for the rightmost column to enable
        '~~ an apropriate header design in case columns headers are specified.
        If LogColWidth(t_col) = 0 Then LogColWidth(t_col) = Len(t_string)
    Else
        Select Case t_align
            Case "L":   s = RTrim$(t_string)
            Case "C":   s = Trim$(t_string)
            Case "R":   s = LTrim$(t_string)
        End Select
        If Len(t_string) > t_length And t_length > 0 _
        Then TruncateIfApplicable = Left(s, t_length)
    End If
    
End Function

Private Function VarAsArray(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsArray"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsArray = CollectionAsArray(v_items)
        Case TypeName(v_items) = "File":        VarAsArray = FileAsArray(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsArray = DictionaryAsArray(v_items)
        Case TypeName(v_items) Like "*()":      VarAsArray = v_items
        Case VarType(v_items) = vbArray:        VarAsArray = v_items
        Case IsArray(v_items):                  VarAsArray = v_items
        Case VarType(v_items) = vbString:       VarAsArray = StringAsArray(v_items)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Private Function VarAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsCollection = v_items
        Case TypeName(v_items) = "Dictionary":  Set VarAsCollection = DictionaryAsCollection(v_items)
        Case TypeName(v_items) = "File":        Set VarAsCollection = FileAsCollection(v_items)
        Case TypeName(v_items) Like "*()":      Set VarAsCollection = ArrayAsCollection(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsCollection = ArrayAsCollection(v_items)
        Case VarType(v_items) = vbBoolean:      cll.Add BooleanAsString(v_items)
                                                Set VarAsCollection = cll
        Case VarType(v_items) = vbString:       Set VarAsCollection = StringAsCollection(v_items)
        Case Else:                              cll.Add v_items
                                                Set VarAsCollection = cll

    End Select
    
End Function

Private Function VarAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsDictionary = CollectionAsDictionary(v_items)
        Case TypeName(v_items) = "File":        Set VarAsDictionary = FileAsDictionary(v_items)
        Case TypeName(v_items) = "Dictionary":  Set VarAsDictionary = DictionaryAsDictionary(v_items) ' sort key ascending
        Case TypeName(v_items) Like "*()":      Set VarAsDictionary = ArrayAsDictionary(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsDictionary = ArrayAsDictionary(v_items)
        Case VarType(v_items) = vbString:       Set VarAsDictionary = StringAsDictionary(v_items)
        Case VarType(v_items) = vbBoolean:      dct.Add BooleanAsString(v_items), vbNullString
                                                Set VarAsDictionary = dct
    End Select
    
End Function

Private Function VarAsFile(ByVal v_items As Variant, _
                  Optional ByVal v_file As String, _
                  Optional ByVal v_append As Boolean = True) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsFile"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsFile = CollectionAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) = "File"
        Case TypeName(v_items) = "Dictionary":  Set VarAsFile = DictionaryAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) Like "*()":      Set VarAsFile = ArrayAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbArray:        Set VarAsFile = ArrayAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbString:       Set VarAsFile = StringAsFile(v_items, v_file, v_append)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Private Function VarAsString(ByVal v_items As Variant) As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------

    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsString = CollectionAsString(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsString = DictionaryAsString(v_items)
        Case TypeName(v_items) = "File":        VarAsString = FileAsString(v_items)
        Case TypeName(v_items) = "Nothing":     VarAsString = vbNullString
        Case TypeName(v_items) Like "*()":      VarAsString = ArrayAsString(v_items)
        Case VarType(v_items) = vbArray:        VarAsString = ArrayAsString(v_items)
        Case VarType(v_items) = vbBoolean:      VarAsString = BooleanAsString(v_items)
        Case VarType(v_items) = vbString:       VarAsString = StringAsString(v_items)
        Case Else:                              VarAsString = CStr(v_items)
    End Select
    
End Function

Private Function VerificationId() As String
    VerificationId = sTestId & "." & lVerificationNumber
End Function

Private Function WbkExists(ByVal ex_wbk As Variant) As Boolean
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
    Const PROC = "WbkExists"
    
    On Error GoTo eh
    
    WbkExists = FSo.FileExists(ex_wbk)
    If Not WbkExists Then
        On Error Resume Next
        WbkExists = Application.Workbooks(ex_wbk) = Application.Workbooks(ex_wbk).Name
    End If
                
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsOpen(ByVal w_wbk As String, _
                  Optional ByRef w_open_wbk As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a Workbook (w_wbk) is open.
' ----------------------------------------------------------------------------
    Const PROC  As String = "WbkIsOpen"
    
    On Error Resume Next
    Set w_open_wbk = Application.Workbooks(w_wbk)
    WbkIsOpen = Err.Number = 0
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

