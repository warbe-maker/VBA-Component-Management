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
' a log file for performed tests and their Result (Passed or Failed).
' - The class module by default initializes a "Test" folder in the
'   application's parent folder.
' - Each individual test is identified by a unique test-number
' - The Result (of the tested procedure) and the expected Result may be
'   provided in any form (Array, Collection, Dictionary, File, String, or
'   Boolean). When both had been provided for a test they are compared which
'   results in a failed or passed status.
' - Result/ResultExpected comparisons are numbered which means that an
'   individual test may have several Result verifications
' - A regression-mode allows to specify "asserted errors" beforehand in order
'   to not interrupt a test sequence. Furthermore, all results are collected
'   and may finally be displayed (ResultLogSummary). In contrast, i.e. in
'   non-regression-mode - which is the default) any failed test is immediately
'   displayed (VrfctnLogsWrite)
'
' Public services:
' ----------------
' AssertedErrors              Error numbers regarded asserted for a specific
'                             test. Effective only when the Standard Module
'                             mErH is installed and indicated by the Cond.
'                             Comp. Argument `mErH = 1`.
' Comment               P   w Comment on Test, for the whole id or Verification.
'                             (may be overwritten by a "FailedNote" in case a
'                             verification failed)
' ExcludeFromComparison P   w Specifies a like string which is excluded from
'                             the evaluation/comparision of a Result with an
'                             expected Result.
' FolderUnZip                 Unzips a zipped folder to a target folder.
' FolderZip                   Zips a folder into a zip file.
' Guide                       Non modal message popup providing testing
'                             guidance.
' Instructions          M     Displays instructions for a required
'                             interaction by the "tester". When the
'                             _Common VBA Message Component_ `mMsg`/`fMsg`
'                             is installed (indicated by the Cond. Comp.
'                             Arg. `mMsg = 1`) the mMsg.Dsply service is
'                             used to display the instructions non-modal,
'                             else the instructions are provided in the
'                             Application.StatusBar.
' InstructionsDone      M     Unload the displayed instructions if any.
'                             Note: The form window is identified via the
'                             displayed title which is the key to the
'                             displayed instance of the `fMsg` UserForm.
' InstructionsPosLeft   P r/w Popup message position
' InstructionsPosTop    P r/w Popup message position
' ResultLogSummary      M     Writes a verifications summary log to the
'                             log Worksheet.
' Result                P r/w Variant expression registers/returns the Result
'                             of a test, may be a file object, any other
'                             object or any other data type.
' Result                P r   Returns Passed or Failed
' ResultExpected        P r/w The expected Result of a test, a file object,
'                             any other object or any other data type
' CleanUp               M     Removes all files provided either by the
'                             TempTestItem or the TempFileFullName service.
' TestedComp            P   w Name of the tested VBComponent in a test
'                             procedure.
' TestedProc            P   w The name of the tested procedure (a Sub,
'                             Function, Property, or Method).
' TestedProcType        P   w Tested procedure type (Sub, Function, ...)
' TestFolder            P r/w String expression, specifies the test folder
' TestFile              M     Returns the file object with the
'                             TestFileFullName, raises an erro when not
'                             existing.
' TestFileFullName            Returns a test file's full name based on the
'                             current `TestId` and the specifies name scheme
'                             which defaults to "Test-<id>-<vn>-<nm>.<xt>".
' TestFileNameScheme    P   w Specidies the scheme a test file's name is
'                             assembled by default
' TestId                P   w Obligatory, distinct identifcation of a test,
'                             provided as the very first Property for one
'                             or more verifications (ResultExpected/Result
'                             pairs).
' TempFileFullName      M
' TempTestItem          P r/w Saves the name of a temporary test file in a
'                             Collection for being removed when the class
'                             terminates or explicitly by the CleanUp service.
' TimerStart                  Initiates a maximum precision timer
' TimerElapsedTicks     M     Returns the elapsed ticks since start
' TimerEnd                    Returns the elapsed time since start in
'                             milliseconds whereby the timer overhead is
'                             deducted in order to return a Result of 0
'                             milliseconds when nothing is measured between
'                             TimerStart and TimerEnd.
' TimerOverhead         P r   Returns the milliseconds the TimerStart/TimerEnd
'                             itself consumes when no taking any actual
'                             measurement.
' Title                 P   w Title of a test procedure, valid for all tests
'                             and verifications in it.
' Verification          P   w A string describing what specifically is verified
'                             by the tests ResultExpected/Result.
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
' Requires: Reference to Microsoft Scripting Runtime

' W. Rauschenberger, Berlin Feb 2025
' ----------------------------------------------------------------------------
Public fso                  As New FileSystemObject

Private Const VR_PASSED                 As String = " Passed "          ' Verification Result
Private Const VR_FAILED                 As String = " F a i l e d ! "   ' Verification Result
Private Const SCHEME_ID                 As String = "<id>"
Private Const SCHEME_NAME               As String = "<nm>"
Private Const SCHEME_VER_NO             As String = "<vn>"
Private Const SCHEME_EXT                As String = "<xt>"
'~~ Array indices for the collection and retrieval of Result details and Result summary log
Private Const IX_LOGITEM_COMMENT        As Long = 1
Private Const IX_LOGITEM_COMP_NAME      As Long = 2
Private Const IX_LOGITEM_EXEC_TIME      As Long = 3
Private Const IX_LOGITEM_PROC_NAME      As Long = 4
Private Const IX_LOGITEM_PROC_TYPE      As Long = 5
Private Const IX_LOGITEM_RES            As Long = 6
Private Const IX_LOGITEM_RES_EXP        As Long = 7
Private Const IX_LOGITEM_TEST_ID        As Long = 8
Private Const IX_LOGITEM_TITLE          As Long = 9
Private Const IX_LOGITEM_VRFCTN         As Long = 10
Private Const IX_LOGITEM_VRFCTN_NO      As Long = 11
Private Const IX_LOGITEM_VRFCTN_RSLT    As Long = 12
Private Const IX_LOGITEM_VR_TYPE        As Long = 13

#If VBA7 Then
    Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
    Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As Any, xlApp As Object) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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
    Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As Any, xlApp As Object) As Long
    Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As Any, xlApp As Object) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const WIN_NORMAL = 1         'Open Normal
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

Private aLogHeader                  As Variant
Private aLogColFrmt                 As Variant
Private aLogDetails                 As Variant
Private aLogSummary                 As Variant
Private appLog                      As Application
Private wbkLog                      As Workbook
Private wshLog                      As Worksheet
Private bResult                     As Boolean
Private bResultExpected             As Boolean
Private cllLogDetails               As Collection
Private cllTempTestItems            As Collection
Private cllTimerStack               As New Collection
Private cllTicksStackOvrhd          As New Collection
Private cllVrfctnSummary            As Collection
Private cTicksElapsed               As Currency
Private cTicksEnd                   As Currency
Private cTicksOverhead              As Currency
Private cTicksStart                 As Currency
Private cTimerFrequency             As Currency
Private dTimerMsecs                 As Double
Private dTimerOverhead              As Double
Private lCols                       As Long
Private lMaxVrfctn                  As Long
Private lVrfctnNumber               As Long
Private rngLogSheetHeader           As Range
Private sComment                    As String
Private sFailedNote                 As String
Private sTestResultSummaryLogFile   As String
Private sNameTestFile               As String
Private sNameTestResultExpectedFile As String
Private sNameTestResultFile         As String
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
Private sVrfctnSummaryComment       As String
Private sVrfctnType                 As String
Private sTestFileNameScheme         As String
Private v                           As Variant
Private vResult                     As Variant
Private vResultExpected             As Variant
Private vTempTestItem               As Variant
Private siInstructionsPosLeft       As Single
Private siInstructionsPosTop        As Single
Private siInstructionsTitle         As String
Private sLogBookFullName            As String
Private sLogBookName                As String

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
    
    Set cllTempTestItems = New Collection
    sTestFileNameScheme = "Test-<id>-<vn>-<nm>.<xt>"                     ' default tes file name scheme
    TestFolder = ThisWorkbook.Path & "\Test"                             ' default test folder
    TestResultSummaryLogFile = sTestFolder & "\Test-Result-Summary.log"  ' default regression test summary log file's full name
    
    ' ------------------------------------------------------------------------
    ' Provide the milliseconds the timer itself consumes when it not does any
    ' actual measurement, computed when the class is initiated and deducted
    ' from any measurement at TimerEnd. This is the way the timer service
    ' provided (TimerStart, TimerEnd) will return a figure with a minimum
    ' unavoidable precision failure which is a fraction of a milliseconds.
    ' ------------------------------------------------------------------------
    sTestFileExtension = "txt"
    dTimerMsecs = 0
    
    siInstructionsPosTop = 10
    siInstructionsPosLeft = 10
    lMaxVrfctn = Max(Len(VR_PASSED), Len(VR_FAILED))
    
    sLogBookName = "VerificationsLog.xlsx"
    sLogBookFullName = sTestFolder & "\" & sLogBookName

    '~~ Initialize Log headers                                      Initialize column formats (indentation, alignment, merge equal content)
          LogItemTestId(aLogHeader) = "Test" & vbLf & "Id":               LogItemTestId(aLogColFrmt) = Array(0, xlCenter)
        LogItemCompName(aLogHeader) = "Tested" & vbLf & "Component":    LogItemCompName(aLogColFrmt) = Array(1, xlLeft)
           LogItemTitle(aLogHeader) = "Test Title/Description":            LogItemTitle(aLogColFrmt) = Array(1, xlLeft, True)
        LogItemProcName(aLogHeader) = "Tested" & vbLf & "Procedure":    LogItemProcName(aLogColFrmt) = Array(1, xlLeft)
        LogItemProcType(aLogHeader) = "Procedure" & vbLf & "Type":      LogItemProcType(aLogColFrmt) = Array(1, xlLeft)
        LogItemVrfctnNo(aLogHeader) = "Verif." & vbLf & "No":           LogItemVrfctnNo(aLogColFrmt) = Array(0, xlCenter)
    LogItemVrfctnResExp(aLogHeader) = "Result Expected":            LogItemVrfctnResExp(aLogColFrmt) = Array(0, xlCenter)
       LogItemVrfctnRes(aLogHeader) = "Result":                        LogItemVrfctnRes(aLogColFrmt) = Array(0, xlCenter)
      LogItemVrfctnDesc(aLogHeader) = "Verification Description":     LogItemVrfctnDesc(aLogColFrmt) = Array(1, xlLeft, True)
       LogItemVrfydType(aLogHeader) = "Verified" & vbLf & "Type":      LogItemVrfydType(aLogColFrmt) = Array(1, xlLeft)
         LogItemVxcTime(aLogHeader) = "Exec-Time" & vbLf & "(msec)":     LogItemVxcTime(aLogColFrmt) = Array(1, xlRight)
         LogItemComment(aLogHeader) = "Comment":                         LogItemComment(aLogColFrmt) = Array(1, xlLeft, True)

End Sub

Private Sub Class_Terminate()

    Dim v   As Variant
    
    Set cllVrfctnSummary = Nothing
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

    LogTerminate

End Sub

Private Property Let Arry(Optional ByRef a_arr As Variant, _
                          Optional ByVal a_indices As Variant = Empty, _
                                   ByVal a_var As Variant)
' ----------------------------------------------------------------------------
' Common WRITE to an array service. The service returns an array (a_arr) with
' the provided item (a_var) either:
' - Simply added, when no index/indices (a_indices) are provided
' - Having added or replaced an item at a given index/indices (a_indices) by
'   concidering that the returned array has new from/to specifics for any of
'   its dimensions at any level whereby the from specific for any dimension
'   remains the same
' - Having created a 1 to 8 dimensions array regarding the provided indices
'   (a_indices) with the provided item added or replaced
' - Having re-dimensioned the provided/input array to the provided indices
'   (a_indices) with the provided item (a_var) added or updated
'
' The index/indices (a_indices) may be provided as:
' - a single integer, indicating that the provided is or the returned will be
'   a 1-dimensional array
' - a string of indices delimited by a comma, indicating that the provided or
'   the returned array is multi-dimensional
' - an array or Collection of indices (for a multi-dimensional array)
'
' Note: In contrast to VBA's ReDim statement this service is able to
'       extend any "to" specification of any dimension (not only the last
'       one) with the only constraint that the "from" specification of any
'       dimension will remain the same.
'
' See ArryReDim for re-specifying any of the dimensions ans even adding new
' dimensions)
'
' Constraints:
' - For a yet not dimensioned and/or not allocated array items may be added
'   by simply not specifying an index
' - For an already dimensioned and/or allocated array the provision of an
'   index for each of its dimensions is obligatory.
'
' Uses: ArryBounds
'
' W. Rauschenberger Berlin, Jan 2025
' ----------------------------------------------------------------------------
    Const PROC = "Arry(Let)"
        
    On Error GoTo eh
    Dim cllSpecNdcs     As New Collection   ' the specified indices as Collection
    Dim cllDimSpecs     As Collection   ' the provided/input array's dimension from/to specifics
    Dim lBase           As Long
    Dim cllBounds       As Collection
    Dim cllBoundsOut    As Collection
    Dim lDimsArry       As Long         ' the provided/input array's number of dimensions
    Dim lDimsSpec       As Long         ' the specifies dimensions derived from the provided indices
    Dim lDimsOut        As Long
    
    lBase = LBound(Array(1))
    
    '~~ Get the provided array's number of dimesion (lDimsArry) and their from/to specifics (cllDimSpecs)
    ArryDims a_arr, cllDimSpecs, lDimsArry      ' DimsArray will be 0 when yet not allocated or not an array
    
    '~~ Get the provided indices as Collection
    Set cllSpecNdcs = ArryNdcsCll(a_indices)
    lDimsSpec = cllSpecNdcs.Count               ' may be 0 when none had been provided
    
    If lDimsArry <> 0 Then
        '~~ The array has at least one Item
        Select Case True
            Case lDimsArry > 1 And lDimsSpec <> lDimsArry
                Err.Raise AppErr(4), ErrSrc(PROC), "For an allocated multidimensional array the provided a_indices are incomplete!"
            
            Case lDimsSpec = 0 And lDimsArry > 1
                '~~ When for an allocated multi-dim array no index has been provided an error is raised
                Err.Raise AppErr(3), ErrSrc(PROC), "To write to a multi-dimensional array an appropriate index is required!"
            
            Case lDimsSpec = 0 And lDimsArry = 1
                '~~ When for an allocated 1-dim array no index is provided, the Item is added to a 1-dim array
                ReDim Preserve a_arr(LBound(a_arr) To UBound(a_arr) + 1)
                a_arr(UBound(a_arr)) = a_var
                
            Case lDimsArry = 1 And cllSpecNdcs(1) > UBound(a_arr)
                '~~ When for an Item in a 1-dim array an index beyond the current specs is provided the array is redimed/epanded accordingly
                ReDim Preserve a_arr(cllDimSpecs(1)(1) To cllSpecNdcs(1))
                a_arr(cllSpecNdcs(1)) = a_var
                
            Case lDimsArry = 1 And cllSpecNdcs(1) <= UBound(a_arr)
                a_arr(cllSpecNdcs(1)) = a_var
            
            Case lDimsArry > 1 And lDimsArry = lDimsSpec
                '~~ The dimensions specified are identical with those of the provided array
                '~~ The dimensios' index may still differ
                ArryBounds a_arr, cllSpecNdcs, cllBoundsOut, cllBounds, lDimsOut
                If lDimsOut = 0 Or lDimsOut = 1 And IsArray(cllBoundsOut(cllBoundsOut.Count)) Then
                    '~~ Either no bounds are out or the out-bound dimension is the last one
                    '~~ which can be re-dimed by VBA's ReDim.
                    Select Case lDimsArry
                        Case 1: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2))
                                a_arr(cllSpecNdcs(1)) = a_var
                        
                        Case 2: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2), cllBounds(2)(1) To cllBounds(2)(2))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2)) = a_var
                        
                        Case 3: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2), cllBounds(2)(1) To cllBounds(2)(2), cllBounds(3)(1) To cllBounds(3)(2))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)) = a_var
                        
                        Case 4: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2), cllBounds(2)(1) To cllBounds(2)(2), cllBounds(3)(1) To cllBounds(3)(2), cllBounds(4)(1) To cllBounds(4)(2))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)) = a_var
                        
                        Case 5: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)) = a_var
                        
                        Case 6: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)) = a_var
                        
                        Case 7: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)) = a_var
                        
                        Case 8: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)) = a_var
                    End Select
                Else
                    Select Case lDimsArry
                        Case 1: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",")
                                a_arr(cllSpecNdcs(1)) = a_var
                        
                        Case 2: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2)) = a_var
                        
                        Case 3: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)) = a_var
                        
                        Case 4: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)) = a_var
                        
                        Case 5: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)) = a_var
                        
                        Case 6: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",") _
                                               , "6:" & Join(cllBounds(6), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)) = a_var
                        
                        Case 7: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",") _
                                               , "6:" & Join(cllBounds(6), ",") _
                                               , "7:" & Join(cllBounds(7), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)) = a_var
                        
                        Case 8: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",") _
                                               , "6:" & Join(cllBounds(6), ",") _
                                               , "7:" & Join(cllBounds(7), ",") _
                                               , "8:" & Join(cllBounds(8), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)) = a_var
                    End Select
                
                End If
            Case Else
                Err.Raise AppErr(6), ErrSrc(PROC), "This is a case of writing to an allocated array yet not considered/implemented!"
        End Select
        
    Else
        '~~ The provided array may yet not have been specified of is empty
        If lDimsSpec = 0 Then
            '~~ Writing to a yet un-allocated or yet un-specified array with
            '~~ no index provided the Item is the first of a new 1-dim array
            ReDim a_arr(lBase To lBase)
            a_arr(lBase) = a_var
        Else
            '~~ For a yet not allocated array an index for a 1- or multi-dimensional array had been provided
            Select Case lDimsSpec
                Case 1: ReDim a_arr(cllSpecNdcs(1))
                        a_arr(cllSpecNdcs(1)) = a_var
                Case 2: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2)):                                                                                        a_arr(cllSpecNdcs(1), cllSpecNdcs(2)) = a_var
                Case 3: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)):                                                                          a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)) = a_var
                Case 4: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)):                                                            a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)) = a_var
                Case 5: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)):                                              a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)) = a_var
                Case 6: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)):                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)) = a_var
                Case 7: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)):                  a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)) = a_var
                Case 8: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)):    a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)) = a_var
            End Select
        End If
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get Arry(Optional ByRef a_arr As Variant, _
                          Optional ByVal a_indices As Variant = Nothing) As Variant
' ----------------------------------------------------------------------------
' Common, universal READ from array service supporting up to 8 dimensions.
' The service returns:
' - The Item of a provided array (a_arr) at a given index (a_indices) which
'   might be a single integer or an array or Collection of integers indicating
'   the index for each dimension of a multi-dimensional array.
' - The default (a_default) when
'   - the provided array (a_arr) has no content
'   - the provided array isn't one
'   - the provided index is (indices are) out of the bound/s of any dimension.
'
' W. Rauschenberger Berlin, Jan 2025
' ----------------------------------------------------------------------------
    Dim lDims  As Long
    Dim bObject As Boolean
    
    Select Case TypeName(a_arr)
        Case "Byte()":     Arry = 0
        Case "Integer()":  Arry = 0
        Case "Long()":     Arry = 0
        Case "Single()":   Arry = 0
        Case "Double()":   Arry = 0
        Case "Currency()": Arry = 0
        Case "Date()":     Arry = #12:00:00 AM#
        Case "String()":   Arry = vbNullString
        Case "Boolean()":  Arry = False
        Case "Variant()":  Arry = Empty
        Case "Object()":   Set Arry = Nothing: bObject = True
        Case Else:         Arry = Empty
    End Select
    
    lDims = ArryDims(a_arr) ' This will return 0 for anything not an array or not a specified array
    
    If lDims > 0 Then
        Set a_indices = ArryNdcsCll(a_indices) ' transforms any kind of provided inedx/indices into a Collection (1 to n)
    End If
    
    On Error Resume Next
    Select Case lDims
        Case 0:
        Case 1: If bObject Then Set Arry = a_arr(a_indices(1)) Else Arry = a_arr(a_indices(1))
        Case 2: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2)) Else Arry = a_arr(a_indices(1), a_indices(2))
        Case 3: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3))
        Case 4: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4))
        Case 5: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5))
        Case 6: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6))
        Case 7: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7))
        Case 8: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7), a_indices(8)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7), a_indices(8))
    End Select
    
xt:
End Property

Private Property Get ArryItem(Optional ByVal a_arr As Variant, _
                              Optional ByVal a_ndcs As Variant) As Variant
' ---------------------------------------------------------------------------
' Returns from an array (a_arr) the Item addressed by the indices (a_ndcs)
' which might be up to 5 dimensions.
' ---------------------------------------------------------------------------
    Select Case UBound(a_ndcs)
        Case 0: ArryItem = a_arr(a_ndcs(0))
        Case 1: ArryItem = a_arr(a_ndcs(0), a_ndcs(1))
        Case 2: ArryItem = a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2))
        Case 3: ArryItem = a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3))
        Case 4: ArryItem = a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4))
    End Select
End Property

Private Property Let ArryItem(Optional ByVal a_arr As Variant, _
                              Optional ByVal a_ndcs As Variant, _
                                       ByVal a_item As Variant)
' ---------------------------------------------------------------------------
' Writes an Item (a_Item) to an array (a_arr) by means of provided indices
' (a_ndcs) which covers up to 8 dimensions.
' ---------------------------------------------------------------------------
    
    Select Case UBound(a_ndcs)
        Case 0: a_arr(a_ndcs(0)) = a_item
        Case 1: a_arr(a_ndcs(0), a_ndcs(1)) = a_item
        Case 2: a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2)) = a_item
        Case 3: a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3)) = a_item
        Case 4: a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4)) = a_item
        Case 5: a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5)) = a_item
        Case 6: a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6)) = a_item
        Case 7: a_arr(a_ndcs(0), a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7)) = a_item
    End Select
    
End Property

' ----------------------------------------------------------------------------
' Unified Collection interface.
' Read:  Returns any index, when beyond count the default (c_default)
' Write: - When no index is provided the Item is added
'        - When an existing index is provided the Item is replaced
'        - When an index beyond the current number of items in the Collection
'          is provided the Item is added at the given position and any items
'          in between are provided with the default (c_default)
' ----------------------------------------------------------------------------
Private Property Get Coll(Optional ByRef c_cll As Collection, _
                          Optional ByVal c_index As Long = -999, _
                          Optional ByVal c_default As Variant = vbNullString) As Variant
    On Error Resume Next
    If c_cll Is Nothing Then Set c_cll = New Collection
    If c_index <= c_cll.Count Then
        If VBA.IsObject(c_cll(c_index)) _
        Then Set Coll = c_cll(c_index) _
        Else Coll = c_cll(c_index)
    Else
        If VBA.IsObject(c_default) _
        Then Set Coll = c_default _
        Else Coll = c_default
    End If
    
End Property

Private Property Let Coll(Optional ByRef c_cll As Collection, _
                          Optional ByVal c_index As Long = -999, _
                          Optional ByVal c_default As Variant = vbNullString, _
                                   ByVal c_Item As Variant)
    Dim cll As New Collection
    Dim i   As Long
    
    If c_cll Is Nothing Then Set c_cll = New Collection
    With cll
        Select Case True
            Case c_index = -999
                c_cll.Add c_Item
            Case c_index <= c_cll.Count
                For i = 1 To c_cll.Count
                    Select Case True
                        Case i <> c_index
                            .Add c_cll(i)
                        Case i = c_index
                            .Add c_Item
                    End Select
                Next i
                
            Case c_index > c_cll.Count
                For i = 1 To c_cll.Count
                    .Add c_cll(i)
                Next i
                Do While cll.Count + 1 < c_index
                    .Add c_default
                Loop
                .Add c_Item
        End Select
    End With
    Set c_cll = cll
    Set cll = Nothing

End Property

Public Property Let Comment(ByVal s As String):                                                         sComment = s:                                                   End Property

Private Property Let Dict(Optional ByRef d_dct As Dictionary, _
                          Optional ByVal d_key As Variant, _
                                   ByVal d_Item As Variant)
    If d_dct Is Nothing Then Set d_dct = New Dictionary
    If d_dct.Exists(d_key) Then d_dct.Remove d_key
    d_dct.Add d_key, d_Item
    
End Property

Friend Property Get InstructionsPosLeft() As Single:                                                    InstructionsPosLeft = siInstructionsPosLeft:                    End Property

Friend Property Let InstructionsPosLeft(si As Single):                                                  siInstructionsPosLeft = si:                                     End Property

Friend Property Get InstructionsPosTop() As Single:                                                     InstructionsPosTop = siInstructionsPosTop:                      End Property

Friend Property Let InstructionsPosTop(si As Single):                                                   siInstructionsPosTop = si:                                      End Property

Friend Property Get InstructionsTitle() As String:                                                      InstructionsTitle = siInstructionsTitle:                        End Property

Private Property Get LogItemComment(Optional ByRef l_items As Variant) As Variant:                       LogItemComment = Arry(l_items, IX_LOGITEM_COMMENT):             End Property

Private Property Let LogItemComment(Optional ByRef l_items As Variant, ByVal l_item As Variant):         Arry(l_items, IX_LOGITEM_COMMENT) = l_item:                     End Property

Private Property Get LogItemCompName(Optional ByRef l_items As Variant) As Variant:                      LogItemCompName = Arry(l_items, IX_LOGITEM_COMP_NAME):          End Property

Private Property Let LogItemCompName(Optional ByRef l_items As Variant, ByVal l_item As Variant):        Arry(l_items, IX_LOGITEM_COMP_NAME) = l_item:                   End Property

Private Property Get LogItemProcName(Optional ByRef l_items As Variant) As Variant:                      LogItemProcName = Arry(l_items, IX_LOGITEM_PROC_NAME):          End Property

Private Property Let LogItemProcName(Optional ByRef l_items As Variant, ByVal l_item As Variant):        Arry(l_items, IX_LOGITEM_PROC_NAME) = l_item:                   End Property

Private Property Get LogItemProcType(Optional ByRef l_items As Variant) As Variant:                      LogItemProcType = Arry(l_items, IX_LOGITEM_PROC_TYPE):          End Property

Private Property Let LogItemProcType(Optional ByRef l_items As Variant, ByVal l_item As Variant):        Arry(l_items, IX_LOGITEM_PROC_TYPE) = l_item:                   End Property

Private Property Get LogItemTestId(Optional ByRef l_items As Variant) As Variant:                        LogItemTestId = Arry(l_items, IX_LOGITEM_TEST_ID):              End Property

Private Property Let LogItemTestId(Optional ByRef l_items As Variant, ByVal l_item As Variant):          Arry(l_items, IX_LOGITEM_TEST_ID) = l_item:                     End Property

Private Property Get LogItemTitle(Optional ByRef l_items As Variant) As Variant:                         LogItemTitle = Arry(l_items, IX_LOGITEM_TITLE):                 End Property

Private Property Let LogItemTitle(Optional ByRef l_items As Variant, ByVal l_item As Variant):           Arry(l_items, IX_LOGITEM_TITLE) = l_item:                       End Property

Private Property Get LogItemVrfctnDesc(Optional ByRef l_items As Variant) As Variant:                    LogItemVrfctnDesc = Arry(l_items, IX_LOGITEM_VRFCTN):           End Property

Private Property Let LogItemVrfctnDesc(Optional ByRef l_items As Variant, ByVal l_item As Variant):      Arry(l_items, IX_LOGITEM_VRFCTN) = l_item:                      End Property

Private Property Get LogItemVrfctnNo(Optional ByRef l_items As Variant) As Variant:                      LogItemVrfctnNo = Arry(l_items, IX_LOGITEM_VRFCTN_NO):          End Property

Private Property Let LogItemVrfctnNo(Optional ByRef l_items As Variant, ByVal l_item As Variant):        Arry(l_items, IX_LOGITEM_VRFCTN_NO) = l_item:                   End Property

Private Property Get LogItemVrfctnRes(Optional ByRef l_items As Variant) As Variant:                     LogItemVrfctnRes = Arry(l_items, IX_LOGITEM_RES):               End Property

Private Property Let LogItemVrfctnRes(Optional ByRef l_items As Variant, ByVal l_item As Variant):       Arry(l_items, IX_LOGITEM_RES) = l_item:                         End Property

Private Property Get LogItemVrfctnResExp(Optional ByRef l_items As Variant) As Variant:                  LogItemVrfctnResExp = Arry(l_items, IX_LOGITEM_RES_EXP):        End Property

Private Property Let LogItemVrfctnResExp(Optional ByRef l_items As Variant, ByVal l_item As Variant):    Arry(l_items, IX_LOGITEM_RES_EXP) = l_item:                     End Property

Private Property Get LogItemVrfctnResult(Optional ByRef l_items As Variant) As Variant:                  LogItemVrfctnResult = Arry(l_items, IX_LOGITEM_VRFCTN_RSLT):    End Property

Private Property Let LogItemVrfctnResult(Optional ByRef l_items As Variant, ByVal l_item As Variant):    Arry(l_items, IX_LOGITEM_VRFCTN_RSLT) = l_item:                 End Property

Private Property Get LogItemVrfydType(Optional ByRef l_items As Variant) As Variant:                     LogItemVrfydType = Arry(l_items, IX_LOGITEM_VR_TYPE):        End Property

Private Property Let LogItemVrfydType(Optional ByRef l_items As Variant, ByVal l_item As Variant):       Arry(l_items, IX_LOGITEM_VR_TYPE) = l_item:                  End Property

Private Property Get LogItemVxcTime(Optional ByRef l_items As Variant) As Variant:                       LogItemVxcTime = Arry(l_items, IX_LOGITEM_EXEC_TIME):           End Property

Private Property Let LogItemVxcTime(Optional ByRef l_items As Variant, ByVal l_item As Variant):         Arry(l_items, IX_LOGITEM_EXEC_TIME) = l_item:                   End Property

Friend Property Get NameTestFile() As String:                                                           NameTestFile = sNameTestFile:                                   End Property

Friend Property Get NameTestResultExpectedFile() As String:                                             NameTestResultExpectedFile = sNameTestResultExpectedFile:       End Property

Friend Property Get NameTestResultFile() As String:                                                     NameTestResultFile = sNameTestResultFile:                       End Property

Friend Property Get Result(Optional ByVal r_exclude As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
' Returns the most current registered test Result.
' ------------------------------------------------------------------------------
    
    r_exclude = r_exclude
    If VBA.IsObject(vResult) _
    Then Set Result = vResult _
    Else Result = vResult

End Property

Friend Property Let Result(Optional ByVal r_exclude As String = vbNullString, _
                                    ByVal r_var As Variant)
' ------------------------------------------------------------------------------
' Registers a Result in the dctResult Directory. An existing corresponding
' Result expected in the dctResultExpected Directory triggers the
' ResultVersusExpected.
' ------------------------------------------------------------------------------
    If VarType(r_var) = vbInteger Then r_var = CLng(r_var)
    If r_exclude <> vbNullString Then sExcludeFromComparison = r_exclude
    If VBA.IsObject(r_var) Then Set vResult = r_var Else vResult = r_var
    bResult = True
    If bResultExpected Then ResultVersusExpected r_var, vResultExpected
    
End Property

Friend Property Get ResultExpected(Optional ByVal r_exclude As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
' Returns the most current registered expected test Result.
' ------------------------------------------------------------------------------
    
    r_exclude = r_exclude
    If VBA.IsObject(vResultExpected) _
    Then Set ResultExpected = vResultExpected _
    Else ResultExpected = vResultExpected

End Property

Friend Property Let ResultExpected(Optional ByVal r_exclude As String = vbNullString, _
                                            ByVal r_var As Variant)
' ------------------------------------------------------------------------------
' Registers an expected Result in the dctResultExpected Directory. An existing
' corresponding test Result in the dctResult Directory triggers the
' ResultVersusExpected.
' ------------------------------------------------------------------------------
    If VarType(r_var) = vbInteger Then r_var = CLng(r_var)
    If r_exclude <> vbNullString Then sExcludeFromComparison = r_exclude
    If VBA.IsObject(r_var) Then Set vResultExpected = r_var Else vResultExpected = r_var
    bResultExpected = True
    If bResult Then ResultVersusExpected vResult, r_var
    
End Property

Friend Property Let TempTestItem(Optional ByVal t_test_no As String = vbNullString, _
                                          ByVal t_Item As Variant)
' ------------------------------------------------------------------------------
' - When no test-number (t_test_no) is provided and the provided test Item
'   (t_Item) exists (may be a file, a folder, or a Workbook), a record is kept
'   in a file (sTestItems) for being removed (when a file or a folder)
'   or closed (when a Workbook) with the CleanUp service.
' - When a test-number is provided the full default name of an existing test
'   file or folder becomes the current test Item.
' ------------------------------------------------------------------------------
    Const PROC = "TempTestItem(Let)"
        
    If t_test_no = vbNullString Then
        '~~ Keep a record of the test Item
        Select Case True
            Case TypeName(t_Item) = "String"
                cllTempTestItems.Add t_Item
                vTempTestItem = t_Item
            Case VBA.IsObject(t_Item)
                cllTempTestItems.Add t_Item
                Set vTempTestItem = t_Item
            Case Else:                          Err.Raise AppErr(1), ErrSrc(PROC), "Unexpected kind of test Item (record not kept)!"
        End Select
    Else
        '~~ Provide the test Item
        vTempTestItem = TestFileFullName
    End If
    
End Property

Friend Property Let TestedComp(ByVal s As String):                  sTestedComp = s:                                            End Property

Friend Property Let TestedProc(ByVal s As String):                  sTestedProc = s:                                            End Property

Friend Property Let TestedProcType(ByVal s As String):              sTestedProcType = s:                                        End Property

Friend Property Get TestFileExtension() As String:                  TestFileExtension = sTestFileExtension:                     End Property

Friend Property Let TestFileExtension(ByVal s As String):           sTestFileExtension = Replace("." & s, "..", "."):           End Property

Public Property Let TestFileNameScheme(ByVal s_scheme As String)
' ----------------------------------------------------------------------------
' Determines the schem of test file names.
' ----------------------------------------------------------------------------
    Const PROC          As String = "TestFileNameScheme"
    
    Dim s As String
    s = s_scheme
    s = Replace(s, SCHEME_ID, vbNullString)
    s = Replace(s, SCHEME_NAME, vbNullString)
    s = Replace(s, SCHEME_VER_NO, vbNullString)
    s = Replace(s, SCHEME_EXT, vbNullString)
    s = Replace(s, "-", vbNullString)
    s = Replace(s, ".", vbNullString)
    If s <> "Test" Then _
    Err.Raise AppErr(1), ErrSrc(PROC), "The provided test file name scheme is invalid! " & vbLf & _
                                       "It contains other strings than ""<id>"", ""<nm>"", ""<vn>"", ""<xt>"", ""-"", and "".""."

End Property

Friend Property Get TestFolder() As String:                         TestFolder = sTestFolder:                                   End Property

Friend Property Let TestFolder(ByVal s As String)
    
    If sTestFolder <> vbNullString Then
        If sTestFolder <> s Then
            If fso.FolderExists(s) Then fso.DeleteFolder s
        End If
    End If
    sTestFolder = s
    If Not fso.FolderExists(sTestFolder) _
    Then fso.CreateFolder sTestFolder
    ChDir sTestFolder

End Property

Friend Property Get TestId() As String:                             TestId = sTestId:                                           End Property

Friend Property Let TestId(ByVal s As String)
' ------------------------------------------------------------------------------
' Essential, obligatory, distinct identification of a test, provided as the very
' first Property for each test. Resets a previous test to prepare for a new one.
' since the id is used in the test file's name it must be valid for being part
' of it.
' ------------------------------------------------------------------------------
    Const PROC = "TestId(Let)"
    
    If s = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Specifying a TestId is obligatory!"
        
    If Not IsValidFileName(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
                   "The provided Testid is not valid as (part of a) file or Worksheet name! || " & _
                   "The id is used in test file's name (.TestFileName, TestResultFileName, TestResultExpectedFileName) " & _
                   "and as a Worksheet name for failed test verifications. Thus it needs to be valid for becoming part of these names."
    sTestId = s
    vTempTestItem = vbNullString
    lVrfctnNumber = 1
    sVerification = vbNullString
    TestFileNames
    bResult = False
    bResultExpected = False
    Set wshLog = Nothing ' enforce initiation of a new log sheet
    ArryErase aLogDetails
    Set cllLogDetails = Nothing
    
End Property

Private Property Let TestResultSummaryLogFile(ByVal s As String)
    
    If sTestResultSummaryLogFile <> vbNullString Then
        If fso.FileExists(sTestResultSummaryLogFile) Then
            fso.DeleteFile sTestResultSummaryLogFile
        End If
    End If
    sTestResultSummaryLogFile = s
    
End Property

Friend Property Get TimerExecTimeMsecs() As Currency:               TimerExecTimeMsecs = dTimerMsecs:                           End Property

Public Property Get TimerOverheadMsec() As Double:                  TimerOverheadMsec = dTimerOverhead:                         End Property

Friend Property Get TimerOverheadTicks() As Currency:               TimerOverheadTicks = cTicksOverhead:                        End Property

Friend Property Get Title() As String:                              Title = sTitle:                                             End Property

Friend Property Let Title(ByVal s As String)
    sTitle = s
    lVrfctnNumber = 0
End Property

Friend Property Let TitleRegression(ByVal s As String):             sTitleRegression = s:                                       End Property

Private Property Get TxtFile(Optional ByVal t_file As String, _
                             Optional ByVal t_append As Boolean) As String
    Const PROC = "TxtFile(Get)"
    
    On Error GoTo eh
    Dim iFileNumber As Integer
    
    iFileNumber = FreeFile
    Open t_file For Input As #iFileNumber
    TxtFile = Input$(LOF(iFileNumber), iFileNumber)
    Close #iFileNumber

xt: Exit Property
 
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let TxtFile(Optional ByVal t_file As String, _
                             Optional ByVal t_append As Boolean, _
                                     ByVal t_strng As String)
    Const PROC = "TxtFile"
    
    On Error GoTo eh
    Dim iFileNumber As Integer
    
    iFileNumber = FreeFile
    If t_append _
    Then Open t_file For Append As #iFileNumber _
    Else Open t_file For Output As #iFileNumber
    Print #iFileNumber, t_strng
    Close #iFileNumber
                              
xt: Exit Property
 
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get Verification() As String
    Const PROC = "Verification(Get)"
    
    If sVerification = vbNullString And sTitle <> vbNullString Then sVerification = sTitle
    If sVerification = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "Neither a .Verification not a .Title had been specified for the evaluation/comparison of a test Result versus a Result expected (test-id " & sTestId & ")." & "||" & _
         "When it comes to a Result evaluation a specified .Verification or alternatively a given .Title is obligatory."

    Verification = sVerification
    
End Property

Friend Property Let Verification(ByVal s As String)
' ----------------------------------------------------------------------------
' Obligatory prior any ResultVersusExpected. However, when a test identified by a
' u nique .TestId has only one .Result/.ResultExpected a missing .Verification
' is the Title is used alternatively - when it
' had been provided. When neither has been provided an error is raised. For
' any addidtional verification (.Result/.ResultExpected) the provision of a
' description (this property) is obligatory.
' Note: Verifications are indicated by a .number attached to the TestId.
' ----------------------------------------------------------------------------
    
    sVerification = s
    TestFileNames
    lVrfctnNumber = lVrfctnNumber + 1 ' for a possible next verification within the same TestId
    bResultExpected = False
    bResult = False
    
End Property

Friend Property Get VrfctnNumber() As Long: VrfctnNumber = lVrfctnNumber:  End Property

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

Private Function ArryAsColl(ByVal a_arr As Variant) As Collection
' ----------------------------------------------------------------------------
' Return an array's (a_arr) items as Collection.
'
' See also https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = ""
    
    Dim cll As New Collection
    Dim v   As Variant
    
    If ArryDims(a_arr) <> 1 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided array has more than one dimension but multidimensions are not supported!"
    
    With cll
        For Each v In a_arr
            .Add v
        Next v
    End With
    Set ArryAsColl = cll
    Set cll = Nothing
    
End Function

Public Function ArryAsDict(ByVal a_arr As Variant) As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary with all items of an array (a_arr) with the indices
' delimited by a comma as key.
' Note: The function handlesmultidimensional array up to 5 dimensions.
' Uses: ArryAsDictAdd
' ------------------------------------------------------------------------------
    
    Dim dct As New Dictionary
    
    ArryAsDictAdd a_arr, dct, "", 1, ArryDims(a_arr)
    Set ArryAsDict = dct
    
End Function

Private Sub ArryAsDictAdd(ByVal a_arr As Variant, _
                          ByRef a_dct As Dictionary, _
                          ByRef a_key As String, _
                          ByRef a_dim As Integer, _
                          ByVal a_dims As Integer)
' ------------------------------------------------------------------------------
' Adds an item of an array (a_arr), addressed by indices - derived from the key
' (a_key) - indices delimited by commas - to a Dictionary (a_dct) with the
' provided key (a_key). Empty items in the array (a_arr) are ignored.
' ------------------------------------------------------------------------------
    Const PROC = "ArryAsDictAdd"
    
    On Error GoTo eh
    Dim i       As Long
    Dim sKey    As String
    Dim ArrNdcs As Variant
    Dim vItem   As Variant
    
    For i = LBound(a_arr, a_dim) To UBound(a_arr, a_dim)
        If a_key = vbNullString _
        Then sKey = i _
        Else sKey = a_key & "," & i
        
        If a_dim < a_dims Then
            ArryAsDictAdd a_arr, a_dct, sKey, a_dim + 1, a_dims
        Else
            ArrNdcs = Split(sKey, ",")
            vItem = ArryItem(a_arr, ArrNdcs)
            If Not IsError(vItem) And Not IsEmpty(vItem) Then
                a_dct.Add sKey, vItem
            End If
        End If
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Function ArryAsFile(ByVal v_items As Variant, _
                    Optional ByRef v_file_name As String = vbNullString, _
                    Optional ByVal v_file_append As Boolean = True) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
      
    If v_file_name = vbNullString Then v_file_name = TempFileFullName
    StringAsFile ArryAsStrg(v_items), v_file_name, v_file_append
    Set ArryAsFile = fso.GetFile(v_file_name)
    
End Function

Private Sub ArryAsRnge(ByVal a_arr As Variant, _
                       ByRef a_rng As Range, _
              Optional ByVal a_one_col As Boolean = False)
' ----------------------------------------------------------------------------
' Transferes the content of an array (a_arr) to a range (a_rng) amd returns
' the resulting range.
' ----------------------------------------------------------------------------
    Const PROC = "ArryAsRnge"
    
    On Error GoTo eh
    Dim rTarget As Range

    If a_one_col Then
        '~~ One column, n rows
        Set rTarget = a_rng.Cells(1, 1).Resize(UBound(a_arr) + 1, 1)
        rTarget.value = Application.Transpose(a_arr)
    Else
        '~~ One column, n rows
        Set rTarget = a_rng.Cells(1, 1).Resize(1, UBound(a_arr) + 1)
        rTarget.value = a_arr
    End If
    Set a_rng = rTarget
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Function ArryAsStrg(ByVal c_array As Variant, _
                      Optional ByVal c_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (c_array) as string with the items delimited (c_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArryAsStrg = Join(c_array, c_delim)
End Function

Private Function ArryBounds(ByVal a_arr As Variant, _
                            ByVal a_indices As Variant, _
                   Optional ByRef a_out_bounds As Collection, _
                   Optional ByRef a_in_bounds As Collection, _
                   Optional ByRef a_out As Long) As Boolean
' ---------------------------------------------------------------------------
' Returns:
' - TRUE when all dimensions addressed by indices (a_indices) are
'   within the bounds of the respective dimension in array (a_arr)
' - FALSE when any of the provided indices (a_indices) is out of the bounds
'   of the provided array (a_arr)
' - Returns the dimesions which are out of bounds as Collection with items
'   in-bound empty and out-bound with the new bound
' - Returns the complete dimension specifics which combine the "from" spec
'   of the provided array with the new "to" specs in case they are greater
'   than the present ones
'
' Precondition: The indices are provided (a_indices) is either as a single
'               integer - when the array (a_arr) is a 1-dim array - or an
'               array of integers, each specifying the index for one
'               dimension.
'
' Uses: Coll
'
' W. Rauschenberger, Berlin Jan 2025
' ---------------------------------------------------------------------------
    Dim aBounds(1 To 2)     As Variant
    Dim aBoundsOut(1 To 2)  As Variant
    Dim cllBoundsIn         As New Collection
    Dim cllBoundsOut        As New Collection
    Dim cllSpecNdcs         As Collection
    Dim i                   As Long
    Dim lDimsArry           As Long
    Dim lDimsSpec           As Long
    
    lDimsArry = ArryDims(a_arr)
    Set cllSpecNdcs = ArryNdcsCll(a_indices)
    lDimsSpec = cllSpecNdcs.Count
    
    If lDimsSpec > lDimsArry Then GoTo xt
    For i = 1 To cllSpecNdcs.Count
        aBounds(1) = Min(cllSpecNdcs(i), LBound(a_arr, i))
        aBounds(2) = Max(cllSpecNdcs(i), UBound(a_arr, i))
        Coll(cllBoundsIn, i) = aBounds
        If cllSpecNdcs(i) < LBound(a_arr, i) Or cllSpecNdcs(i) > UBound(a_arr, i) Then
            aBoundsOut(1) = LBound(a_arr, i)
            aBoundsOut(2) = UBound(a_arr, i)
            Coll(cllBoundsOut, i) = aBoundsOut
            a_out = a_out + 1
        Else
            Coll(cllBoundsOut, i) = Empty
        End If
    Next i
    
    Set a_in_bounds = cllBoundsIn
    Set a_out_bounds = cllBoundsOut
    ArryBounds = cllBoundsOut.Count > 0
    Set cllBoundsIn = Nothing
    Set cllBoundsOut = Nothing
xt:
End Function

Private Function ArryDefault(ByVal a_arry As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the default value of an array (a_arry) based on the TypeName.
' When the provided argument is not an array, Empty is returned.
' ----------------------------------------------------------------------------
    
    If Not IsArray(a_arry) Then
        ArryDefault = Empty
    Else
        Select Case TypeName(a_arry)
            Case "Byte()":     ArryDefault = 0
            Case "Integer()":  ArryDefault = 0
            Case "Long()":     ArryDefault = 0
            Case "Single()":   ArryDefault = 0
            Case "Double()":   ArryDefault = 0
            Case "Currency()": ArryDefault = 0
            Case "Date()":     ArryDefault = #12:00:00 AM#
            Case "String()":   ArryDefault = vbNullString
            Case "Boolean()":  ArryDefault = False
            Case "Variant()":  ArryDefault = Empty
            Case "Object()":   Set ArryDefault = Nothing
            Case Else:         ArryDefault = Empty
        End Select
    End If
    
End Function

Private Function ArryDiffs(ByVal v_res As Variant, _
                           ByVal v_exp As Variant, _
                  Optional ByRef v_res_prfx As String = vbNullString, _
                  Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the an array (v_res) differs from another (v_exp).
' ----------------------------------------------------------------------------
    Const PROC = "ArryDiffs"
    
    On Error GoTo eh
    Dim dctRes  As Dictionary
    Dim dctExp  As Dictionary
    Dim dctDone As New Dictionary
    
    Set dctRes = ArryAsDict(v_res)
    Set dctExp = ArryAsDict(v_exp)

    If dctRes.Count <> dctExp.Count Then
        sFailedNote = "Number of items Result/expected differs!"
        LogCollectVrfctnDetails vbNullString, vbNullString, VR_FAILED
    Else
        LogCollectVrfctnDetails vbNullString, vbNullString, vbNullString
    End If
    For Each v In dctRes
        If Not dctExp.Exists(v) Then
            sFailedNote = "Item not expected!"
            LogCollectVrfctnDetails "Item(" & v & ") = " & CStr(dctRes(v)), vbNullString, VR_FAILED
            ArryDiffs = True
        Else
            If VarDiffs(dctRes(v), dctExp(v), "Item(" & v & ") = ", v_exp_prfx) Then
                ArryDiffs = True
            End If
            If Not dctDone.Exists(v) Then dctDone.Add v, vbNullString
        End If
    Next v
    
    For Each v In dctExp
        If Not dctDone.Exists(v) Then
            sFailedNote = "Item(" & v & ") missing in Result!"
            LogCollectVrfctnDetails vbNullString, "Item(" & v & ") = " & CStr(dctExp(v)), VR_FAILED
            ArryDiffs = True
        End If
    Next v
    
    Set dctDone = Nothing
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Private Function ArryDims(ByVal a_arr As Variant, _
                 Optional ByRef a_dim_specs As Collection, _
                 Optional ByRef a_dims As Long) As Long
' ----------------------------------------------------------------------------
' Returns:
' - the dimensions of an array (a_arr), optionally also as argument (a_dims)
' - the from to specs (a_dimsfrom, a_dimsto) delimited by kommas for each dim
' - an arry with the from/to specs, representing th dimensions 1 to n
' ----------------------------------------------------------------------------
    Dim arrSpecs(1 To 2) As Variant
    Dim cll              As New Collection
    Dim i                As Long
    
    With cll
        For i = 1 To 8
            On Error Resume Next
            arrSpecs(1) = LBound(a_arr, i)
            If Err.Number <> 0 Then
                Exit For
            Else
                arrSpecs(2) = UBound(a_arr, i)
                .Add arrSpecs
            End If
        Next i
    End With
        
xt: ArryDims = cll.Count
    Set a_dim_specs = cll
    a_dims = cll.Count
    Set cll = Nothing

End Function

Private Function ArryErase(ByRef a_arr As Variant)
    If IsArray(a_arr) Then Erase a_arr
End Function

Private Function ArryHasElements(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the array (a_arr) has at least one element.
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    On Error Resume Next
    v = a_arr(LBound(a_arr))
    ArryHasElements = Err.Number = 0
    Err.Clear

End Function

Private Function ArryNdcsCll(ParamArray a_indices() As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns provided indices (a_indices) as Collection whereby the indices may
' be provided: - as integers (one for each dimension)
'              - as an array of integers
'              - as a string of integers delimited by a , (comma).
' ----------------------------------------------------------------------------
    Const PROC = "ArryIncides"
    
    Dim cll As New Collection
    Dim v   As Variant
    
    On Error GoTo xt
    With cll
        If UBound(a_indices) >= LBound(a_indices) Then
            Select Case True
                Case Not ArryIsAllocated(a_indices)
                    '~~ No indices had been provided
                    Exit Function
                Case IsArray(a_indices(LBound(a_indices)))
                    '~~ The first item is an array if incides
                    For Each v In a_indices(LBound(a_indices))
                        If IsInteger(CInt(Trim(v))) _
                        Then .Add v _
                        Else Err.Raise AppErr(1), ErrSrc(PROC), "At least one of the items provided as array is not an integer value!"
                    Next v
                Case TypeName(a_indices(LBound(a_indices))) = "Collection"
                    '~~ The first item is a Collection of indices
                    Set cll = a_indices(LBound(a_indices))
                Case TypeName(a_indices(LBound(a_indices))) = "String"
                    '~~ The first item is a string of incides delimited by a comma
                    For Each v In Split(a_indices(LBound(a_indices)), ",")
                        If IsNumeric(v) Then
                            .Add CInt(Trim(v))
                        Else
                            Err.Raise AppErr(2), ErrSrc(PROC), "At least one of the items provided a string delimited by a comma is not an integer value!"
                        End If
                    Next v
                Case a_indices(LBound(a_indices)) = Empty
                Case Else
                    For Each v In a_indices
                        If IsInteger(v) _
                        Then .Add v _
                        Else Err.Raise AppErr(3), ErrSrc(PROC), "Any of the provided indices is not an integer value!"
                    Next v
            End Select
        End If
    End With
    
xt: On Error GoTo 0
    Set ArryNdcsCll = cll
    Set cll = Nothing
    
End Function

Private Function ArryIsAllocated(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the array (a_arr) is allocated, i.e. has at least one Item.
' ----------------------------------------------------------------------------
    
    On Error Resume Next
    ArryIsAllocated = UBound(a_arr) >= LBound(a_arr)
    On Error GoTo 0
    Err.Clear
    
End Function

Private Function ArryItems(ByVal a_arr As Variant, _
                  Optional ByVal a_default_excluded As Boolean = False) As Long
' ----------------------------------------------------------------------------
' Returns the number of items in a multi-dimensional array or a nested array.
' The latter is an array of which one or more items are again possibly multi-
' dimensional arrays. An unallocated array returns 0. When items which return
' a type specific default are excluded (a_default_excluded = True) only
' "active= items/elements are counted.
' ----------------------------------------------------------------------------
    Dim lItems  As Long
    Dim lDims   As Long
    Dim v       As Variant
    Dim i       As Long
    Dim vDflt   As Variant
    If Not IsArray(a_arr) Then Exit Function
    
    lDims = ArryDims(a_arr)
    If Not a_default_excluded Then
        lItems = 1
        For i = 1 To lDims
            lItems = lItems * (UBound(a_arr, i) - LBound(a_arr, i) + 1)
        Next i
    Else
        vDflt = ArryDefault(a_arr)
        For Each v In a_arr
            If IsArray(v) Or TypeName(v) Like "*()" Then
                lItems = lItems + ArryItems(v, a_default_excluded)
            ElseIf Not TypeName(v) = "Error" And Not v = vDflt _
                Then lItems = lItems + 1
            End If
        Next v
    End If
    ArryItems = lItems
    
End Function

Private Function ArryNoOfElements(ByVal a_arr As Variant) As Long
' ----------------------------------------------------------------------------
' Returns the number of items/elements in a mustidimesional array (a_arr).
' Multidimensional arrays are covered by unloading tehm into a Dictionary.
' ----------------------------------------------------------------------------
    Dim dct As Dictionary
    
    Set dct = ArryAsDict(a_arr)
    ArryNoOfElements = dct.Count
    
End Function

'Private Sub ArryUnloadToDict(ByVal a_arr As Variant, _
'                                    ByRef a_dct As Dictionary, _
'                           Optional ByRef a_ndc As String)

Private Sub ArryReDim(ByRef a_arr As Variant, _
                 ParamArray a_dim() As Variant)
' ------------------------------------------------------------------------------
' Returns a provided multidimensional array (a_arr) with new dimension specifics
' (a_dim) whereby the new dimension specs (a_dim) are provided as strings
' following the format: "<dimension>:<from>,<to>" whereby <dimension> is either
' adressing a dimesion in the current array (a_arr), i.e. before the Redim has
' taken place, or a + for a new dimension. Since only new or dimensions with
' changed from/to specs are provided the information will be used to compile the
' final new target array's dimensions which must not exceed 8.
'
' Uses: ArryReDimSpecs
'
' Requires References to: - Microsoft Scripting Runtime
'                         - Microsoft VBScript Regular Expressions 5.5
' ------------------------------------------------------------------------------
    Const PROC = "ArryReDim"
    
    Dim arrOut          As Variant
    Dim arrUnloaded     As Variant
    Dim dct             As New Dictionary
    Dim i               As Long
    Dim sIndices        As String
    Dim sIndicesPrfx    As String
    
    On Error GoTo xt
    If UBound(a_arr) >= LBound(a_arr) Then
        On Error GoTo eh
        
        ReDim arrUnloaded(1 To ArryItems(a_arr))    ' redim the target array with the total number of items in the input array
        arrUnloaded = ArryUnload(a_arr)             ' unload the input array in a 2-dim array with Item 1 = indices and Item 2 = the input array's Item
        ArryReDimSpecs a_arr, a_dim, sIndicesPrfx, dct ' obtain the new and or changed dimension specifics
        
        '~~ Get confirmed that the total number of dimensions not exieeds the maximum suported
        If dct.Count > 8 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), _
                       "The target array's number of dimensions resulting from the provided " & _
                       "array and the provided dimesion specifics exeeds the maximum number of 8 dimensions!"
            
        '~~ Redim the new target array considering a maximum of 8 possible dimensions
        Select Case dct.Count
            Case 1: ReDim arrOut(dct(1)(0) To dct(1)(1))
            Case 2: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1))
            Case 3: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1), dct(3)(0) To dct(3)(1))
            Case 4: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1), dct(3)(0) To dct(3)(1), dct(4)(0) To dct(4)(1))
            Case 5: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1), dct(3)(0) To dct(3)(1), dct(4)(0) To dct(4)(1), dct(5)(0) To dct(5)(1))
            Case 6: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1), dct(3)(0) To dct(3)(1), dct(4)(0) To dct(4)(1), dct(5)(0) To dct(5)(1), dct(6)(0) To dct(6)(1))
            Case 7: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1), dct(3)(0) To dct(3)(1), dct(4)(0) To dct(4)(1), dct(5)(0) To dct(5)(1), dct(6)(0) To dct(6)(1), dct(7)(0) To dct(7)(1))
            Case 8: ReDim arrOut(dct(1)(0) To dct(1)(1), dct(2)(0) To dct(2)(1), dct(3)(0) To dct(3)(1), dct(4)(0) To dct(4)(1), dct(5)(0) To dct(5)(1), dct(6)(0) To dct(6)(1), dct(7)(0) To dct(7)(1), dct(8)(0) To dct(8)(1))
        End Select
        
        '~~ Re-load the unloaded array to the new specified re-dimed array and return it replacing the source array
        For i = LBound(arrUnloaded) To UBound(arrUnloaded)
            sIndices = sIndicesPrfx & arrUnloaded(i, 1)
            On Error Resume Next
            ArryItem(arrOut, sIndices) = arrUnloaded(i, 2)
        Next i
        a_arr = arrOut
    End If
    
xt: Err.Clear
    On Error GoTo 0
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub ArryReDimSpecs(ByVal a_arr As Variant, _
                           ByVal a_specs As Variant, _
                           ByRef a_prefx As String, _
                           ByRef a_dims As Dictionary)
    Const PROC = "ArryReDimSpecs"
    
    On Error GoTo eh
    Dim aItem       As Variant
    Dim lBase       As Long
    Dim lDims       As Long
    Dim lDimsNew    As Long
    Dim i           As Long
    Dim regex       As Object
    Dim sSpec       As String
    Dim v           As Variant
    
    lBase = LBound(Array("x"))
    lDims = ArryDims(a_arr)
    
    On Error GoTo xt
    If UBound(a_specs) >= LBound(a_specs) Then
        
        Set regex = CreateObject("VBScript.RegExp")
        '~~ 1. Load a Dictionary with any new specified dimensions (those whihc start with a + sign)
        For Each v In a_specs
            With regex
                .Pattern = "^(\+|\d)?\s*:\s*\d+(?:\s*,\s*\d+)?$"
                .IgnoreCase = False
                .Global = False
                If Not .Test(v) _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided dimension specification does not conform with expectations!" & vbLf & _
                                                        "A valid spec starts with a ""+"" (plus) or an integer from 1 to 8 followed " & _
                                                        "by a "":"" (semicolon), followed by to integers delimited by a "","" (comma)."
            End With
            
            If Trim$(Split(v, ":")(0)) = "+" Then
                lDimsNew = lDimsNew + 1
                sSpec = Trim(Split(v, ":")(1))
                If InStr(sSpec, ",") = 0 Then
                    '~~ The +: is followed by a single integer value.
                    '~~ This is interpreted as the Ubound of the new dimension
                    '~~ while the LBound defaults to the Base Options which is either 0 or 1
                    ReDim aItem(0 To 1)
                    aItem(0) = lBase
                    aItem(1) = CLng(sSpec)
                Else
                    aItem = Split(sSpec, ",")
                End If
                For i = LBound(aItem) To UBound(aItem)
                    aItem(i) = Trim(aItem(i))
                Next i
                a_dims.Add lDimsNew, aItem
            End If
        Next v
        If a_dims.Count > 0 Then
            '~~ Assemble the indices prefix from the new specified dimensions
            For Each v In a_dims
                a_prefx = a_prefx & a_dims(v)(0) & ","
            Next v
        End If
        
        '~~ 2. Add any redims specified - if any - for existing dimensions
        '~~    with redim specifications when available
        For Each v In a_specs
            If Trim$(Split(v, ":")(0)) <> "+" Then
                lDimsNew = lDimsNew + 1
                sSpec = Trim(Split(v, ":")(1))
                If InStr(sSpec, ",") = 0 Then
                    '~~ The dimension indicating digit is followed by a single integer value.
                    '~~ This is interpreted as the Ubound of the new dimension
                    '~~ while the LBound defaults to the Base Options which is either 0 or 1
                    ReDim aItem(0 To 1)
                    aItem(0) = lBase
                    aItem(1) = CLng(sSpec)
                Else
                    aItem = Split(sSpec, ",")
                End If
                For i = LBound(aItem) To UBound(aItem)
                    aItem(i) = Trim(aItem(i))
                Next i
                a_dims.Add lDimsNew, aItem
            End If
        Next v
    
        '~~ 3. Completed the dimensions with those in of the source Array's which not yet are collected
        For i = 1 To lDims
            If Not a_dims.Exists(i) Then
                On Error Resume Next
                a_dims.Add i + lDimsNew, Split(LBound(a_arr, i) & "," & UBound(a_arr, i), ",")
                If Err.Number <> 0 Then Exit For ' the previous i was the last dimension of the source array
            End If
        Next i
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Function ArryUnload(ByVal a_arr As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns an 2-dim array with all items of a multidimensional - max 8 - array
' (a_arr) with the indices delimited by a comma as the first Item and the
' array's Item as the second one. I.e. that a multidimendional array is unloaded
' into a "flat" array.
' ------------------------------------------------------------------------------
    Const PROC = "ArryUnload"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim cllDimSpecs As New Collection
    
    ReDim arr(1 To ArryItems(a_arr), 1 To 2)
    ArryDims a_arr, cllDimSpecs
    
    ArryUnloadAdd a_arr, arr, "", 1, cllDimSpecs.Count
    ArryUnload = arr
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Private Sub ArryUnloadAdd(ByVal a_arr As Variant, _
                          ByRef a_arr_out As Variant, _
                          ByRef a_indices As String, _
                          ByRef a_dim As Integer, _
                          ByVal a_dims As Integer)
' ------------------------------------------------------------------------------
' Adds an Item of an array (a_arr), addressed by indices - derived from the key
' (a_key) which is composed of them delimited by commas - to a Dictionary
' (a_dct) with the provided kex (a_key).
' ------------------------------------------------------------------------------
    Const PROC = "ArryUnloadAdd"
    
    On Error GoTo eh
    Dim i           As Long
    Dim sIndices    As String
    
    For i = LBound(a_arr, a_dim) To UBound(a_arr, a_dim)
        If a_indices = vbNullString _
        Then sIndices = i _
        Else sIndices = a_indices & "," & i
        
        If a_dim < a_dims Then
            ArryUnloadAdd a_arr, a_arr_out, sIndices, a_dim + 1, a_dims
        Else
            a_arr_out(i + 1, 1) = sIndices                   ' The Item's indices
            On Error Resume Next
            a_arr_out(i + 1, 2) = ArryItem(a_arr, sIndices) ' The Item
        End If
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub ArryUnloadToColl(ByRef a_arry As Variant, _
                             ByRef a_coll As Collection)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim i           As Long
    Dim cllInColl   As New Collection

    With a_coll
        If IsArray(a_arry) And (Not IsEmpty(a_arry)) Then
            If Not IsArray(a_arry(LBound(a_arry))) Then
                '~~ One dimensional array
                For i = LBound(a_arry) To UBound(a_arry)
                    .Add a_arry(i)
                Next i
            Else
                '~~ Multidimensional array unloaded by recursive calls
                For i = LBound(a_arry) To UBound(a_arry)
                    ArryUnloadToColl a_arry(i), cllInColl
                    .Add cllInColl
                Next i
            End If
        End If
    End With
    
End Sub

'' ----------------------------------------------------------------------------
'' Returns all items in a multidimensional array (a_arr) as Dictionary (a_dct)
'' with the elements as Item and the indices delimited by a comma as key.
'' ----------------------------------------------------------------------------
'    Dim i        As Long
'    Dim iCurrent As String
'
'    If a_dct Is Nothing Then Set a_dct = New Dictionary
'
'    If IsArray(a_arr) Then
'        For i = LBound(a_arr) To UBound(a_arr)
'            iCurrent = a_ndc & i
'            ' Recursively call for nested arrays
'            ArryUnloadToDict a_arr(i), a_dct, iCurrent & ","
'        Next i
'    Else ' is an Item
'        a_dct.Add Left(a_ndc, Len(a_ndc) - 1), a_arr
'    End If
'
'End Sub
'
Public Sub AssertedErrors(ParamArray c_app_errs() As Variant)
' ----------------------------------------------------------------------------
' When mErH is not installed (Cond. Comp. Arg. `mErh=1`) this is ignored.
' ----------------------------------------------------------------------------
    Const PROC = "AssertedErrors"

#If mErH Then
    On Error GoTo xt
    If ArryIsAllocated(c_app_errs) Then
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

Private Function BooleanAsStrg(ByVal b As Boolean) As String
    If b Then BooleanAsStrg = "TRUE" Else BooleanAsStrg = "FALSE"
End Function

Private Function checkHwnds(ByRef xlApps() As Application, hwnd As LongPtr) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "checkHwnds"
    
    On Error GoTo eh
    Dim i As Long
    
    If UBound(xlApps) = 0 Then GoTo xt

    For i = LBound(xlApps) To UBound(xlApps)
        If xlApps(i).hwnd = hwnd Then
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
' Removes/Deletes any Item collected in cllTempTestItems.
' ------------------------------------------------------------------------------
    Const PROC = "CleanUp"
    
    On Error GoTo eh
    Dim i   As Long
    Dim fle As File
    Dim v   As Variant
    
    With fso
        For Each fle In .GetFolder(sTestFolder).Files
            '~~ Remove any temporary test file
            If fle.Name Like "rad*" Then
                .DeleteFile fle.Path
            End If
            '~~ Remove when explicitly provided named
            If ArryIsAllocated(t_names) Then
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

                      
Private Function CollAsArry(ByVal v_items As Collection) As Variant
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
    CollAsArry = arr
    
End Function

Private Function CollAsDict(ByVal v_items As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each Item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each Item is returned as Item.'
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
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set CollAsDict = KeySort(dct)
    Set dct = Nothing

End Function

Private Function CollAsFile(ByVal c_items As Collection, _
                        Optional ByRef c_file As String = vbNullString, _
                        Optional ByVal c_file_append As Boolean = True) As File
' ----------------------------------------------------------------------------
' Transfers the items of a Collection (c_coll) as records/lines to a
' file (c_file_name), optionally appended (c_file_append).
' ----------------------------------------------------------------------------

    If c_file = vbNullString Then c_file = TempFileFullName
    StringAsFile CollAsStrg(c_items), c_file, c_file_append
    Set CollAsFile = fso.GetFile(c_file)

End Function

Private Function CollAsStrg(ByVal c_coll As Collection, _
                           Optional ByRef c_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an Item is an object its
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
    CollAsStrg = s

End Function

Private Function CollDiffs(ByVal v_res As Collection, _
                           ByVal v_exp As Collection, _
                  Optional ByRef v_res_prfx As String = vbNullString, _
                  Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the verification of a Collection (v_res) differs from
' another (v_exp).
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long

    If v_res.Count <> v_exp.Count Then
        sFailedNote = "Number of items Result/expected differs!"
        LogCollectVrfctnDetails vbNullString, vbNullString, VR_FAILED
    Else
        LogCollectVrfctnDetails vbNullString, vbNullString, vbNullString
    End If
    '~~ Result versus Result expected
    For i = 1 To Min(v_res.Count, v_exp.Count)
        If VarDiffs(v_res(i), v_exp(i), "Item(" & i & ") = ", v_exp_prfx) Then
            CollDiffs = True
        End If
    Next i
    
    '~~ results without any expected
    For j = i To v_res.Count
        sFailedNote = j & ". Item not expected!"
        LogCollectVrfctnDetails "Item(" & i & ") = " & v_res(j), vbNullString, VR_FAILED
        CollDiffs = True
    Next j
    
    '~~ Expected with a missing Result
    For j = i To v_exp.Count
        sFailedNote = j & ". Item missing!"
        LogCollectVrfctnDetails vbNullString, "Item(" & j & ") = " & v_exp(j), VR_FAILED
        CollDiffs = True
    Next j
    
End Function

Private Function DictAsArry(ByVal v_items As Dictionary) As Variant
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
    DictAsArry = arr
    
End Function

Private Function DictAsColl(ByVal v_items As Dictionary) As Collection
' ----------------------------------------------------------------------------
' Returns the Dictionary (v_items) with the keys in ascending order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In v_items
            .Add v
        Next v
    End With
    Set DictAsColl = cll
    Set cll = Nothing
    
End Function

Private Function DictAsDict(ByVal v_items As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns the Dictionary (v_items) with the keys in ascending order.
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    With dct
        For Each v In v_items
            .Add v, vbNullString
        Next v
    End With
    Set DictAsDict = KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function DictAsFile(ByVal v_items As Dictionary, _
                         Optional ByRef v_file_name As String = vbNullString, _
                         Optional ByVal v_file_append As Boolean = True) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    If v_file_name = vbNullString Then v_file_name = TempFileFullName
    StringAsFile DictAsStrg(v_items), v_file_name, v_file_append
    Set DictAsFile = fso.GetFile(v_file_name)

End Function

Private Function DictAsStrg(ByVal d_items As Dictionary, _
                           Optional ByRef d_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a Dictionary's keys as string with each key delimited by a vbCrLf.
' Keys are converted into a string, if an Item is an object its Name property
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
    DictAsStrg = s

End Function

Private Function DictDiffs(ByVal v_res As Dictionary, _
                           ByVal v_exp As Dictionary, _
                  Optional ByRef v_res_prfx As String = vbNullString, _
                  Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a Dictionary (v_res) differs from another (v_exp).
' ----------------------------------------------------------------------------
    Const PROC = "DictDiffs"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim dctDone As New Dictionary
    
    If v_res.Count <> v_exp.Count Then
        sFailedNote = "Number of items Result/expected differ!"
        LogCollectVrfctnDetails vbNullString, vbNullString, VR_FAILED
    Else
        LogCollectVrfctnDetails vbNullString, vbNullString, vbNullString
    End If
    
    For Each v In v_res
        If Not v_exp.Exists(v) Then
            sFailedNote = "Item not expected!"
            LogCollectVrfctnDetails VarAsStrg(v) & " = " & VarAsStrg(v_res(v)), vbNullString, VR_FAILED
            DictDiffs = True
        Else
            dctDone.Add v, vbNullString
            If VarDiffs(v_res(v), v_exp(v), VarAsStrg(v) & " = ", v_exp_prfx) Then
                DictDiffs = True
            End If
        End If
    Next v
    
    For Each v In v_exp
        If Not dctDone.Exists(v) Then
            sFailedNote = "Item with key """ & VarAsStrg(v) & """ missing!"
            LogCollectVrfctnDetails vbNullString, VarAsStrg(v) & " = " & VarAsStrg(v_exp(v)), VR_FAILED
            DictDiffs = True
        End If
    Next v
    
    Set dctDone = Nothing
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DsplyFile(ByVal d_file_name As String)
    ShellRun d_file_name
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsTestAid." & sProc
End Function

Private Function Exclude(ByVal e_strng As String, _
                         ByVal e_excld As String) As String
' ------------------------------------------------------------------------------
' Returns a string (e_string) with the "exclude from Result evaluation" string
' excluded.
' Examples: "????-??-?? " exclude this like pattern
'           "?=?"         exclude this like pattern
'
' Note: The string (e_strng) may contains substrings delimited by a vbCrLf and
'       for each of them the exclusion applies.
' ------------------------------------------------------------------------------
                         
    Dim a           As Variant
    Dim b           As Boolean
    Dim bEx         As Long     ' contains "?"or "#"
    Dim bExAtend    As Boolean  ' string ends with *
    Dim bExAtStart  As Boolean  ' string begins with *
    Dim bExBetween  As Boolean  ' string enclosed in *
    Dim i           As Long
    Dim lPos        As Long
    Dim s           As String
    Dim sEx         As String
    
    If e_excld = vbNullString Then
        Exclude = e_strng
        GoTo xt
    End If
    sEx = e_excld
    If InStr(sEx, "?") <> 0 Or InStr(sEx, "#") <> 0 Then bEx = True
    Select Case True
        Case Left(sEx, 1) = "*" And Right(sEx, 1) <> "*":   bExAtStart = True
                                                            sEx = Right(sEx, Len(sEx) - 1)
        Case Left(sEx, 1) <> "*" And Right(sEx, 1) = "*":   bExAtend = True
                                                            sEx = Left(sEx, Len(sEx) - 1)
        Case Left(sEx, 1) = "*" And Right(sEx, 1) = "*":    bExBetween = True
                                                            sEx = Mid(sEx, 2, Len(sEx) - 2)
        Case Else:                                          bExBetween = True
    End Select
    
    a = StringAsArry(e_strng, vbCrLf)
    For i = LBound(a) To UBound(a)
        s = a(i)
        If s Like "*" & sEx & "*" Then
            lPos = ExcludePosition(s, sEx, bEx)
            If lPos <> 0 Then
                b = True
                Select Case True
                    Case bEx And bExAtStart:    a(i) = Left(s, Len(s) - Len(sEx))
                    Case bEx And bExAtend:      a(i) = Right(s, Len(s) - Len(sEx))
                    Case Else:                  a(i) = Left(s, lPos - 1) & Right(s, Len(s) - Len(sEx) - lPos + 1)
                End Select
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
        If e_like Then
            If sPattern Like e_pattern Then
                ExcludePosition = i
                GoTo xt
            End If
        Else
            If sPattern = e_pattern Then
                ExcludePosition = i
                GoTo xt
            End If
        End If
    Next i

xt: Exit Function
End Function

Private Function ExcludeString(ByVal e_strng As String, _
                               ByVal e_excld As String, _
                               ByVal e_like As Boolean) As String
' ------------------------------------------------------------------------------
' Returns a string (e_excld_strng_from) with a substring beginning (e_excld_strng)
' which may contain ? excluded.
' ------------------------------------------------------------------------------
    
    Dim i       As Long
    Dim lStart  As Long
    
    '~~ Get start of the string
    For i = 1 To Len(e_strng) - Len(e_excld)
        If Not e_like Then
            If Mid(e_strng, i, Len(e_excld)) = e_excld Then
                lStart = i
                Exit For
            End If
        Else
            If Mid(e_strng, i, Len(e_excld)) Like e_excld Then
                lStart = i
                Exit For
            End If
        End If
    Next i
    Select Case lStart
        Case 0:                                 ExcludeString = e_strng
        Case 1:                                 ExcludeString = Right(e_strng, Len(e_strng) - Len(e_excld))
        Case Len(e_strng) - Len(e_excld) + 1:    ExcludeString = Left(e_strng, Len(e_strng) - Len(e_excld))
        Case Else:                              ExcludeString = Left(e_strng, lStart - 1) & Right(e_strng, (Len(e_excld) - lStart + 1))
    End Select

End Function

Private Function FileAsArry(ByVal f_file As Variant, _
                    Optional ByVal f_empty_excluded) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim sSplit  As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    For Each v In Split(FileAsStrg(f_file, sSplit), sSplit)
        If f_empty_excluded Then
            If Trim$(v) <> vbNullString Then Arry(arr) = v
        Else
            Arry(arr) = v
        End If
    Next v
    FileAsArry = arr
    
End Function

Private Function FileAsColl(ByVal v_items As File) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "FileAsColl"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim sSplit  As String
    
    With cll
        For Each v In Split(FileAsStrg(v_items.Path, sSplit), sSplit)
            .Add v
        Next v
    End With
    Set FileAsColl = cll
    Set cll = Nothing

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsDict(ByVal v_items As File) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the file's records (vulgo text lines) are returned as
'            Directory keys, lines will be unified. I e. line will exist only
'            once. To make this restriction productive, the number of
'            occurrences of each line is returned as Item.
' ----------------------------------------------------------------------------
    Dim dct     As New Dictionary
    Dim l       As Long
    Dim s       As String
    Dim sSplit  As String
    
    With dct
        For Each v In Split(FileAsStrg(v_items.Path, sSplit), sSplit)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set FileAsDict = KeySort(dct)
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
    
    With fso
        Select Case True
            Case f_rename And Not f_append:     If f_file_in.Path = .GetParentFolderName(f_file_out) _
                                                Then f_file_in.Name = .GetFileName(f_file_out) _
                                                Else Err.Raise AppErr(1), ErrSrc(PROC), "File cannot be renamed when the provided file and the new file's name " & _
                                                                                        "do not point to the same location!"
            Case Not f_rename And Not f_append: .CopyFile f_file_in.Path, f_file_out
            Case Not f_rename And f_append:     StringAsFile FileAsStrg(f_file_in, sSplit), f_file_out, True
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

Private Function FileAsStrg(Optional ByVal f_file_full_name As String, _
                              Optional ByRef f_split As String, _
                              Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    On Error Resume Next
    Dim s As String
    
    s = TxtFile(f_file_full_name)
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
        s = FileAsStrgEmptyExcluded(s)
    End If
    FileAsStrg = s
    
xt:
End Function

Private Function FileAsStrgEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    
    s_s = FileAsStrgTrimmed(s_s) ' leading and trailing empty already excluded
    Do While InStr(s_s, vbCrLf & vbCrLf) <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStrgEmptyExcluded = s_s
    
End Function

Private Function FileAsStrgTrimmed(ByVal s_s As String, _
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
    
    FileAsStrgTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Private Function FileDiffs(ByVal v_res As File, _
                           ByVal v_exp As File, _
                  Optional ByRef v_res_prfx As String = vbNullString, _
                  Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a file (v_res) differs form another (v_exp) by considering
' possibly exclusions from comparing lines.
' Precondition: Both files are test streams.
' ----------------------------------------------------------------------------
    Dim sResult         As String
    Dim sResultExpected As String
    

    sResult = Exclude(VarAsStrg(v_res), sExcludeFromComparison)
    sResultExpected = Exclude(VarAsStrg(v_exp), sExcludeFromComparison)
    If sResult = sResultExpected Then
        LogCollectVrfctnDetails "File " & v_res.Name, "File " & v_exp.Name, VR_PASSED
    Else
        '~~ Details verified only when failed
        LogCollectVrfctnDetails "File " & v_res.Name, "File " & v_exp.Name, VR_FAILED
        FileDiffs = ArryDiffs(VarAsArry(sResult), VarAsArry(sResultExpected), v_res_prfx, v_exp_prfx)
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
    TxtFile(f_target) = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)

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
    If fso.FolderExists(sTarget) Then fso.DeleteFolder sTarget
    fso.CreateFolder sTarget
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
    
    If Not fso.FolderExists(sSource) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A folder to zip """ & sSource & """ does not exist!"
    
    FldrZip sSource, sTarget

End Sub

Function GetAllExcelInstances() As Collection
    
    Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
    Dim hwnd        As LongPtr
    Dim hwndChild   As LongPtr
    Dim xlApp       As Object
    Dim className   As String * 255
    Dim retVal      As Long
    Dim xlApps      As New Collection

    hwnd = FindWindowEx(0, 0, "XLMAIN", vbNullString)

    Do While hwnd <> 0
        hwndChild = FindWindowEx(hwnd, 0, "XLDESK", vbNullString)
        hwndChild = FindWindowEx(hwndChild, 0, "EXCEL7", vbNullString)

        If hwndChild <> 0 Then
            retVal = GetClassName(hwndChild, className, 255)
            If Left(className, retVal) = "EXCEL7" Then
                AccessibleObjectFromWindow hwndChild, &HFFFFFFF0, IID_IDispatch, xlApp
                ' Add the xlApp object to the collection
                If Not xlApp Is Nothing Then
                    xlApps.Add xlApp
                End If
            End If
        End If

        hwnd = FindWindowEx(0, hwnd, "XLMAIN", vbNullString)
    Loop

    Set GetAllExcelInstances = xlApps
End Function

Function GetObjectCount() As Integer
    Dim obj    As Object
    Dim lCount As Long

    On Error Resume Next
    Do
        lCount = lCount + 1
        Set obj = GetObject(sLogBookFullName)
    Loop Until obj Is Nothing
    On Error GoTo 0

    GetObjectCount = lCount - 1

End Function

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

Private Function IsInteger(ByVal i_value As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is numeric, and has no decimals.
' ----------------------------------------------------------------------------
    If Not i_value = Empty Then
        If IsNumeric(i_value) _
        Then IsInteger = (i_value = Int(i_value))
    End If
End Function

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
         "is the target format. However, the current Item is an object which does " & _
         "not have a Name property!"
    
End Function

Function IsValidFileName(ByVal i_string As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the string (i_string) is valid as a file name or valid as
' part of a file name.
' ----------------------------------------------------------------------------
    
    Dim aInvalidChars   As Variant
    Dim aReservedNames  As Variant
    Dim sName           As String
    Dim v               As Variant
    
    IsValidFileName = True ' turned to False when invalid

    '~~ Check for invalid characters
    aInvalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each v In aInvalidChars
        If InStr(i_string, v) <> 0 Then
            IsValidFileName = False
            GoTo xt
        End If
    Next v
    
    '~~ Check for reserved names
    aReservedNames = Array("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
    sName = UCase(i_string)
    For Each v In aReservedNames
        If InStr(sName, v) <> 0 Then
            IsValidFileName = False
            GoTo xt
        End If
    Next v
    
xt: Exit Function
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
    With dct
        For i = LBound(arr) To UBound(arr)
            vKey = arr(i)
            .Add key:=vKey, item:=k_dct.item(vKey)
        Next i
    End With
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub LogCollectVrfctnDetails(ByVal v_res As String, _
                                    ByVal v_exp As String, _
                                    ByVal v_ver As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
                                      
           LogItemTitle(aLogDetails) = sTitle
           LogItemTestId(aLogDetails) = "'" & sTestId
        LogItemVrfctnNo(aLogDetails) = "'" & Format(lVrfctnNumber, "00")
      LogItemVrfctnDesc(aLogDetails) = sVerification
       LogItemVrfydType(aLogDetails) = sVrfctnType
    LogItemVrfctnResult(aLogDetails) = v_ver
       LogItemVrfctnRes(aLogDetails) = v_res
    LogItemVrfctnResExp(aLogDetails) = v_exp
    If dTimerMsecs = 0 _
    Then LogItemVxcTime(aLogDetails) = "n/a" _
    Else LogItemVxcTime(aLogDetails) = "'" & Format(dTimerMsecs, "#,##0.000")
        LogItemComment(aLogDetails) = sComment
    If sFailedNote <> vbNullString _
    Then LogItemComment(aLogDetails) = sFailedNote
    
    If cllLogDetails Is Nothing Then Set cllLogDetails = New Collection
    cllLogDetails.Add aLogDetails
    
End Sub

Private Sub LogInitApp()
' ----------------------------------------------------------------------------
' Create a new instance of Excel. In case this component runs in a non-Excel
' application one is created, else this excel application is used.
' ----------------------------------------------------------------------------
    Const PROC = "LogInitApp"
    
    On Error GoTo eh
    LogWbkClose ' in case open
    
    If ThisWorkbook.Parent.Name Like "*Excel" Then
        Set appLog = ThisWorkbook.Parent
    Else
        '~~ establish a dedicated Excel Application instance
        Set appLog = CreateObject("Excel.Application")
        With appLog
'            .Visible = False
            .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
            .DisplayFormulaBar = False
            .DisplayStatusBar = False
            .DisplayScrollBars = True
            .WindowState = xlNormal
        End With
    End If
    Set wbkLog = Nothing
    Set wshLog = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogInitWbk()
' ------------------------------------------------------------------------------
' Close any open, delete it, and re-establish it
' -----------------------------------------------------------------------------
        
    '~~ Get rid of any open and/or existing log book
    On Error Resume Next
    Set wbkLog = appLog.Workbooks(sLogBookName)
    On Error GoTo 0
    If Not wbkLog Is Nothing Then wbkLog.Close False
    If fso.FileExists(sLogBookFullName) Then fso.DeleteFile sLogBookFullName
    
    Set wbkLog = appLog.Workbooks.Add
    wbkLog.SaveAs sLogBookFullName
'    Workbooks(sLogBookName).Windows(1).Visible = False
    Set wshLog = Nothing
    
xt: Exit Sub

End Sub

Private Sub LogSheetBorder(ByVal l_border As Border)
    
    With l_border
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With

End Sub

Private Sub LogSheetFinalize()
' ----------------------------------------------------------------------------
' Format, merge, etc.
' Column specs: <col>, <width>, <linebreak>
' ----------------------------------------------------------------------------
    Const PROC = "LogSheetFinalize"
    
    On Error GoTo eh
        
    If appLog.Workbooks.Count = 1 Then
        With appLog
            Stop
            .Left = 10 ' Set the horizontal position (in points)
            .Top = 10 ' Set the vertical position (in points)
'            .Visible = True
        End With
    End If
    
    wbkLog.Activate
    wshLog.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    
    Application.EnableEvents = False
    wbkLog.Save
    Application.EnableEvents = True
    wbkLog.Windows(1).Visible = True
    
    appLog.ScreenUpdating = True
    appLog.Visible = True ' in case it had been established explicitely
    wshLog.Visible = xlSheetVisible

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogSheetFinalizeAdjustColumns(ParamArray l_spec() As Variant)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim col   As Range
    Dim lCols As Long
    Dim cll   As Collection
    Dim dct   As New Dictionary
    Dim i     As Long
    Dim lCol  As Long
    
    If ArryIsAllocated(l_spec) Then
        '~~ Obtain finalization column specs
        For i = LBound(l_spec) To UBound(l_spec) Step 3
            Set cll = New Collection
            lCol = l_spec(i)
            cll.Add l_spec(i + 1)
            cll.Add l_spec(i + 2)
            dct.Add lCol, cll
            Set cll = Nothing
        Next i
    End If
    
    If dct.Count <> 0 Then
        With wshLog
            For Each col In .UsedRange.Columns
                If dct.Exists(col.Column) Then
                    col.EntireColumn.ColumnWidth = dct(col.Column)(1)
                    col.WrapText = dct(col.Column)(2)
                    lCols = lCols + 1
                End If
                If lCols = dct.Count Then Exit For
            Next col
        End With
    End If

End Sub

Private Sub LogSheetFinalizeFormatings()
    
    Dim cel As Range
    Dim rng As Range
    
    With wshLog
        For Each cel In .UsedRange
            If cel.value <> vbNullString Then
                cel.Errors(xlNumberAsText).Ignore = True
            End If
        Next cel
        With .UsedRange
            .Font.Size = 10
            .Columns.AutoFit
            For Each rng In .Columns
                rng.EntireColumn.ColumnWidth = rng.EntireColumn.ColumnWidth + 5
            Next rng
            .Columns.AutoFit
            .Font.Size = 9
        End With
        With .UsedRange
            .VerticalAlignment = xlCenter
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            LogSheetBorder .Borders(xlEdgeLeft)
            LogSheetBorder .Borders(xlEdgeTop)
            LogSheetBorder .Borders(xlEdgeBottom)
            LogSheetBorder .Borders(xlEdgeRight)
            LogSheetBorder .Borders(xlInsideVertical)
            LogSheetBorder .Borders(xlInsideHorizontal)
        End With
        Range("A1").EntireColumn.ColumnWidth = 2.86
    End With

End Sub

Private Sub LogSheetFinalizeMergeColsEqual(ParamArray l_cols() As Variant)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "LogSheetFinalizeMergeColsEqual"
    
    On Error GoTo eh
    Dim sValue      As Variant
    Dim rngMerge    As Range
    Dim cel         As Range
    Dim col         As Range
    Dim xlHAlgnmnt  As Long
    Dim aCols       As Variant
    
    If Not ArryIsAllocated(l_cols) Then Exit Sub
    For Each v In l_cols
        Arry(aCols, v) = True
    Next v
    
    For Each col In wshLog.UsedRange.Columns
        If col.Column = 1 Then GoTo nc
        If Arry(aCols, col.Column - 1) Then
            For Each cel In col.Cells
                If cel.row <= rngLogSheetHeader.Rows.Count Then GoTo nc ' skip header rows
                If cel.MergeCells Then GoTo nc                 ' skip any already merged
                Select Case True
                    Case rngMerge Is Nothing
                        sValue = cel.value
                        Set rngMerge = cel
                        xlHAlgnmnt = cel.HorizontalAlignment
                    Case cel.value = sValue
                        Set rngMerge = appLog.Union(rngMerge, cel)
                    Case cel.value <> sValue
                        If Not rngMerge Is Nothing Then
                            If rngMerge.Rows.Count > 1 Then
                                With rngMerge
                                    Application.DisplayAlerts = False
                                    .Merge
                                    Application.DisplayAlerts = False
                                    .HorizontalAlignment = xlHAlgnmnt
                                    .VerticalAlignment = xlCenter
                                End With
                            End If
                        End If
                        sValue = cel.value
                        Set rngMerge = cel
                        xlHAlgnmnt = cel.HorizontalAlignment
                End Select
nc:         Next cel
            If Not rngMerge Is Nothing Then
                If rngMerge.Rows.Count > 1 Then
                    With rngMerge
                        Application.DisplayAlerts = False
                        .Merge
                        Application.DisplayAlerts = False
                        .HorizontalAlignment = xlHAlgnmnt
                        .VerticalAlignment = xlCenter
                   End With
                End If
            End If
            Set rngMerge = Nothing
        End If
    Next col
    
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogSheetHeader(ByVal l_title As String, _
                      ParamArray l_header() As Variant)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "LogSheetHeader"
    
    On Error GoTo eh
    Dim lCols     As Long
    Dim celFreeze As Range
    
    Application.ScreenUpdating = False
    lCols = UBound(l_header) + 1
    
    With wshLog
        If l_title <> vbNullString Then
            .Range("B2").value = l_title
            .Range(.Cells(2, 2), .Cells(2, lCols + 1)).Merge
            ArryAsRnge l_header, .Range("B3")
            Set rngLogSheetHeader = .Range(.Cells(2, 2), .Cells(3, lCols + 1))
            Set celFreeze = .Range(.Cells(4, 2), .Cells(4, 2))
        Else
            ArryAsRnge l_header, .Range("B2")
            Set rngLogSheetHeader = .Range(.Cells(2, 2), .Cells(2, lCols + 1))
            Set celFreeze = .Range(.Cells(3, 2))
        End If
    End With
    
    With rngLogSheetHeader
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    ' Freeze the used range = the header
    wbkLog.Activate
    wshLog.Activate
    
    celFreeze.Select
    ActiveWindow.FreezePanes = True
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogSheetInitialize(ByVal l_name As String, _
                               ByVal l_at_end As Boolean, _
                          ParamArray l_format() As Variant)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "LogSheetInitialize"
    
    On Error GoTo eh
    Dim i       As Long
    Dim lCol    As Long
    Dim lFormat As Long
    Dim lIndent As Long
    Dim rngCol  As Range
    Dim wsh     As Worksheet
        
    If Not ArryIsAllocated(l_format) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No columns specification (alignment and indentaion) provided!"
    
    On Error Resume Next
    If appLog Is Nothing Then LogInitApp
    If wbkLog Is Nothing Then LogInitWbk
    Set wshLog = Nothing
    Set wshLog = wbkLog.Worksheets(l_name)
    If Not wshLog Is Nothing Then
        wshLog.UsedRange.Clear
    Else
        If l_at_end _
        Then Set wshLog = wbkLog.Worksheets.Add(After:=wbkLog.Worksheets(wbkLog.Worksheets.Count)) _
        Else Set wshLog = wbkLog.Worksheets.Add(Before:=wbkLog.Worksheets(1))
        wshLog.Name = l_name
    End If
    
    lCol = 1
    For i = LBound(l_format) To UBound(l_format) Step 2
        lCol = lCol + 1
        lIndent = l_format(i)
        lFormat = l_format(i + 1)
        Set rngCol = wshLog.Cells(, lCol)
        With rngCol.EntireColumn
            .HorizontalAlignment = lFormat ' e.g. xlLeft
            .IndentLevel = lIndent
        End With
    Next i
    
    '~~ Remove all non "Test-* Worksheets
    For Each wsh In wbkLog.Worksheets
        If Not wsh.Name Like "Test-*" Then
            wsh.Visible = xlSheetHidden
        End If
    Next wsh
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogSheetInitializeDetails()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "LogSheetInitializeDetails"
    
    On Error GoTo eh
    Dim s As String
    
    If appLog Is Nothing Then LogInitApp
    On Error Resume Next
    s = wbkLog.Name
    If s = vbNullString Then LogInitWbk
    
    aLogDetails = cllLogDetails(1)
    
    If wshLog Is Nothing Then
        LogSheetInitialize "Test-" & sTestId, True _
                 , 0, xlCenter _
                 , 1, xlLeft _
                 , 0, xlCenter _
                 , 1, xlLeft _
                 , 1, xlLeft _
                 , 0, xlCenter _
                 , 0, xlRight _
                 , 1, xlLeft
    End If
                                                             
    LogSheetHeader LogItemTestId(aLogDetails) & ": " & LogItemTitle(aLogDetails) _
                 , LogItemVrfctnNo(aLogHeader) _
                 , LogItemVrfctnDesc(aLogHeader) _
                 , LogItemVrfydType(aLogHeader) _
                 , LogItemVrfctnRes(aLogHeader) _
                 , LogItemVrfctnResExp(aLogHeader) _
                 , LogItemVrfctnResult(aLogHeader) _
                 , LogItemVxcTime(aLogHeader) _
                 , LogItemComment(aLogHeader)
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogSheetInitializeSummary()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim s As String
    
    If appLog Is Nothing Then LogInitApp
    On Error Resume Next
    s = wbkLog.Name
    If s = vbNullString Then LogInitWbk
    
    LogSheetInitialize "Test-Summary", False _
               , 0, xlCenter _
               , 1, xlLeft _
               , 0, xlCenter _
               , 1, xlLeft _
               , 1, xlLeft _
               , 1, xlLeft _
               , 0, xlCenter _
               , 1, xlLeft _
               , 1, xlRight _
               , 1, xlLeft
    
    LogSheetHeader "Regression Test Summary" _
                 , LogItemTestId(aLogHeader) _
                 , LogItemTitle(aLogHeader) _
                 , LogItemVrfctnNo(aLogHeader) _
                 , LogItemVrfctnDesc(aLogHeader) _
                 , "Proc-Type" _
                 , LogItemCompName(aLogHeader) _
                 , LogItemProcName(aLogHeader) _
                 , LogItemVrfctnResult(aLogHeader) _
                 , LogItemVxcTime(aLogHeader) _
                 , LogItemComment(aLogHeader)
    
End Sub

Public Sub LogSheetWrite(ParamArray l_log() As Variant)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim lRowNew As Long
    
    With wshLog
        lRowNew = .UsedRange.Rows.Count + 1
        ArryAsRnge l_log, .Cells(lRowNew + 1, 2)
    End With

End Sub

Private Sub LogTerminate()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "LogTerminate"

    On Error GoTo eh
    Dim wbk As Workbook
    
    If appLog Is Nothing Then GoTo xt
    For Each wbk In appLog.Workbooks
        If wbk.FullName = sLogBookFullName Then
            wbkLog.Close True
            Set wbkLog = Nothing
            Set wshLog = Nothing
            Exit For
        End If
    Next wbk
    
    If appLog.Workbooks.Count = 0 Then appLog.Quit

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogWbkClose()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim obj As Object
    Dim wbk As Workbook
    
    On Error Resume Next
    Set obj = GetObject(sLogBookFullName)
    If Not obj Is Nothing Then
        Set appLog = obj.Parent
        For Each wbk In appLog.Workbooks
            If wbk.FullName = sLogBookFullName Then
                wbk.Close True
                Exit For
            End If
        Next wbk
        If appLog.Workbooks.Count = 0 _
        Then appLog.Quit
    
    End If

End Sub

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

Private Function ObjectAsStrg(ByVal o As Object) As String
    
    On Error Resume Next
    ObjectAsStrg = o.Name
    If Err.Number <> 0 Then ObjectAsStrg = "Object_" & o.HashCode

End Function

Private Sub ResultLogDetails()
' ------------------------------------------------------------------------------
' Displays the Result of a test versus its expected Result. The Result is
' presented in three columns: Passed/Failed Result Result Expected.
' Called by ResultVersusExpected.
' ------------------------------------------------------------------------------
    Const PROC = "ResultLogDetails"
    
    On Error GoTo eh
        
    If cllLogDetails Is Nothing Then GoTo xt
    If cllLogDetails.Count = 0 Then GoTo xt
    
    Application.ScreenUpdating = False
    If wshLog Is Nothing Then LogSheetInitializeDetails
    
    For Each v In cllLogDetails
        aLogDetails = v
        LogSheetWrite LogItemVrfctnNo(aLogDetails) _
                    , LogItemVrfctnDesc(aLogDetails) _
                    , LogItemVrfydType(aLogDetails) _
                    , LogItemVrfctnRes(aLogDetails) _
                    , LogItemVrfctnResExp(aLogDetails) _
                    , LogItemVrfctnResult(aLogDetails) _
                    , LogItemVxcTime(aLogDetails) _
                    , LogItemComment(aLogDetails)
    Next v
    Set cllLogDetails = Nothing
    
    LogSheetFinalizeFormatings
    LogSheetFinalizeMergeColsEqual 1, 2
    LogSheetFinalize
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ResultLogSummary()
' ----------------------------------------------------------------------------
' Writes the collected verification summary to the log Worksheet.
' ----------------------------------------------------------------------------
    Const PROC = "ResultLogSummary"
    
    On Error GoTo eh
    Dim v As Variant
    
    If cllVrfctnSummary.Count = 0 Then Exit Sub
    LogSheetInitializeSummary
    
    For Each v In cllVrfctnSummary
        aLogSummary = v
        LogSheetWrite LogItemTestId(aLogSummary) _
                    , LogItemTitle(aLogSummary) _
                    , LogItemVrfctnNo(aLogSummary) _
                    , LogItemVrfctnDesc(aLogSummary) _
                    , LogItemProcType(aLogSummary) _
                    , LogItemCompName(aLogSummary) _
                    , LogItemProcName(aLogSummary) _
                    , LogItemVrfctnResult(aLogSummary) _
                    , LogItemVxcTime(aLogSummary) _
                    , LogItemComment(aLogSummary) _

    Next v
    LogSheetFinalizeFormatings
    LogSheetFinalizeMergeColsEqual 1, 2, 10
    LogSheetFinalizeAdjustColumns 3, 40, True
    LogSheetFinalize
    Set cllVrfctnSummary = Nothing
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ResultVersusExpected(ByVal v_res As Variant, _
                                 ByVal v_exp As Variant)
' ----------------------------------------------------------------------------
' When both, a .Result (v_res) and a .ResultExpected (v_exp) had been provided
' each element of the two (file lines, cokllectio items, range values, etc.
' are compared (VrfyVarDiffs) which provides each element's Result in a triple
' in a collection (Result, expected, passed/failed). When the Result of the
' verification is "failed" a corresponding failed-log-file is displayed.
' In regression-test mode the summary Result is collected for being finally
' displayed.
' ----------------------------------------------------------------------------
    Const PROC = "ResultVersusExpected"
    
    Dim sVrfctnSummaryResult    As String
    
    Select Case True
        Case IsArray(v_res) And IsArray(v_exp)
        Case TypeName(v_res) = TypeName(v_exp)
        Case Else: Err.Raise AppErr(1), ErrSrc(PROC), _
                             "The Result and the ResultExpected are different types of objects but need to be the same!" & vbCrLf & _
                             "Result = " & TypeName(v_res) & ", ResultExpected = " & TypeName(v_exp)
    End Select
    
    If VarDiffs(v_res, v_exp) Then
        sVrfctnSummaryResult = VR_FAILED
        ResultLogDetails
        sVrfctnSummaryComment = "See log sheet " & wshLog.Name & " for details"
    Else
        sVrfctnSummaryResult = VR_PASSED
        sVrfctnSummaryComment = vbNullString
    End If
            
    '~~ Verification summary, writen to the log Worksheet by an explicit call of ResultLogSummary
    '~~ usually in the Regression-Test procecure"
          LogItemTestId(aLogSummary) = "'" & sTestId
           LogItemTitle(aLogSummary) = sTitle
        LogItemVrfctnNo(aLogSummary) = "'" & Format(lVrfctnNumber, "00")
      LogItemVrfctnDesc(aLogSummary) = sVerification
        LogItemProcType(aLogSummary) = sTestedProcType
        LogItemCompName(aLogSummary) = sTestedComp
        LogItemProcName(aLogSummary) = sTestedProc
    LogItemVrfctnResult(aLogSummary) = sVrfctnSummaryResult
         LogItemComment(aLogSummary) = sComment:    sComment = vbNullString
    If sFailedNote <> vbNullString _
    Then LogItemComment(aLogSummary) = sFailedNote
    sFailedNote = vbNullString
    
    If dTimerMsecs = 0 _
    Then LogItemVxcTime(aLogSummary) = "n/a" _
    Else LogItemVxcTime(aLogSummary) = "'" & Format(dTimerMsecs, "#,##0.000")
    If cllVrfctnSummary Is Nothing Then Set cllVrfctnSummary = New Collection
    cllVrfctnSummary.Add aLogSummary
    
xt: dTimerMsecs = 0
    sVerification = vbNullString ' Another pair of .Result/.ResultExpected requires a new specified Verification.
    bResult = False
    bResultExpected = False
    Exit Sub

End Sub

Private Function RngeAddress(ByVal r_rng As Range) As String
    RngeAddress = Replace(r_rng.Address, "$", vbNullString)
End Function

Private Function RngeAsDict(ByVal r_rng As Range, _
                          Optional ByRef r_addr As Dictionary = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all cells in a range (r_rng) with the row and the
' column as key ("RnCn").
' ----------------------------------------------------------------------------

    Dim cel     As Range
    Dim dctAddr As New Dictionary
    Dim dctRng  As New Dictionary
    Dim lCol    As Long
    Dim lRow    As Long
    
    '~~ Loop through each cell in the ranges
    For lCol = 1 To r_rng.Columns.Count
        For lRow = 1 To r_rng.Rows.Count
            Set cel = r_rng.Cells(lRow, lCol)
            dctRng.Add "R" & lRow & "C" & lCol, cel.value
            dctAddr.Add "R" & lRow & "C" & lCol, RngeAddress(cel)
        Next lRow
    Next lCol
    Set RngeAsDict = dctRng
    If Not r_addr Is Nothing Then Set r_addr = dctAddr
    Set dctRng = Nothing
    Set dctAddr = Nothing
    
End Function

Private Function RngeDiffs(ByVal v_res As Range, _
                           ByVal v_exp As Range) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a range (v_res) differs from another (v_exp).
' ----------------------------------------------------------------------------
    Dim dctRes      As Dictionary
    Dim dctAddr     As New Dictionary
    Dim dctExp      As Dictionary
    Dim dctDone     As New Dictionary
    
    If v_res.Columns.Count <> v_exp.Columns.Count _
    Or v_res.Rows.Count <> v_exp.Rows.Count Then
        sFailedNote = "Number of columns/rows Result/expected differ!"
        LogCollectVrfctnDetails "Range " & Replace(v_res.Address, "$", vbNullString), "Range " & Replace(v_exp.Address, "$", vbNullString), VR_FAILED
    Else
        LogCollectVrfctnDetails "Range " & Replace(v_res.Address, "$", vbNullString), "Range " & Replace(v_exp.Address, "$", vbNullString), vbNullString
    End If
    Set dctRes = RngeAsDict(v_res, dctAddr)
    Set dctExp = RngeAsDict(v_exp)
    
    For Each v In dctRes
        If Not dctExp.Exists(v) Then
            sFailedNote = "Cell not expected!"
            LogCollectVrfctnDetails v & "=" & dctRes(v), vbNullString, VR_FAILED
            RngeDiffs = True
        Else
            dctDone.Add v, vbNullString
            If StrgDiffs(CStr(dctRes(v)), CStr(dctExp(v)), v & " = ", v & " = ") Then
                RngeDiffs = True
            End If
        End If
    Next v
    
    For Each v In dctExp
        If Not dctDone.Exists(v) Then
            sFailedNote = "Cell missing in Result!"
            LogCollectVrfctnDetails vbNullString, v & "=" & dctExp(v), VR_FAILED
            RngeDiffs = True
        End If
    Next v
    
xt: Exit Function

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

Public Sub SleepMsecs(ByVal s_msecs As Long)
    Sleep s_msecs
End Sub

Private Function SplitIndctr(ByVal s_strng As String, _
                    Optional ByRef s_indctr As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the split indicator (s_indctr) of a string (s_strng) as string and
' as argument. The dedection is bypassed in case one (s_indctr) has already
' been provided.
' Note: By intention, the list of possibly identified split indicatory is kept
'       to a minimum in order not to interfere with numeric values as strings,
'       list, sentenses, etc.
' ----------------------------------------------------------------------------
    If s_indctr = vbNullString Then
        Select Case True
            Case InStr(s_strng, vbCrLf) <> 0: s_indctr = vbCrLf
            Case InStr(s_strng, vbLf) <> 0:   s_indctr = vbLf      ' e.g. in case of a downloaded file's_strng complete string
            Case InStr(s_strng, "|&|") <> 0:  s_indctr = "|&|"
        End Select
    End If
    SplitIndctr = s_indctr

End Function

Private Function StrgDiffs(ByVal v_res As String, _
                           ByVal v_exp As String, _
                  Optional ByRef v_res_prfx As String = vbNullString, _
                  Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a strings (v_res) differs from another (v_exp).
' ----------------------------------------------------------------------------
    
    If v_res <> v_exp Then
        StrgDiffs = True
        LogCollectVrfctnDetails v_res_prfx & v_res, v_exp_prfx & v_exp, VR_FAILED
    Else
        LogCollectVrfctnDetails v_res_prfx & v_res, v_exp_prfx & v_exp, VR_PASSED
    End If

End Function

Private Function StringAsArry(ByVal s_strng As String, _
                      Optional ByVal s_split As String = vbNullString, _
                      Optional ByVal s_trim As Variant = True) As Variant
' ----------------------------------------------------------------------------
' Returns a string (s_strng) split into an array of strings. When no split
' indicator (s_split) is provided it one is found by examination of the
' string (s_strng). When the option (s_trim) is TRUE (the default), "R", or
' "L" the items in the array are returned trimmed accordingly.
' Example 1: arr = StringAsArry("this is a string", " ") is returned as an
'            array with 3 items: "this", "is", "a", "string".
' Example 2: arr = StringAsArry(FileAsStrg(FileName),sSplit,False) is
'            returned as any array with records/lines of the provided file,
'            whereby the lines are not trimmed, i.e. leading spaces are
'            preserved.
'            Note: The not provided split indicator has the advantage that it
'                  is provided by the SplitIndctr service, which in that case
'                  returns either vbCrLf or vbLf, the latter when the file is
'                  a download.
' Example 3: arr = FileAsArry(<file>) return the same as example 2.
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
    StringAsArry = arr
    
End Function

Public Function StringAsColl(ByVal s_items As String, _
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
    Set StringAsColl = cll
    Set cll = Nothing
    
End Function

Private Function StringAsDict(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as Item.
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
    Set StringAsDict = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function StringAsFile(ByVal s_strng As String, _
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
    
    TxtFile(s_file, s_file_append) = s_strng
    Set StringAsFile = fso.GetFile(s_file)
    
End Function

Private Function StringAsStrg(ByVal s_items As String, _
                       Optional ByVal s_delimiter As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns a string (v_Item) with any delimiter delimited by vbCrLf by default.
' ----------------------------------------------------------------------------
    Const PROC = "StringAsStrg"
    
    On Error GoTo eh
    Dim sSplit As String
    sSplit = SplitIndctr(s_items)

    StringAsStrg = Replace(s_items, sSplit, s_delimiter)
    
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
    
    s = Replace(fso.GetTempName, ".tmp", f_ext)
    If f_path = vbNullString Then f_path = CurDir
    s = VBA.Replace(f_path & "\" & s, "\\", "\")
    If f_create_as_textstream Then fso.CreateTextFile s
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
    
    If fso.FileExists(TestFileFullName(t_name, t_ext)) _
    Then Set TestFile = fso.GetFile(TestFileFullName) _
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
    TestFileFullName = sTestFolder & "\" & TestFileName(sTestId, Format(lVrfctnNumber, "00"), t_name, t_ext)

End Function

Private Function TestFileName(ByVal t_id As String, _
                              ByVal t_ver_no As Long, _
                     Optional ByVal t_name As String = vbNullString, _
                     Optional ByVal t_ext As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a test file's name based on the specified scheme which defaults to
' "Test-<id>-<vn>-<nm>.<xt>".
' ----------------------------------------------------------------------------
    Dim s As String
    
    s = sTestFileNameScheme
    s = Replace(s, SCHEME_ID, t_id)
    
    If t_ver_no = 0 Then t_ver_no = 1
    s = Replace(s, SCHEME_VER_NO, t_ver_no)
    
    If t_name <> vbNullString _
    Then s = Replace(s, SCHEME_NAME, t_name) _
    Else s = Replace(s, "-" & SCHEME_NAME & "-", vbNullString)
    If t_ext = vbNullString Then t_ext = TestFileExtension
    s = Replace(s, SCHEME_EXT, t_ext)
    s = Replace(Replace(s, "--", "-"), "..", ".")
    TestFileName = s
    
End Function

Private Sub TestFileNames()
    sNameTestFile = sTestFolder & "\" & TestFileName(sTestId, Format(lVrfctnNumber, "00"), , sTestFileExtension)
    sNameTestResultFile = sTestFolder & "\" & TestFileName(sTestId, Format(lVrfctnNumber, "00"), "Result", sTestFileExtension)
    sNameTestResultExpectedFile = sTestFolder & "\" & TestFileName(sTestId, Format(lVrfctnNumber, "00"), "Result-Expected", sTestFileExtension)
End Sub

Public Function TestResultExpectedFile() As File
' ----------------------------------------------------------------------------
' Returns a file named <test-folder>/.
' ----------------------------------------------------------------------------
    Const PROC = "TestResultExpectedFile"
    
    With fso
        If .FileExists(sNameTestResultExpectedFile) _
        Then Set TestResultExpectedFile = .GetFile(sNameTestResultExpectedFile) _
        Else Err.Raise AppErr(1), ErrSrc(PROC), "A test Result file named """ & sNameTestResultExpectedFile & """ does not exist!"
    End With
End Function

Public Function TestResultFile() As File
' ------------------------------------------------------------------------------
' Returns the default name of a test Result file in the form:
' <tes-folder>\TestResult-<test-number><file-extention>
' when a file with the returned name already exists, it is deleted
' ------------------------------------------------------------------------------
    Const PROC = "TestResultFile"
    
    With fso
        If .FileExists(sNameTestResultFile) _
        Then Set TestResultFile = .GetFile(sNameTestResultFile) _
        Else Err.Raise AppErr(1), ErrSrc(PROC), "A test Result file named """ & sNameTestResultFile & """ does not exist!"
    End With
    
End Function

Public Sub TimerCalibration(Optional ByVal t_sleep_msec As Long = 20, _
                            Optional ByVal t_sleeps As Long = 1, _
                            Optional ByVal t_loops As Long = 1)
' ----------------------------------------------------------------------------
' Determines the timer's overhead caused by the measurement itself.
' ----------------------------------------------------------------------------
    Dim c As Currency
    Dim i As Long, j As Long
    
    QueryPerformanceFrequency cTimerFrequency
    For j = 1 To t_loops
        For i = 1 To t_sleeps
            TimerStart False
            Sleep t_sleep_msec
            TimerEnd
            c = c + (cTicksElapsed - t_sleep_msec)
        Next i
    Next j
    cTicksOverhead = c / (t_sleeps * t_loops)
    If cTicksOverhead < 0 Then
        cTicksOverhead = 0
        dTimerOverhead = 0
    Else
        dTimerOverhead = (cTicksOverhead / cTimerFrequency) * 1000 ' milliseconds
    End If
    
End Sub

Public Function TimerEnd(Optional t_msecs As Double) As Double
' ----------------------------------------------------------------------------
' Returns the elapsed time (t_elpsd) since start in milliseconds reduced by
' the timer overhead - which returns the most possible precise exec time.
' This overhead is the elapsed time between TimerStart and TimerEnd when
' nothing is measured which calculated when the class is initialized.
' ----------------------------------------------------------------------------
    Const PROC = "TimerEnd"
    
    Dim cStart As Currency
    Dim cOvrhd As Currency
    
    With cllTimerStack
        If .Count = 0 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "TimerEnd called without a corresponding TimerStart!"
        cStart = .item(.Count)  ' pop ticks at timer start
        .Remove .Count
    End With
    With cllTicksStackOvrhd
        cOvrhd = .item(.Count) ' pop ticks overhead at timer start
        .Remove .Count
    End With
    
    QueryPerformanceCounter cTicksEnd
    cTicksElapsed = cTicksEnd - cStart - cOvrhd
    dTimerMsecs = (cTicksElapsed / cTimerFrequency) * 1000 ' milliseconds
    If dTimerMsecs < 0 Then dTimerMsecs = 0.001
    t_msecs = dTimerMsecs
    TimerEnd = dTimerMsecs
    
End Function

Public Sub TimerStart(Optional ByVal t_with_calibration = True)
    
    If t_with_calibration Then TimerCalibration
    QueryPerformanceCounter cTicksStart
    cllTimerStack.Add cTicksStart ' queue
    cllTicksStackOvrhd.Add cTicksOverhead
    
End Sub

Private Function VarAsArry(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsArry"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsArry = CollAsArry(v_items)
        Case TypeName(v_items) = "File":        VarAsArry = FileAsArry(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsArry = DictAsArry(v_items)
        Case TypeName(v_items) Like "*()":      VarAsArry = v_items
        Case VarType(v_items) = vbArray:        VarAsArry = v_items
        Case IsArray(v_items):                  VarAsArry = v_items
        Case VarType(v_items) = vbString:       VarAsArry = StringAsArry(v_items)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Private Function VarAsColl(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsColl = v_items
        Case TypeName(v_items) = "Dictionary":  Set VarAsColl = DictAsColl(v_items)
        Case TypeName(v_items) = "File":        Set VarAsColl = FileAsColl(v_items)
        Case TypeName(v_items) Like "*()":      Set VarAsColl = ArryAsColl(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsColl = ArryAsColl(v_items)
        Case VarType(v_items) = vbBoolean:      cll.Add BooleanAsStrg(v_items)
                                                Set VarAsColl = cll
        Case VarType(v_items) = vbString:       Set VarAsColl = StringAsColl(v_items)
        Case Else:                              cll.Add v_items
                                                Set VarAsColl = cll

    End Select
    
End Function

Private Function VarAsDict(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsDict = CollAsDict(v_items)
        Case TypeName(v_items) = "File":        Set VarAsDict = FileAsDict(v_items)
        Case TypeName(v_items) = "Dictionary":  Set VarAsDict = DictAsDict(v_items) ' sort key ascending
        Case TypeName(v_items) Like "*()":      Set VarAsDict = ArryAsDict(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsDict = ArryAsDict(v_items)
        Case VarType(v_items) = vbString:       Set VarAsDict = StringAsDict(v_items)
        Case VarType(v_items) = vbBoolean:      dct.Add BooleanAsStrg(v_items), vbNullString
                                                Set VarAsDict = dct
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
        Case TypeName(v_items) = "Collection":  Set VarAsFile = CollAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) = "File"
        Case TypeName(v_items) = "Dictionary":  Set VarAsFile = DictAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) Like "*()":      Set VarAsFile = ArryAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbArray:        Set VarAsFile = ArryAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbString:       Set VarAsFile = StringAsFile(v_items, v_file, v_append)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Private Function VarAsStrg(ByVal v_items As Variant) As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsStrg"
    
    On Error GoTo eh
    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsStrg = CollAsStrg(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsStrg = DictAsStrg(v_items)
        Case TypeName(v_items) = "File":        VarAsStrg = FileAsStrg(v_items)
        Case TypeName(v_items) = "Nothing":     VarAsStrg = vbNullString
        Case TypeName(v_items) Like "*()":      VarAsStrg = ArryAsStrg(v_items)
        Case VarType(v_items) = vbArray:        VarAsStrg = ArryAsStrg(v_items)
        Case VarType(v_items) = vbBoolean:      VarAsStrg = BooleanAsStrg(v_items)
        Case VarType(v_items) = vbString:       VarAsStrg = StringAsStrg(v_items)
        Case VBA.IsObject(v_items):             VarAsStrg = ObjectAsStrg(v_items)
        Case Else:                              VarAsStrg = CStr(v_items)
    End Select
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function VarDiffs(ByVal v_res As Variant, _
                          ByVal v_exp As Variant, _
                 Optional ByRef v_res_prfx As String = vbNullString, _
                 Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a variant (v_res) differs from another (v_exp).
' Precondition: Both are of the same VarType.
' ----------------------------------------------------------------------------
    Const PROC = "VarDiffs"
    
    On Error GoTo eh
    Dim vResult             As Variant
    Dim vResultExpected     As Variant
            
    sVrfctnType = TypeName(v_res)
    If VBA.IsObject(v_res) Then
         Select Case sVrfctnType
            Case "Workbook"
               VarDiffs = WrkbDiffs(v_res, v_exp)
            Case "Collection":  VarDiffs = CollDiffs(v_res, v_exp)
            Case "Dictionary":  VarDiffs = DictDiffs(v_res, v_exp)
            Case "File":        VarDiffs = FileDiffs(v_res, v_exp)
            Case "Range":       VarDiffs = RngeDiffs(v_res, v_exp)
            Case Else:          VarDiffs = Not v_res Is v_exp
        End Select
    Else
        Select Case True
            Case IsArray(v_res) And IsArray(v_exp)
                If ArryNoOfElements(v_res) = ArryNoOfElements(v_exp) Then
                    LogCollectVrfctnDetails "Array", "Array", vbNullString
                Else
                    sFailedNote = "Number of elements/items differ!"
                    LogCollectVrfctnDetails "Array", "Array", VR_FAILED
                End If
                VarDiffs = ArryDiffs(v_res, v_exp, v_res_prfx)
            Case TypeName(v_res) = "String"
                vResult = Exclude(VarAsStrg(v_res), sExcludeFromComparison)
                vResultExpected = Exclude(VarAsStrg(v_exp), sExcludeFromComparison)
                VarDiffs = StrgDiffs(vResult, vResultExpected, v_res_prfx, v_exp_prfx)
            Case Else
                VarDiffs = StrgDiffs(VarAsStrg(v_res), VarAsStrg(v_exp), v_res_prfx, v_exp_prfx)
        End Select
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WrkbDiffs(ByVal v_res As Workbook, _
                           ByVal v_exp As Workbook, _
                  Optional ByRef v_res_prfx As String = vbNullString, _
                  Optional ByRef v_exp_prfx As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    WrkbDiffs = StrgDiffs(v_res.Name, v_exp.Name, v_res_prfx, v_exp_prfx)

End Function

Private Function WrkbExists(ByVal ex_wbk As Variant) As Boolean
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
    Const PROC = "WrkbExists"
    
    On Error GoTo eh
    
    WrkbExists = fso.FileExists(ex_wbk)
    If Not WrkbExists Then
        On Error Resume Next
        WrkbExists = Application.Workbooks(ex_wbk) = Application.Workbooks(ex_wbk).Name
    End If
                
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WrkbIsOpen(ByVal w_wbk As String, _
                   Optional ByRef w_open_wbk As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a Workbook (w_wbk) is open.
' ----------------------------------------------------------------------------
    Const PROC  As String = "WrkbIsOpen"
    
    On Error Resume Next
    Set w_open_wbk = Application.Workbooks(w_wbk)
    WrkbIsOpen = Err.Number = 0
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

