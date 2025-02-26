Attribute VB_Name = "mBasic"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mBasic: Common VBA services (declarations, procedures,
' ======================= functions potentially usefull  in any VB-Project.
' When not used as a whole parts may optionally just be copied.
'
' Note: This component is supposed to be autonomous. I.e. is does not require
'       any other installed component. The Common VBA Message Services
'       (fMsg and mMsg) and the Common VBA Error Services (mErH) are optional
'       and only used when installed which is indicated by corresponding
'       Conditional Compile Arguments `mMsg = 1` and `mErH = 1`.
'
' See the Summary of services in the README
' https://github.com/warbe-maker/VBA-Basics/blob/master/README.md#summary-of-services
'
' supplemented by Specifics and use
' https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#specifics-and-usage-examples-for-mbasic-vba-services
'
' Requires:
' ---------
' Reference to "Microsoft Scripting Runtime"
' Reference to "Microsoft Visual Basic Application Extensibility .."
'
' W. Rauschenberger, Berlin Feb 2025
' ----------------------------------------------------------------------------
Public Const DCONCAT    As String = "||"    ' For concatenating and error with a general message (info) to the error description
Public Const DGT        As String = ">"
Public Const DLT        As String = "<"
Public Const DAPOST     As String = "'"
Public Const DKOMMA     As String = ","
Public Const DBSLASH    As String = "\"
Public Const DDOT       As String = "."
Public Const DCOLON     As String = ":"
Public Const DEQUAL     As String = "="
Public Const DSPACE     As String = " "
Public Const DEXCL      As String = "!"
Public Const DQUOTE     As String = """"    ' one " character

' Common xl constants grouped ----------------------------
Public Enum YesNo   ' ------------------------------------
    xlYes = 1       ' System constants (identical values)
    xlNo = 2        ' grouped for being used as Enum Type.
End Enum            ' ------------------------------------
Public Enum xlOnOff ' ------------------------------------
    xlOn = 1        ' System constants (identical values)
    xlOff = -4146   ' grouped for being used as Enum Type.
End Enum            ' ------------------------------------
Public Enum enAlign
    enAlignLeft = 1
    enAlignRight = 2
    enAlignCentered = 3
End Enum
Public Enum enDctOpt
    enAdd
    enReplace
    enIncrement
    enCollect
    enCollectSorted
End Enum

#If Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As Long
    Private Declare PtrSafe Function GetWindowLongPtr Lib "User32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetForegroundWindow Lib "User32.dll" () As Long
    Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
#End If

Private Const ERROR_BAD_FORMAT = 11&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_SUCCESS = 32&
Private Const GWL_STYLE As Long = -16
Private Const LOGPIXELSX = 88               ' Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72  ' A point is defined as 1/72 inches
Private Const WIN_NORMAL = 1         'Open Normal
Private Const WS_THICKFRAME As Long = &H40000

Private cyTimerTicksBegin       As Currency
Private cyTimerTicksEnd         As Currency
Private TimerSystemFrequency    As Currency

#If Not mMsg = 1 Then
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
#End If

Public Property Let Arry(Optional ByRef a_arr As Variant, _
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
    Set cllSpecNdcs = ArryIndices(a_indices)
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

Public Property Get Arry(Optional ByRef a_arr As Variant, _
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
    Const PROC = "Arry(Get)"
    
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
        Set a_indices = ArryIndices(a_indices) ' transforms any kind of provided inedx/indices into a Collection (1 to n)
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

Property Get ArryItem(Optional ByRef a_arr As Variant, _
                      Optional ByVal a_indices As Variant = Nothing, _
                      Optional ByVal a_default As Variant = vbNullString) As Variant
' ---------------------------------------------------------------------------
' Returns from an array (a_arr) the item addressed by indices (a_indices)
' which might be up to 8 dimensions provided as an array or a string with the
' indices delimited by a comma.
' ---------------------------------------------------------------------------
    Dim cllIndices As Collection
    
    Set cllIndices = ArryIndices(a_indices)
    
    On Error Resume Next
    Select Case cllIndices.Count
        Case 1: ArryItem = a_arr(cllIndices(1))
        Case 2: ArryItem = a_arr(cllIndices(1), cllIndices(2))
        Case 3: ArryItem = a_arr(cllIndices(1), cllIndices(2), cllIndices(3))
        Case 4: ArryItem = a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4))
        Case 5: ArryItem = a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5))
        Case 6: ArryItem = a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5), cllIndices(6))
        Case 7: ArryItem = a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5), cllIndices(6), cllIndices(7))
        Case 8: ArryItem = a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5), cllIndices(6), cllIndices(7), cllIndices(8))
    End Select

End Property

Property Let ArryItem(Optional ByRef a_arr As Variant, _
                      Optional ByVal a_indices As Variant = Nothing, _
                      Optional ByVal a_default As Variant = vbNullString, _
                               ByVal a_item As Variant)
' ---------------------------------------------------------------------------
' Writes an Item (a_Item) to an array (a_arr) by means of provided indices
' (a_indices) which covers up to 8 dimensions.
' ---------------------------------------------------------------------------
    Dim cllIndices As Collection
    
    Set cllIndices = ArryIndices(a_indices)
    
    On Error Resume Next ' not assignable items are ignored
    Select Case cllIndices.Count
        Case 1: a_arr(cllIndices(1)) = a_item
        Case 2: a_arr(cllIndices(1), cllIndices(2)) = a_item
        Case 3: a_arr(cllIndices(1), cllIndices(2), cllIndices(3)) = a_item
        Case 4: a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4)) = a_item
        Case 5: a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5)) = a_item
        Case 6: a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5), cllIndices(6)) = a_item
        Case 7: a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5), cllIndices(6), cllIndices(7)) = a_item
        Case 8: a_arr(cllIndices(1), cllIndices(2), cllIndices(3), cllIndices(4), cllIndices(5), cllIndices(6), cllIndices(7), cllIndices(8)) = a_item
    End Select
    
End Property

' ----------------------------------------------------------------------------------
' Universal read from and write to a Collection (as Property Get/Let).
' - Read: When the provided argument (c_argmnt) is an integer which addresses a
'         not existing index or an index of which the element = Empty, Empty is
'         returned, else the element's content.
'         When the provided argument is not an integer the index of the first element
'         which is identical with the provided argument is returned. When no Item
'         is identical Empty is returned.
' - Write/Let: Writes items with any index by filling/adding the gap with Empty
'   items.
' ----------------------------------------------------------------------------------
Public Property Get Coll(Optional ByRef c_coll As Collection, _
                         Optional ByVal c_argmnt As Variant = Empty, _
                         Optional ByVal c_dflt As Variant = Empty) As Variant
    Const PROC = "Coll(Get)"
    
    Dim i As Long
    
    If c_coll Is Nothing Then Set c_coll = New Collection
    Select Case True
        Case c_argmnt = Empty:           Coll = c_dflt
        Case IsInteger(c_argmnt) _
            And c_argmnt > c_coll.Count: Coll = c_dflt
        Case IsInteger(c_argmnt) _
            And c_argmnt > 1 _
            And c_argmnt <= c_coll.Count
            If IsObject(c_coll(c_argmnt)) _
            Then Set Coll = c_coll(c_argmnt) _
            Else Coll = c_coll(c_argmnt)
        Case Else
            '~~ When an argument is provided which is not an integer, the index of
            '~~ the first item is returned which is identical with the provided argument
            With c_coll
                If IsObject(c_argmnt) Then
                    For i = 1 To .Count
                        If .Item(i) Is c_argmnt Then
                            Coll = i
                            GoTo xt
                        End If
                    Next i
                Else
                    For i = 1 To .Count
                        If .Item(i) = c_argmnt Then
                            Coll = i
                            GoTo xt
                        End If
                    Next i
                End If
            End With
    End Select

xt:
End Property

Public Property Let Coll(Optional ByRef c_coll As Collection, _
                         Optional ByVal c_argmnt As Variant = Empty, _
                         Optional ByVal c_dflt As Variant = Empty, _
                                  ByVal c_var As Variant)
    Const PROC = "Coll(Let)"
    
    If c_coll Is Nothing Then Set c_coll = New Collection
    With c_coll
        Select Case True
            Case c_argmnt = Empty
                '~~ When no index is provided the Item is simply added
                .Add c_var
            Case c_argmnt > .Count
                '~~ For any indey beyond count elements = Empty are added up to the provided index - 1
                Do While .Count < c_argmnt - 1
                    .Add Empty
                Loop
                .Add c_var
            Case c_argmnt <= .Count _
              And Not IsObject(c_coll(c_argmnt))
                '~~ Update!
                .Remove c_argmnt
                .Add c_var, c_argmnt
            Case c_argmnt <= c_coll.Count _
              And IsObject(c_coll(c_argmnt))
                '~~ Replace value in object
              
        End Select
    End With

End Property

Public Property Let dict(Optional ByRef d_dct As Dictionary = Nothing, _
                         Optional ByVal d_key As Variant, _
                         Optional ByVal d_arg As Variant = enAdd, _
                                  ByVal d_Item As Variant)
    Const PROC = "Dict_Let"
    
    Dim arr As Variant
    Dim cll As Collection
    Dim i   As Long
    Dim j   As Long
    Dim tmp As Variant

    If d_dct Is Nothing Then Set d_dct = New Dictionary
    Select Case d_arg
        Case enAdd:         If Not d_dct.Exists(d_key) Then d_dct.Add d_key, d_Item
        Case enCollect:
                            If Not d_dct.Exists(d_key) Then
                                Set cll = New Collection
                                cll.Add d_Item
                                d_dct.Add d_key, cll
                            Else
                                If Not TypeName(d_dct(d_key)) = "Collection" _
                                Then Err.Raise AppErr(2), ErrSrc(PROC), "The option ""collect"" is not applicable since the existing Item is not a Collection!"
                                Set cll = d_dct(d_key)
                                cll.Add d_Item
                                d_dct.Remove d_key
                                d_dct.Add d_key, cll
                            End If
        Case enCollectSorted
                            If Not d_dct.Exists(d_key) Then
                                Set cll = New Collection
                                cll.Add d_Item
                                d_dct.Add d_key, cll
                            Else
                                If Not TypeName(d_dct(d_key)) = "Collection" _
                                Then Err.Raise AppErr(3), ErrSrc(PROC), "The option ""collect"" is not applicable since the existing Item is not a Collection!"
                                Set cll = d_dct(d_key)
                                cll.Add d_Item
                                
                                '~~ Sort items by bubble sort
                                With cll
                                    ReDim arr(0 To .Count - 1)
                                    For i = 0 To .Count - 1
                                        arr(i) = .Item(i + 1)
                                    Next i
                                End With
                                For i = LBound(arr) To UBound(arr) - 1
                                    For j = i + 1 To UBound(arr)
                                        If arr(i) > arr(j) Then
                                            tmp = arr(j)
                                            arr(j) = arr(i)
                                            arr(i) = tmp
                                        End If
                                    Next j
                                Next i
                                Set cll = New Collection
                                With cll
                                    For i = LBound(arr) To UBound(arr)
                                        .Add arr(i)
                                    Next i
                                End With
                                d_dct.Remove d_key
                                d_dct.Add d_key, cll
                            End If
        Case enIncrement
                            If Not d_dct.Exists(d_key) Then
                                d_dct.Add d_key, d_Item
                            Else
                                If Not IsNumeric(d_dct(d_key)) _
                                Then Err.Raise AppErr(1), ErrSrc(PROC), "The option ""increment"" is not applicable since the Item is not numeric!"
                                tmp = d_dct(d_key)
                                tmp = tmp + d_Item
                                d_dct.Remove d_key
                                d_dct.Add d_key, tmp
                            End If
        Case enReplace:     If d_dct.Exists(d_key) Then d_dct.Remove d_key: d_dct.Add d_key, d_Item
    End Select
    
End Property

' ----------------------------------------------------------------------------
' Universal Dictionary add, replace, increase Item value service.
' Specific: when the Item is a string in the form "-n" or "+n" whereby n is a
' numeric value, the Item with the corresponding key is incremented or
' decremented provided it exists or it is added with the negative or positive
' value as Item.
' - Get/Read:       Returns a default when an Item is not existing
' - Let/Write:      Handles the provided Item in correspondence with the
'                   provided argument (d_arg) which defaults to add.
' - Argument d_arg: With Read it specifies the default returend when the key
'                   doesnt exist, with write
' W. Rauschenberger Berlin, Oct 2024
' ----------------------------------------------------------------------------
Public Property Get dict(Optional ByRef d_dct As Dictionary = Nothing, _
                         Optional ByVal d_key As Variant, _
                         Optional ByVal d_arg As Variant = Nothing) As Variant
                         
    If d_dct Is Nothing Then
        If VarType(d_arg) = vbObject Then Set dict = d_arg Else dict = d_arg
    Else
        If d_dct.Exists(d_key) Then
            If VarType(d_dct(d_key)) = vbObject Then Set dict = d_dct(d_key) Else dict = d_dct(d_key)
        Else
            If VarType(d_arg) = vbObject _
            Then Set dict = d_arg _
            Else dict = d_arg
        End If
    End If
    
End Property

Private Property Get TimerSecsElapsed() As Currency:        TimerSecsElapsed = TimerTicksElapsed / SysFrequency:        End Property

Private Property Get TimerSysCurrentTicks() As Currency:    getTickCount TimerSysCurrentTicks:                          End Property

Private Property Get TimerTicksElapsed() As Currency:       TimerTicksElapsed = cyTimerTicksEnd - cyTimerTicksBegin:    End Property

Public Function Align(ByVal a_string As String, _
             Optional ByVal a_align As enAlign = enAlignLeft, _
             Optional ByVal a_length As Long = 0, _
             Optional ByVal a_fill As String = " ", _
             Optional ByVal a_margin As String = vbNullString, _
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
'   when the string is one Item of serveral organized in columns when the
'   column delimiter is a vbNullString. When the column delimiter is a |
'   usually a marign of a single space is used
' - The string (a_strn) may contain already leading or trailing spaces of
'   the left are preserved when the string is left aligned and the right
'   spaces are preserved when the string is right aligned.
'   spaces.
' - The function is also used to align items arragend in columns.
'
' About column arranged alignment (a_col_arranged = True):
' --------------------------------------------------------
' - The provided length (a_length) is regarded the maximum. I.e. when the
'   string (a_string) is longer it is  truncated to the right
' - The returned string has any margin (left and right) added
' - When a fill (a_fill) other than a single space the returned string has
'   at least one fill (a_fill) added. In the example below the fill is " -"
'   a_string = "xxx", a_lenght = 10, a_margin = " ", returns
'   " xxx -------- " which is a final lenght of 14.
'
' About normal alignment (a_col_arranged = True):
' -----------------------------------------------
' - The provided length is the final length returned. I.e. when the provided
'   string (a_string) exceeds the length (a_length) the Result is truncated
'   to the right
' - Any specified margin is ignored
' - Fills (a_fill) are added to the right to the extent the length of the
'   provided string (a_string) allows it without exceeding the final length
'   (a_lenght).
'
' Uses: AlignLeft, AlignRight, AlignCenter
'
' W. Rauschenberger, Berlin Jun 2024
' ----------------------------------------------------------------------------
    Const PROC = "Align"
    
    On Error GoTo eh
    Dim lLength As Long
    
    If a_string = vbNullString Then
        '~~ A non provided string results in one filled with a_fill
        '~~ enclosed in margins (a_margin)
        Align = a_margin & String$(a_length, a_fill) & a_margin
        GoTo xt
    End If
                
    If a_length = 0 _
    Then lLength = Len(a_string) _
    Else lLength = a_length

    If a_fill = vbNullString Then a_fill = " "
    Select Case a_align
        Case enAlignLeft:       Align = AlignLeft(a_string, lLength, a_fill, a_margin, a_col_arranged)
        Case enAlignRight:      Align = AlignRght(a_string, lLength, a_fill, a_margin, a_col_arranged)
        Case enAlignCentered:   Align = AlignCntr(a_string, lLength, a_fill, a_margin, a_col_arranged)
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function AlignCntr(ByVal a_string As String, _
                          ByRef a_length As Long, _
                 Optional ByVal a_fill As String = " ", _
                 Optional ByVal a_margin As String = vbNullString, _
                 Optional ByVal a_col_arranged = False) As String
' ----------------------------------------------------------------------------
' Returns a string (a_string) centered within a guven width (a_length). When
' truncate (a_truncate) is False (the default), the returned with (a_with)
' may be extended to the width of the string (a_steing). The Result may not be
' centered exactly.
'
' Uses: AlignTruncated, AlignWidthExMargin
' ----------------------------------------------------------------------------
    Const PROC = "AlignCntr"
    
    On Error GoTo eh
    Dim lLoop       As Long
    Dim s           As String
    Dim sFillLeft   As String
    Dim sFillRight  As String
    Dim sFill       As String
    Dim lWidth      As Long
    Dim lLoss       As Long

    If a_length = 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A zero length is nor supported by alignment centered!"

    s = a_string
    Select Case a_fill
        Case " ":           sFill = " ": sFillLeft = " ":          sFillRight = " ":            lLoss = 0
        Case "-":           sFill = "-": sFillLeft = vbNullString: sFillRight = vbNullString:   lLoss = 2
        Case " -", "- ":    sFill = "-": sFillLeft = " ":          sFillRight = " ":            lLoss = 4
        Case "=":           sFill = "=": sFillLeft = vbNullString: sFillRight = vbNullString:   lLoss = 2
        Case " =", "= ":    sFill = "=": sFillLeft = " ":          sFillRight = " ":            lLoss = 4
        Case Else
            Err.Raise AppErr(2), ErrSrc(PROC), _
            "The specified fill is not applicable for a centered alignment! " & _
            "Accepted is "" "", ""-"", ""="", "" -"", ""- "", "" ="", ""= """
    End Select
    
    '~~ Determine the final width/length of the returned string - excluding margins
    If a_col_arranged Then
        lWidth = AlignWidthExMargin(a_fill, a_length, enAlignCentered)
        s = AlignTruncated(s, enAlignCentered, lWidth - lLoss)
    Else
        lWidth = a_length
        If Len(s) > lWidth Then
            AlignCntr = Left(s, lWidth)
            GoTo xt
        End If
    End If
    
    Do While Len(s) < lWidth
        lLoop = lLoop + 1
        Select Case True
            Case Len(s) = lWidth
                Exit Do
            Case Len(s & sFillRight) <= lWidth
                s = s & sFillRight
                sFillRight = sFill
                If Len(s) < lWidth Then
                    s = sFillLeft & s ' add fill left
                    sFillLeft = sFill
                End If
        End Select
        If lLoop > lWidth Then
            Stop
        End If
    Loop
    AlignCntr = s
    If a_col_arranged Then AlignCntr = a_margin & s & a_margin _

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function AlignLeft(ByRef a_string As String, _
                          ByRef a_length As Long, _
                 Optional ByVal a_fill As String = " ", _
                 Optional ByVal a_margin As String = vbNullString, _
                 Optional ByVal a_col_arranged As Boolean = False) As String
' -----------------------------------------------------------------------------------
' Returns a string (a_strng) left aligned by:
' - adding fill characters to the right up to the specified length (a_length),
'   whereby the filling may start with a single space
' - enclose the Result in left and right margins - which default to a vbNullString
'   when a_col_arranged = True
' Specifics:
' - Left spaces provided with the string (a_string) are preserved
' - The final length is the specified length plus a left and a right marging - which
'   may be a vbNullString
' - It is considered that the string may be an column Item which may look ate the end
'   "| xxxx" when a column delimiter  | is used, else just "xxxx" when the column
'   delimiter is already a single space. In that case the margin (a_margin) would be
'   a vbNullString.
' Precondition: The string (a_string) not/no longer contains any implicit alignment
'               specs.
' Uses: AlignTruncated, AlignWidthExMargin
' -----------------------------------------------------------------------------------

    Dim s           As String
    Dim sFillStart  As String
    Dim sFill       As String
    Dim sFillEnd    As String
    Dim lLoop       As Long
    Dim lWidth      As Long
    Dim lLoss       As Long
    
    '~~ Preparations for the final string which will have the following syntax:
    '~~ - a_col_arranged = True: <margin><string><fill-start>[<fill>][<fill>]...<fill-end><margin>
    '~~ - a_col_arranged = False: <margin><string>
    s = a_string
    Select Case a_fill
        Case " ":   sFillStart = vbNullString: sFill = " ": sFillEnd = " ":             lLoss = 0
        Case "-":   sFillStart = vbNullString: sFill = "-": sFillEnd = vbNullString:    lLoss = 1
        Case " -":  sFillStart = " ":          sFill = "-": sFillEnd = vbNullString:    lLoss = 2
        Case ".":   sFillStart = vbNullString: sFill = ".": sFillEnd = vbNullString:    lLoss = 1
        Case " .":  sFillStart = " ":          sFill = ".": sFillEnd = vbNullString:    lLoss = 2
        Case ".:":  sFillStart = vbNullString: sFill = ".": sFillEnd = ":":             lLoss = 2
        Case " .:": sFillStart = " ":          sFill = ".": sFillEnd = ":":             lLoss = 3
    End Select
        
    '~~ Determine the final width/length of the returned string - excluding margins
    If a_col_arranged Then
        lWidth = AlignWidthExMargin(a_fill, a_length, enAlignLeft)
        s = AlignTruncated(s, enAlignLeft, lWidth - lLoss)
    Else
        lWidth = a_length
        If Len(s) > lWidth Then
            AlignLeft = Left(s, lWidth)
            Exit Function
        End If
    End If

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
    AlignLeft = s
    If a_col_arranged Then AlignLeft = a_margin & s & a_margin _
    
End Function

Public Function AlignRght(ByVal a_string As String, _
                 Optional ByRef a_length As Long = 0, _
                 Optional ByVal a_fill As String = " ", _
                 Optional ByVal a_margin As String = vbNullString, _
                 Optional ByVal a_col_arranged As Boolean = False) As String
' -----------------------------------------------------------------------------------
' Returns a string (a_string) aligned right with at the left filled (a_fill) in a
' given length (a_lenght), enclosed in specified margins (a_margin).
'
' Uses: AlignTruncated, AlignWidthExMargin
' -----------------------------------------------------------------------------------
    Const PROC = "AlignRght"
    
    On Error GoTo eh
    Dim lLoop       As Long
    Dim lWidth      As Long
    Dim s           As String
    Dim sFill       As String
    Dim sFillEnd    As String
    Dim sFillStart  As String
    Dim lLoss       As Long
    
    If a_length = 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A zero length is nor supported by alignment right!"
    
    s = a_string
    If a_fill = vbNullString Then a_fill = " "
    Select Case a_fill
        Case " ":           sFillStart = vbNullString: sFill = " ": sFillEnd = " ":             lLoss = 0
        Case ".":           sFillStart = vbNullString: sFill = ".": sFillEnd = vbNullString:    lLoss = 1
        Case " .", ". ":    sFillStart = " ":          sFill = ".": sFillEnd = vbNullString:    lLoss = 2
        Case "-":           sFillStart = vbNullString: sFill = "-": sFillEnd = vbNullString:    lLoss = 1
        Case " -", "- ":    sFillStart = " ":          sFill = "-": sFillEnd = vbNullString:    lLoss = 2
    End Select
    
    '~~ Determine the final width/length of the returned string - excluding margins
    If a_col_arranged Then
        lWidth = AlignWidthExMargin(a_fill, a_length, enAlignRight)
        s = AlignTruncated(s, enAlignRight, lWidth - lLoss)
    Else
        lWidth = a_length
        If Len(s) > a_length Then
            AlignRght = Left(s, a_length)
            GoTo xt
        End If
    End If

    Do
        lLoop = lLoop + 1
        Select Case True
            Case Len(s) >= lWidth:                                Exit Do
            Case Len(sFillStart & s) = lWidth:                    s = sFillStart & s:           sFillStart = vbNullString
            Case Len(sFillEnd & s) = lWidth:                      s = sFillEnd & s              ' last fill add
            Case Len(sFillEnd & sFill & sFillStart & s) <= lWidth: s = sFill & sFillStart & s:   sFillStart = vbNullString
            Case Else:                                            s = sFill & s
        End Select
        If lLoop > lWidth Then
            Stop
        End If
    Loop
    AlignRght = s
    If a_col_arranged Then AlignRght = a_margin & s & a_margin _
                   
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AlignTruncated(ByRef a_string As String, _
                                ByVal a_align As enAlign, _
                                ByVal a_space As Long) As String
' -----------------------------------------------------------------------------------
' Returns a string truncated when the space available (a_space) is less then the
' width required by the string (a_string).
' -----------------------------------------------------------------------------------
    Dim s As String
    
    Select Case a_align
        Case enAlignLeft:       s = RTrim$(a_string)
        Case enAlignCentered:   s = Trim$(a_string)
        Case enAlignRight:      s = LTrim$(a_string)
    End Select
        
    If Len(s) > a_space Then s = Left(s, a_space)
    AlignTruncated = s
    
End Function

Private Function AlignWidthExMargin(ByVal a_fill As String, _
                                    ByVal a_length As Long, _
                                    ByVal a_align As enAlign) As Long
' -----------------------------------------------------------------------------------
' The provided lenght indicates the maximum string length. However, for another fill
' but a single space the final returned width always includes the fill (one for
' left or right and two for centered.
' Note: col arranged assumes that the columns do have a delimiter |, which makes a
'       margin appropriate provided one is specified - which is the default.
' -----------------------------------------------------------------------------------
    Dim l As Long
    
    l = a_length ' default

    If a_fill <> " " Then
        l = l + Len(a_fill)   ' one in any case
        If a_align = enAlignCentered Then l = l + Len(a_fill) ' one fill at the left and another one at the right
    End If
    AlignWidthExMargin = l
  
End Function

Public Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error number never conflicts
' with VB runtime error. Thr function returns a given positive number
' (app_err_no) with the vbObjectError added - which turns it to negative. When
' the provided number is negative it returns the original positive "application"
' error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function AppIsInstalled(ByVal exe As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an application (exe) is installed, i.e. the provided name
' is found in the VBA.Environ$ Path.
' ----------------------------------------------------------------------------
    Dim i As Long: i = 1
    Do Until VBA.Environ$(i) Like "Path=*": i = i + 1: Loop
    AppIsInstalled = Environ$(i) Like "*" & exe & "*"
End Function

Public Function ArryAsDict(ByVal a_arr As Variant) As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary with all items of an array (a_arr) with the indices
' delimited by a comma as key.
' Note: The function handles multidimensional array up to 8 dimensions.
' ------------------------------------------------------------------------------
    
    Dim dct As New Dictionary
    
    ArryAsDictAdd a_arr, dct, "", 1, ArryDims(a_arr)
    Set ArryAsDict = dct
    
End Function

Private Sub ArryAsDictAdd(ByVal a_arr As Variant, _
                          ByRef a_dct As Object, _
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
    Dim arrNdcs As Variant
    Dim vItem   As Variant
    
    With a_dct
        For i = LBound(a_arr, a_dim) To UBound(a_arr, a_dim)
            If a_key = vbNullString _
            Then sKey = i _
            Else sKey = a_key & "," & i
            
            If a_dim < a_dims Then
                ArryAsDictAdd a_arr, a_dct, sKey, a_dim + 1, a_dims
            Else
                arrNdcs = Split(sKey, ",")
                vItem = ArryItem(a_arr, arrNdcs)
                If Not IsError(vItem) And Not IsEmpty(vItem) Then
                    .Add sKey, vItem
                End If
            End If
        Next i
    End With
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub ArryAsRnge(ByVal a_arr As Variant, _
                      ByRef a_rng As Range, _
             Optional ByVal a_transpose As Boolean = False)
' ----------------------------------------------------------------------------
' Transferes the content of an array (a_arr) to a range (a_rng) amd returns
' the resulting range.
' ----------------------------------------------------------------------------
    Const PROC = "ArryAsRnge"
    
    On Error GoTo eh
    Dim lDims   As Long
    Dim rTarget As Range
    Dim lBase   As Long
    
    lDims = ArryDims(a_arr)
    lBase = LBound(Array("x"))
    
    Select Case True
        Case lDims = 1 And a_transpose:     '~~ One column, n rows
                                            Set rTarget = a_rng.Cells(1, 1).Resize(UBound(a_arr) + 1, 1)
                                            rTarget.Value = Application.Transpose(a_arr)
                                            Set a_rng = rTarget
                                            
        Case lDims = 1 And Not a_transpose: '~~ n columns, one row
                                            Set rTarget = a_rng.Cells(1, 1).Resize(1, UBound(a_arr) + 1)
                                            rTarget.Value = a_arr
                                            Set a_rng = rTarget
        
        Case lDims = 2 And Not a_transpose: Set rTarget = a_rng.Cells(1, 1).Resize(UBound(a_arr, 1) + 1, UBound(a_arr, 2) + 1)
                                            rTarget.Value = a_arr
                                            Set a_rng = rTarget
        
        Case lDims = 2 And a_transpose:     Set rTarget = a_rng.Cells(1, 1).Resize(UBound(a_arr, 2) + 1, UBound(a_arr, 1) + 1)
                                            rTarget.Value = Application.Transpose(a_arr)
                                            Set a_rng = rTarget
                
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), _
                      "The provided array has more than 2 dimensions but only 1 or two may be transfered to a range!"
    End Select
    
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

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
    Set cllSpecNdcs = ArryIndices(a_indices)
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

Public Function ArryCompare(ByVal a_v1 As Variant, _
                            ByVal a_v2 As Variant, _
                   Optional ByVal a_stop_after As Long = 0, _
                   Optional ByVal a_ignore_case As Boolean = True, _
                   Optional ByVal a_ignore_empty As Boolean = True) As Dictionary
' --------------------------------------------------------------------------
' Returns a Dictionary with n (a_stop_after) lines which are different
' between array 1 (a_v1) and array 2 (a_v2) with the line number as the
' key and the two different lines as Item in the form: '<line>'vbLf'<line>'
' When no differnece is encountered the returned Dictionary is empty.
' When no a_stop_after <> 0 is provided all lines different are returned
' --------------------------------------------------------------------------
    Const PROC = "ArryCompare"
    
    On Error GoTo eh
    Dim l       As Long
    Dim i       As Long
    Dim lMethod As VbCompareMethod
    Dim dct     As New Dictionary
    
    If a_ignore_case Then lMethod = vbTextCompare Else lMethod = vbBinaryCompare
    
    With dct
        If Not mBasic.ArryIsAllocated(a_v1) And mBasic.ArryIsAllocated(a_v2) Then
            If a_ignore_empty Then mBasic.ArryTrimm a_v2
            For i = LBound(a_v2) To UBound(a_v2)
                .Add i + 1, "'" & a_v2(i) & "'" & vbLf
            Next i
        ElseIf mBasic.ArryIsAllocated(a_v1) And Not mBasic.ArryIsAllocated(a_v2) Then
            If a_ignore_empty Then mBasic.ArryTrimm a_v1
            For i = LBound(a_v1) To UBound(a_v1)
                .Add i + 1, "'" & a_v1(i) & "'" & vbLf
            Next i
        ElseIf Not mBasic.ArryIsAllocated(a_v1) And Not mBasic.ArryIsAllocated(a_v2) Then
            GoTo xt
        End If
        
        If a_ignore_empty Then mBasic.ArryTrimm a_v1
        If a_ignore_empty Then mBasic.ArryTrimm a_v2
        
        l = 0
        For i = LBound(a_v1) To Min(UBound(a_v1), UBound(a_v2))
            If StrComp(a_v1(i), a_v2(i), lMethod) <> 0 Then
                .Add i + 1, "'" & a_v1(i) & "'" & vbLf & "'" & a_v2(i) & "'"
                l = l + 1
                If a_stop_after <> 0 And l >= a_stop_after Then
                    GoTo xt
                End If
            End If
        Next i
        
        If UBound(a_v1) < UBound(a_v2) Then
            For i = UBound(a_v1) + 1 To UBound(a_v2)
                .Add i + 1, "''" & vbLf & " '" & a_v2(i) & "'"
                l = l + 1
                If a_stop_after <> 0 And l >= a_stop_after Then
                    GoTo xt
                End If
            Next i
            
        ElseIf UBound(a_v2) < UBound(a_v1) Then
            For i = UBound(a_v2) + 1 To UBound(a_v1)
                .Add i + 1, "'" & a_v1(i) & "'" & vbLf & "''"
                l = l + 1
                If a_stop_after <> 0 And l >= a_stop_after Then
                    GoTo xt
                End If
            Next i
        End If
    End With

xt: Set ArryCompare = dct
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
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

Public Function ArryDiffers(ByVal ad_v1 As Variant, _
                             ByVal ad_v2 As Variant, _
                    Optional ByVal ad_ignore_empty_items As Boolean = False, _
                    Optional ByVal ad_comp_mode As VbCompareMethod = vbTextCompare) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when array (ad_v1) differs from array (ad_v2).
' ----------------------------------------------------------------------------
    Const PROC  As String = "ArryDiffers"
    
    Dim i       As Long
    Dim j       As Long
    Dim va()    As Variant
    
    On Error GoTo eh
    
    If Not mBasic.ArryIsAllocated(ad_v1) And mBasic.ArryIsAllocated(ad_v2) Then
        va = ad_v2
    ElseIf mBasic.ArryIsAllocated(ad_v1) And Not mBasic.ArryIsAllocated(ad_v2) Then
        va = ad_v1
    ElseIf Not mBasic.ArryIsAllocated(ad_v1) And Not mBasic.ArryIsAllocated(ad_v2) Then
        GoTo xt
    End If

    '~~ Leading and trailing empty items are ignored by default
    mBasic.ArryTrimm ad_v1
    mBasic.ArryTrimm ad_v2
    
    If Not ad_ignore_empty_items Then
        On Error Resume Next
        If Not ad_ignore_empty_items Then
            On Error Resume Next
            ArryDiffers = Join(ad_v1) <> Join(ad_v2)
            If Err.Number = 0 Then GoTo xt
            '~~ At least one of the joins resulted in a string exeeding the maximum possible lenght
            For i = LBound(ad_v1) To Min(UBound(ad_v1), UBound(ad_v2))
                If ad_v1(i) <> ad_v2(i) Then
                    ArryDiffers = True
                    Exit Function
                End If
            Next i
        End If
    Else
        i = LBound(ad_v1)
        j = LBound(ad_v2)
        For i = i To mBasic.Min(UBound(ad_v1), UBound(ad_v2))
            While Len(ad_v1(i)) = 0 And i + 1 <= UBound(ad_v1)
                i = i + 1
            Wend
            While Len(ad_v2(j)) = 0 And j + 1 <= UBound(ad_v2)
                j = j + 1
            Wend
            If i <= UBound(ad_v1) And j <= UBound(ad_v2) Then
                If StrComp(ad_v1(i), ad_v2(j), ad_comp_mode) <> 0 Then
                    ArryDiffers = True
                    GoTo xt
                End If
            End If
            j = j + 1
        Next i
        If j < UBound(ad_v2) Then
            ArryDiffers = True
        End If
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function ArryDims(ByVal a_arr As Variant, _
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
    
    For i = 1 To 8
        On Error Resume Next
        arrSpecs(1) = LBound(a_arr, i)
        If Err.Number <> 0 Then
            Exit For
        Else
            arrSpecs(2) = UBound(a_arr, i)
            cll.Add arrSpecs
        End If
    Next i
        
xt: ArryDims = cll.Count
    Set a_dim_specs = cll
    a_dims = cll.Count
    Set cll = Nothing

End Function

Public Sub ArryErase(ParamArray a_args() As Variant)
' ----------------------------------------------------------------------------
' Erases any arguments which are an Array. Non array arguments are ignored.
' Note: erase does not effect an array's specifics regarding dimensions and
' bounds.
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    On Error GoTo xt
    If UBound(a_args) >= LBound(a_args) Then
        For Each v In a_args
            If IsArray(v) Then
                Erase v
            End If
        Next v
    End If
    
xt: Err.Clear
    On Error GoTo 0
End Sub

Public Function ArryIndices(ParamArray a_indices() As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns provided indices (a_indices) as Collection whereby the indices may
' be provided: - as integers (one for each dimension)
'              - as an array of integers
'              - as a string of integers delimited by a , (comma).
' ----------------------------------------------------------------------------
    Const PROC = "ArryIncides"
    
    Dim cll As New Collection
    Dim i   As Long
    Dim v   As Variant
    
    On Error GoTo xt
    If UBound(a_indices) >= LBound(a_indices) Then
        Select Case True
            Case Not ArryIsAllocated(a_indices)
                '~~ No indices had been provided
                Exit Function
            Case IsArray(a_indices(LBound(a_indices)))
                '~~ The first item is an array if incides
                For Each v In a_indices(LBound(a_indices))
                    If IsInteger(CInt(Trim(v))) _
                    Then cll.Add v _
                    Else Err.Raise AppErr(1), ErrSrc(PROC), "At least one of the items provided as array is not an integer value!"
                Next v
            Case TypeName(a_indices(LBound(a_indices))) = "Collection"
                '~~ The first item is a Collection of indices
                Set cll = a_indices(LBound(a_indices))
            Case TypeName(a_indices(LBound(a_indices))) = "String"
                '~~ The first item is a string of incides delimited by a comma
                For Each v In Split(a_indices(LBound(a_indices)), ",")
                    If IsNumeric(v) Then
                        cll.Add CInt(Trim(v))
                    Else
                        Err.Raise AppErr(2), ErrSrc(PROC), "At least one of the items provided a string delimited by a comma is not an integer value!"
                    End If
                Next v
            Case a_indices(LBound(a_indices)) = Empty
            Case Else
                For Each v In a_indices
                    If IsInteger(v) _
                    Then cll.Add v _
                    Else Err.Raise AppErr(3), ErrSrc(PROC), "Any of the provided indices is not an integer value!"
                Next v
        End Select
    End If
    
xt: On Error GoTo 0
    Set ArryIndices = cll
    Set cll = Nothing
    
End Function

Public Function ArryIsAllocated(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the array (a_arr) is allocated, i.e. has at least one Item.
' ----------------------------------------------------------------------------
    
    On Error Resume Next
    ArryIsAllocated = UBound(a_arr) >= LBound(a_arr)
    On Error GoTo 0
    Err.Clear
    
End Function

Public Function ArryItems(ByVal a_arr As Variant, _
                 Optional ByVal a_default_excluded As Boolean = False) As Long
' ----------------------------------------------------------------------------
' Returns the number of items in a multi-dimensional array or a nested array.
' The latter is an array of which one or more items are again possibly multi-
' dimensional arrays. An unallocated array returns 0. When items which return
' a type specific default are excluded (a_default_excluded = True) only
' "active= items/elements are counted.
' ----------------------------------------------------------------------------
    
    Dim lDim    As Long
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

Public Function ArryNextIndex(ByVal a_arr As Variant, _
                              ByRef a_indices() As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns the logical next index for a multidimensional array (a_arr) based
' on a given index (a_indices) which is an array of indices, one for each
' dimension in the provided array (a_arr). When the next index would be one
' above all upper bounds the function returns FALSE, else the indices (a_indices)
' are the logically next one.
' Precondition: The indices array (a_indices) is specified as 1 to x for each
'               dimension.
' ----------------------------------------------------------------------------
    Const PROC = "ArryNextIndex"

    Dim i       As Long
    Dim lDims   As Long
    Dim lNext   As Long

    On Error GoTo xt
    If LBound(a_indices) <= LBound(a_indices) Then
        lDims = ArryDims(a_arr)
        If UBound(a_indices) <> lDims _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided current index does not correctly consider the number of dimensions in the provided array!"
    
        For i = lDims To 1 Step -1
            '~~ Loop through all dimensins from last to first
            lNext = a_indices(i) + 1
            If lNext > UBound(a_arr, i) Then
                a_indices(i) = LBound(a_arr, i)
            Else
                a_indices(i) = lNext
                ArryNextIndex = True
                GoTo xt
            End If
        Next i
    End If
    '~~ When the upper bounds of all dimesions had been specified the function return False
    ArryNextIndex = False

xt: On Error GoTo 0
End Function

Public Sub ArryReDim(ByRef a_arr As Variant, _
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
    
    Set regex = CreateObject("VBScript.RegExp")
    If Not ArryIsAllocated(a_specs) Then GoTo xt
    
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
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub ArryRemoveItems(ByRef a_arry As Variant, _
                  Optional ByVal a_item As Long = -99999, _
                  Optional ByVal a_indx As Long = -99999, _
                  Optional ByVal a_items As Long = 1)
' ------------------------------------------------------------------------------
' Returns arr 1-dim array (a_arry) with the number of elements (a_items) removed
' whereby the start element may be indicated by the item number 1,2,.. (a_item)
' or the index (a_indx), both must be within the array's bounds. Any
' inappropriate provision of arguments raises an application errore. When the
' last Item in of an array is removed the returned array appears erased (i.e is
' no longer allocated).
'
' Uses: ArryDims to check for the number of dimensions which is limited to 1.
'
' W. Rauschenberger, Berlin Feb 2022
' ------------------------------------------------------------------------------
    Const PROC = "ArryRemoveItems"

    On Error GoTo eh
    Dim lItems  As Long ' number of items in the provided array
    Dim i       As Long
    Dim lLbnd   As Long ' the provided array's LBound
    Dim lUbnd   As Long ' the provided array's UBound
      
    On Error GoTo xt    ' when the provided array is not allocated
    If UBound(a_arry) >= LBound(a_arry) Then
        lLbnd = LBound(a_arry)
        lUbnd = UBound(a_arry)
        lItems = lUbnd - lLbnd + 1
            
        Select Case True
            Case Not ArryDims(a_arry) = 1
                Err.Raise AppErr(1), ErrSrc(PROC), "Array must not be multidimensional!"
            
            Case a_item = -99999 And a_indx <> -99999
                If a_indx < lLbnd Or a_indx > lUbnd _
                Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided index is out of the array's bounds!"
            
            Case a_item >= 0 And a_indx = -99999
                '~~ When an item is provided which is within the number of items in the array
                '~~ it is transformed into an index
                If a_item < 1 Or a_item > lItems _
                Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided item number is 0 or exceeds the number of items in the array!"
                ' lBnd = 1 And item = 4 > index = item
                ' lBnd = 4 And item = 2 > index = 5
                a_indx = lLbnd + a_item - 1
            
            Case Else
                Err.Raise AppErr(4), ErrSrc(PROC), "Neither an item nor an index had been provided or both which is conflicting!"
        End Select
        
        For i = a_indx + a_items To lUbnd
            a_arry(i - a_items) = a_arry(i)
        Next i
        ReDim Preserve a_arry(lLbnd To i - a_items - 1)
    
    End If
    
xt: On Error GoTo 0
    Err.Clear
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub ArryTrimm(ByRef a_arr As Variant)
' ------------------------------------------------------------------------------
' Returns the array (a_arr) with all items the leading and trailing spaces
' removed and Empty items removed. Any items with vbCr, vbCrLf, vbLf are ignored.
' When the array contains only empty items the returned array is erased.
'
' Uses: Arry to write/keep non empt items
' ------------------------------------------------------------------------------
    Const PROC  As String = "ArryTrimm"

    On Error GoTo eh
    Dim i   As Long
    Dim arr As Variant
    
    On Error GoTo xt
    If UBound(a_arr) >= LBound(a_arr) Then
        For i = LBound(a_arr) To UBound(a_arr)
            Select Case True
                Case IsError(a_arr(i))
                Case a_arr(i) = Empty
                Case Trim$(a_arr(i)) = vbNullString
                Case Else
                    If Not Len(Trim(a_arr(i))) = 0 _
                    Then Arry(arr) = Trim(a_arr(i))
            End Select
        Next i
    End If
    a_arr = arr
    
xt: On Error GoTo 0
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Function ArryUnload(ByVal a_arr As Variant) As Variant
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
    Dim i       As Long
    Dim sIndices   As String
    Dim arrIndices As Variant
    
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

Public Sub ArryUnloadToDict(ByVal a_arr As Variant, _
                            ByRef a_dct As Dictionary, _
                   Optional ByRef a_ndc As String)
' ----------------------------------------------------------------------------
' Returns all items in a multidimensional array (a_arr) as Dictionary (a_dct)
' with the elements as Item and the indices delimited by a comma as key. The
' procedure is called recursively for nested arrays.
' ----------------------------------------------------------------------------
    Dim i        As Long
    Dim iCurrent As String
    
    If a_dct Is Nothing Then Set a_dct = New Dictionary
    
    If IsArray(a_arr) Then
        If ArryIsAllocated(a_arr) Then
            For i = LBound(a_arr) To UBound(a_arr)
                iCurrent = a_ndc & i
                ArryUnloadToDict a_arr(i), a_dct, iCurrent & ","
            Next i
        End If
    Else ' is an Item
        a_dct.Add Left(a_ndc, Len(a_ndc) - 1), a_arr
    End If

End Sub

Public Sub BoC(ByVal b_id As String, _
      Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Bnd-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.BoC b_id, b_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.BoC b_id, b_args
#End If
End Sub

Public Sub BoP(ByVal b_proc As String, _
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

Public Function CleanTrim(ByVal c_string As String, _
                 Optional ByVal c_conv_non_breaking_spaces As Boolean = True) As String
' ----------------------------------------------------------------------------------
' Returns the string 'c_string' cleaned from any non-printable characters.
' ----------------------------------------------------------------------------------
    Const PROC = "CleanTrim"
    
    On Error GoTo eh
    Dim l       As Long
    Dim aClean  As Variant
    
    aClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                     21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If c_conv_non_breaking_spaces Then c_string = Replace(c_string, Chr$(160), " ")
    For l = LBound(aClean) To UBound(aClean)
        If InStr(c_string, Chr$(aClean(l))) Then c_string = Replace(c_string, Chr$(aClean(l)), vbNullString)
    Next
    
xt: CleanTrim = c_string
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Sub DelayedAction(ByVal d_secs As Long, _
                Optional ByVal d_action As String = vbNullString, _
                Optional ByVal d_action_arg1 As Variant, _
                Optional ByVal d_action_arg2 As Variant, _
                Optional ByVal d_action_arg3 As Variant, _
                Optional ByVal d_action_arg4 As Variant, _
                Optional ByVal d_action_arg5 As Variant)
' ----------------------------------------------------------------------------
' Waits for n (d_sec) seconds and performs a provided action (d_action)
' whereby the action needs to be a fully specified Application.Run action.
' <workbook-name>!<comp>.<proc>
'
' Specific: The action may be passed with up to 5 arguments. Though optional
'           none preceeding must be ommitted. I.e. when the second argument
'           is provided the first one is obligatory.
' ----------------------------------------------------------------------------
                
    Dim v As Variant
    v = Now() + d_secs / 100000
    Do While Now() < v
        DoEvents
    Loop
    If d_action <> vbNullString Then
        Select Case True
            Case Not IsMissing(d_action_arg5): Application.Run d_action, d_action_arg1, d_action_arg2, d_action_arg3, d_action_arg4, d_action_arg5
            Case Not IsMissing(d_action_arg4): Application.Run d_action, d_action_arg1, d_action_arg2, d_action_arg3, d_action_arg4
            Case Not IsMissing(d_action_arg3): Application.Run d_action, d_action_arg1, d_action_arg2, d_action_arg3
            Case Not IsMissing(d_action_arg2): Application.Run d_action, d_action_arg1, d_action_arg2
            Case Not IsMissing(d_action_arg1): Application.Run d_action, d_action_arg1
            Case Else:                         Application.Run d_action
        End Select
    End If
    
End Sub

Public Sub DelayedActionTest()
    
    DelayedAction 1, ""
    DelayedAction 1, ThisWorkbook.Name & "!mBasic.DelayedTest"
    DelayedAction 1, ThisWorkbook.Name & "!mBasic.DelayedTest1", "Arg1"
    DelayedAction 1, ThisWorkbook.Name & "!mBasic.DelayedTest2", "Arg1", "Arg2"
    
End Sub

                                  
Public Sub DelayedTest()
    Debug.Print "DelayedTest performed"
End Sub

Public Sub DelayedTest1(ByVal arg As String)
    Debug.Print "DelayedTest1 with args " & arg
End Sub

Public Sub DelayedTest2(ByVal arg1 As String, _
                        ByVal arg2 As String)
    Debug.Print "DelayedTest2 with args " & arg1 & " and " & arg2
End Sub

Public Function DictAsArry(ByVal d_dct As Dictionary, _
                  Optional ByVal d_dimsfrom As Variant = Nothing, _
                  Optional ByVal d_dimsto As Variant = Nothing) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Dim aResult()   As Variant
    Dim key         As Variant
    Dim keyParts()  As String
    Dim i           As Long
    Dim lDims       As Long
    Dim arrIndices     As Variant
    Dim arr         As Variant
    Dim arrFrom     As Variant
    Dim arrTo       As Variant
    Dim v           As Variant
    
    If d_dct Is Nothing Then GoTo xt
    If d_dct.Count = 0 Then GoTo xt
      
    If d_dimsfrom Is Nothing And d_dimsto Is Nothing Then
        '~~ When no new dimension specs are provided they are retrieved from the dictionar key
        arrFrom = Split(d_dct.Keys()(0), ",")
        lDims = UBound(arrFrom) + 1
        arrTo = Split(d_dct.Keys(d_dct.Count - 1), ",")
    Else
        arrFrom = d_dimsfrom
        arrTo = d_dimsto
    End If
    
    Select Case lDims
        Case 1: ReDim arr(arrFrom(0) To arrTo(0))
        Case 2: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1))
        Case 3: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1), arrFrom(2) To arrTo(2))
        Case 4: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1), arrFrom(2) To arrTo(2), arrFrom(3) To arrTo(3))
        Case 5: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1), arrFrom(2) To arrTo(2), arrFrom(3) To arrTo(3), arrFrom(4) To arrTo(4))
        Case 6: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1), arrFrom(2) To arrTo(2), arrFrom(3) To arrTo(3), arrFrom(4) To arrTo(4), arrFrom(5) To arrTo(5))
        Case 7: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1), arrFrom(2) To arrTo(2), arrFrom(3) To arrTo(3), arrFrom(4) To arrTo(4), arrFrom(5) To arrTo(5), arrFrom(6) To arrTo(6))
        Case 8: ReDim arr(arrFrom(0) To arrTo(0), arrFrom(1) To arrTo(1), arrFrom(2) To arrTo(2), arrFrom(3) To arrTo(3), arrFrom(4) To arrTo(4), arrFrom(5) To arrTo(5), arrFrom(6) To arrTo(6), arrFrom(7) To arrTo(7))
    End Select
    
    For Each v In d_dct
        arrIndices = Split(v, ",")
        ArryItem(arr, arrIndices) = d_dct(v)
    Next v
    DictAsArry = arr
xt:

End Function

Public Sub DictTest()

    Dim dct As Dictionary
    
    dict(dct, "X") = 1              ' Add
    dict(dct, "X") = 2              ' Replace
    Debug.Assert dict(dct, "X") = 2 ' Assert
    
    dict(dct, "X") = "+5"           ' Increment
    Debug.Assert dict(dct, "X") = 7 ' Assert
    
    dict(dct, "X") = "-3"           ' Decrement
    Debug.Assert dict(dct, "X") = 4 ' Assert
    
    '~~ Assert defaults when not existing
    Debug.Assert dict(dct, "B", Nothing) Is Nothing
    Debug.Assert dict(dct, "B", 0) = 0
    Debug.Assert dict(dct, "B") = vbNullString      ' No default returns default
    
End Sub

Public Function DynExec(ByVal d_vbp As VBProject, _
                        ByVal d_proc As String, _
                        ByVal d_args As String) As Variant
' ----------------------------------------------------------------------------
' Dynamically constructs a function which calls a provided procedure (d_proc)
' by passing provided arguments (d_args). In other words the function
' provides a means to dynamically execute a named procedure (d_proc) by means
' of a temporary component created in a VB-Project (d_vbp).
'
' Note: This is one of the very rare working solutions to dynamically execute
'       code. However, is comes with the disadvantage that debugging is
'       hindered substantially since code holds cannot be used when this
'       function is executed.
'       It thus will be appropriate to keep the procedure which calls DynExec
'       short and DynExec the only call executed.
'
' Requires a reference to: "Microsoft Visual Basic Application Extensibility .."
'
' W. Rauschenberger, Berlin Feb 2025
' ----------------------------------------------------------------------------
    Const PROC = "DynExec"
    
    Dim objTempMod As Object
    Dim vResult    As Variant
    Dim sCode      As String
    
    sCode = "Function TempFunction() As Variant" & vbCrLf
    sCode = sCode & "TempFunction = " & d_proc & "(" & d_args & ")" & vbCrLf
    sCode = sCode & "End Function"
    
    On Error GoTo eh
    ' Create a temporary module
    Set objTempMod = d_vbp.VBComponents.Add(vbext_ct_StdModule)
    
    ' Add the dynamic d_sCode to the temporary module
    objTempMod.CodeModule.AddFromString sCode
    
    ' Run the temporary function and get the Result
    DynExec = Application.Run("TempFunction")
    
    ' Remove the temporary module
    Application.VBE.ActiveVBProject.VBComponents.Remove objTempMod
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub EoC(ByVal e_id As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.EoC e_id, e_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.EoC e_id, e_args
#End If
End Sub

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

Public Function ErrMsg(ByVal err_source As String, _
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
    '~~ About
    ErrDesc = err_dscrptn
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    End If
    '~~ Type of error
    If err_no < 0 Then
        ErrType = "Application Error ": ErrNo = AppErr(err_no)
    Else
        ErrType = "VB Runtime Error ":  ErrNo = err_no
        If err_dscrptn Like "*DAO*" _
        Or err_dscrptn Like "*ODBC*" _
        Or err_dscrptn Like "*Oracle*" _
        Then ErrType = "Database Error "
    End If
    
    '~~ Title
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")
    '~~ Description
    ErrText = "Error: " & vbLf & ErrDesc
    '~~ About
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mBasic." & e_proc
End Function

Public Function IsInteger(ByVal i_value As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is numeric, and has no decimals.
' ----------------------------------------------------------------------------
    If Not i_value = Empty Then
        If IsNumeric(i_value) _
        Then IsInteger = (i_value = Int(i_value))
    End If
End Function

Public Function IsString(ByVal v As Variant, _
                Optional ByVal vbnullstring_is_a_string = False) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is neither an object nor numeric.
' ----------------------------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = v
    If Err.Number = 0 Then
        If Not IsNumeric(v) Then
            If (s = vbNullString And vbnullstring_is_a_string) _
            Or s <> vbNullString _
            Then IsString = True
        End If
    End If
End Function

Public Function KeySort(ByRef s_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (s_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If s_dct Is Nothing Then GoTo xt
    If s_dct.Count = 0 Then GoTo xt
    
    With s_dct
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
            .Add key:=vKey, Item:=s_dct.Item(vKey)
        Next i
    End With
    
xt: Set s_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub MakeFormResizable()
' ---------------------------------------------------------------------------
' This part is from Leith Ross                                              |
' Found this Code on:                                                       |
' https://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html |
'                                                                           |
' All credits belong to him                                                 |
' ---------------------------------------------------------------------------
    Const WS_THICKFRAME = &H40000
    Const GWL_STYLE As Long = (-16)
    
    Dim lStyle As LongPtr
    Dim hwnd As LongPtr
    Dim retVal

    hwnd = GetForegroundWindow
    
    lStyle = GetWindowLongPtr(hwnd, GWL_STYLE Or WS_THICKFRAME)
    retVal = SetWindowLongPtr(hwnd, GWL_STYLE, lStyle)

End Sub

Public Function Max(ParamArray m_args() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the maximum of privided arguments (m_args). Arguments supported
' are numeric values, strings, Arrays, and Collections. Of strings items the
' length is considered the numeric value. When an argument is an Array or a
' Collection the items again may be a numeric value, string, Arrays, or
' Collection.
'
' Constraint: When an argument is an Array or a Collection and it contains a
'             single item which again is an Array or a Collection, an error
'             is raised. Nested Arrays and/or Collections may again have items
'             which are an Array or a Collection but only among other numeric
'             or string items. This constraint is caused by the fact that
'             Function calls with a single argument which is an Array or a
'             Collection are considered recursive calls and thus will result in
'             a loop.
'
' W. Rauschenberger, Berlin Feb 2025
' ------------------------------------------------------------------------------
    Const PROC = "Max"
    
    Dim v1      As Variant
    Dim v2      As Variant
    Dim vMax1   As Long
    Dim vMax2   As Variant
    
    On Error GoTo xt ' when no argument is provided
    If UBound(m_args) >= LBound(m_args) Then
        On Error GoTo eh
        For Each v1 In m_args
            Select Case True
                Case IsArray(v1) Or TypeName(v1) = "Collection"
                    If UBound(m_args) = LBound(m_args) Then
                        '~~ Its the one and only argument and thus a possible recursive call
                        For Each v2 In v1
                            Select Case True
                                Case IsArray(v2) Or TypeName(v2) = "Collection"
                                    Err.Raise AppErr(1), ErrSrc(PROC), "Single argument array with single item array (jagged array) not supported!"
                                Case VarType(v2) = vbString
                                    If Len(v2) > vMax1 Then vMax1 = Len(v2)
                                Case IsNumeric(v2)
                                    If v2 > vMax1 Then vMax1 = v2
                            End Select
                        Next v2
                    Else
                        '~~ Resursive Max call with the array as the one and only argument
                        vMax2 = Max(v1)
                        If vMax2 > vMax1 Then vMax1 = vMax2
                    End If
                Case VarType(v1) = vbString
                    If Len(v1) > vMax1 Then vMax1 = Len(v1)
                Case IsNumeric(v1)
                    If v1 > vMax1 Then vMax1 = v1
            End Select
        Next v1
    End If

xt: Max = vMax1
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function Min(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' --------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Public Function PointsPerPixel() As Double
' ----------------------------------------
' Return DPI
' ----------------------------------------
    
    Dim hDC             As Long
    Dim lDotsPerInch    As Long
    
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC

End Function

Public Function ProgramIsInstalled(ByVal sProgram As String) As Boolean
        ProgramIsInstalled = InStr(Environ$(18), sProgram) <> 0
End Function

Public Sub README(Optional ByVal r_base_url As String = "https://github.com/warbe-maker/VBA-Basics", _
                  Optional ByVal r_bookmark As String = vbNullString)
' ----------------------------------------------------------------------------
' By default displays the README for this component, optionally with a
' bookmark (r_bookmark) in the computer's default browser.
' ----------------------------------------------------------------------------
    If r_bookmark = vbNullString Then
        mBasic.ShellRun r_base_url
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        mBasic.ShellRun r_base_url & r_bookmark
    End If
        
End Sub

Public Function RngeAsArry(ByVal r_rng As Range) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    RngeAsArry = r_rng.Value
    
End Function

Public Function SelectFolder(Optional ByVal sTitle As String = "Select a Folder") As String
' ----------------------------------------------------------------------------
' Returns the selected folder or a vbNullString if none had been selected.
' ----------------------------------------------------------------------------
    
    Dim sFolder As String
    
    SelectFolder = vbNullString
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = sTitle
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    SelectFolder = sFolder

End Function

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

Public Function Spaced(ByVal s As String) As String
' ----------------------------------------------------------------------------
' Returns a non-breaking-spaced string with any spaces already in the string
' tripled and leading or trailing spaces unstripped.
' Example: Spaced("this is spaced") returns = "t h i s   i s   s p a c e d"
' ----------------------------------------------------------------------------
    Dim a() As Byte
    Dim i   As Long
    
    If s = vbNullString Then Exit Function
    a = StrConv(Trim$(s), vbFromUnicode)
    Spaced = Chr$(a(LBound(a)))
    For i = LBound(a) + 1 To UBound(a)
        If Chr$(a(i)) = " " _
        Then Spaced = Spaced & Chr$(160) & Chr$(160) & Chr$(160) _
        Else Spaced = Spaced & Chr$(160) & Chr$(a(i))
    Next i

End Function

Public Function StackEd(ByVal s_stck As Collection, _
               Optional ByRef s_Item As Variant = -999999999, _
               Optional ByRef s_lvl As Long = 0) As Variant
' ----------------------------------------------------------------------------
' Returns:
' - TRUE when an Item (s_Item) is stacked at a given level (s_lvl) or
'   when no level is provided, when it is stacked at any level. In the latter
'   case the level (s_lvl) is returned.
' - The stacked Item when none is provided and but a level is.
'
' Restriction: The function works with any kind of object an an Item which is
'              not -999999999, which is regarded no Item is provided.
' ----------------------------------------------------------------------------
    Const PROC = "StckEd"
    
    On Error GoTo eh
    Dim i As Long
    
    If s_stck Is Nothing Then Set s_stck = New Collection
    
    Select Case True
        Case s_lvl <> 0 And s_lvl > s_stck.Count
            GoTo xt

        Case s_lvl = 0 And VarType(s_Item) <> vbObject
            If s_Item = -999999999 Then GoTo xt
            '~~ A specific Item has been provided
            For i = 1 To s_stck.Count
                If s_stck(i) = s_Item Then
                    s_lvl = i
                    StackEd = True
                    GoTo xt
                End If
            Next i
        
        Case s_lvl = 0 And VarType(s_Item) = vbObject
            For i = 1 To s_stck.Count
                If s_stck(i) Is s_Item Then
                    s_lvl = i
                    StackEd = True
                    GoTo xt
                End If
            Next i
                    
        Case VarType(s_Item) <> vbObject
            If s_Item = -999999999 Then
                StackEd = s_stck(s_lvl)
            ElseIf s_stck(s_lvl) = s_Item Then
                StackEd = True
                GoTo xt
            End If
        
        Case VarType(s_Item) = vbObject
            If s_stck(s_lvl) Is s_Item Then
                StackEd = True
                GoTo xt
            End If
    End Select
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function StackIsEmpty(ByVal stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Common Stack Empty check service. Returns True when either there is no stack
' (stck Is Nothing) or when the stack is empty (items count is 0).
' ----------------------------------------------------------------------------
    StackIsEmpty = stck Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = stck.Count = 0
End Function

Public Function StackPop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Common Stack Pop service. Returns the last Item pushed on the stack (stck)
' and removes the Item from the stack. When the stack (stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    If StackIsEmpty(stck) Then GoTo xt
    
    If IsObject(stck(stck.Count)) _
    Then Set StackPop = stck(stck.Count) _
    Else StackPop = stck(stck.Count)
    stck.Remove stck.Count

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Sub StackPush(ByRef stck As Collection, _
                     ByVal stck_item As Variant)
' ----------------------------------------------------------------------------
' Common Stack Push service. Pushes (adds) an Item (stck_Item) to the stack
' (stck). When the provided stack (stck) is Nothing the stack is created.
' ----------------------------------------------------------------------------
    Const PROC = "StckPush"
    
    On Error GoTo eh
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_item

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Function StackTop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Common Stack Top service. Returns the top Item from the stack (stck), i.e.
' the Item last pushed. If the stack is empty a vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckTop"
    
    On Error GoTo eh
    If StackIsEmpty(stck) Then GoTo xt
    If IsObject(stck(stck.Count)) _
    Then Set StackTop = stck(stck.Count) _
    Else StackTop = stck(stck.Count)

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function SysFrequency(Optional ByVal s_f As Currency = 0) As Currency
    If s_f = 0 Then s_f = TimerSystemFrequency
    If s_f = 0 Then getFrequency TimerSystemFrequency
    SysFrequency = TimerSystemFrequency
End Function

Public Function TimedDoEvents(Optional t_source As String = vbNullString, _
                              Optional t_debug_print As Boolean = False) As String
' ---------------------------------------------------------------------------
' Returns the elapsed time in seconds of DoEvents statement.
'
' Background: DoEvents every now and then are concidered to solve problems.
'             However, when looking at the description of DoEvents its effect
'             may appears miraculous. However, stil when it helps is should
'             be known that DoEvents allow keyboard interaction while a
'             process executes. In case of a loop with embedded DoEvents,
'             this may be a godsend. But it as well may cause unpredictable
'             results. This little procedure at least documents in VBE's
'             immediate window the resulting performace delay in milliseconds.
' ---------------------------------------------------------------------------
    Const PROC          As String = "TimedDoEvents"
    Const TIMER_FORMAT  As String = "0.00000"
    
    Dim cBegin      As Currency
    Dim cEnd        As Currency
    Dim cElapsed    As Currency
    
    mBasic.TimerBegin cBegin
    DoEvents
    mBasic.TimerEnd cBegin, cEnd, cElapsed, TIMER_FORMAT
    If t_source <> vbNullString Then t_source = " (" & Trim(t_source) & ")"
    TimedDoEvents = Format((cElapsed / SysFrequency) * 1000, TIMER_FORMAT) & " seconds " & t_source
    If t_debug_print Then Debug.Print ErrSrc(PROC) & ": " & TimedDoEvents
    
End Function

Public Sub TimerBegin(ByRef t_begin As Currency)
    t_begin = TimerSysCurrentTicks
End Sub

Public Function TimerEnd(ByVal t_begin As Currency, _
                Optional ByRef t_end As Currency, _
                Optional ByRef t_elapsed As Currency, _
                Optional ByVal t_format As String = "hh:mm:ss.0000") As String
' ---------------------------------------------------------------------------
' Returns, based on provided begin-ticks (t_begin)
' - the end-ticks (t_end)
' - the elapsed ticks (t_elapsed)
' - the elapsed time in the provided format (t_format)
' ---------------------------------------------------------------------------
    t_end = TimerSysCurrentTicks
    t_elapsed = ((t_end - t_begin) / SysFrequency) * 1000
    TimerEnd = Format(t_elapsed, t_format)
    
End Function

