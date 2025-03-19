Attribute VB_Name = "mBasic"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mBasic: Common Component providing basic VBA services
' ======================= potentially usefull in any VB-Project. When not used
' as a whole parts may optionally just be copied.
'
' As with all (my) Common Components this component is supposed to be
' autonomous in the sense it does not require any other installed component.
' The Common Component "VBA Message Services" (fMsg and mMsg) and
' the Common Component "VBA Error Services" (mErH) are optional by means
' of Conditional Compile Arguments `mMsg = 1` and `mErH = 1`. I.e. these
' indicate when the corresponding components are installed and then will
' automatically be used. When not installed all services, in case of an error
' display a best possible error message - with a debugging option however.
'
' See the Summary of services by README "mBasic"
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
                         Optional ByVal a_ndcs As Variant, _
                                  ByVal a_var As Variant)
' ----------------------------------------------------------------------------
' Common WRITE to an array service. The service returns an array (a_arr) with
' the provided Item (a_var) either:
' - Simply added, when no index/indices (a_ndcs) are provided
' - Having added or replaced an Item at a given index/indices (a_ndcs) by
'   concidering that the returned array has new from/to specifics for any of
'   its dimensions at any level whereby the from specific for any dimension
'   remains the same
' - Having created a 1 to 8 dimensions array regarding the provided indices
'   (a_ndcs) with the provided Item added or replaced
' - Having re-dimensioned the provided/input array to the provided indices
'   (a_ndcs) with the provided Item (a_var) added or updated
'
' The index/indices (a_ndcs) may be provided as:
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
    Dim aSpecsIn        As Variant              ' the provided array's dimensions specifics/bounds
    Dim bIsAllocated    As Boolean
    Dim cllBounds       As New Collection
    Dim cllBoundsOut    As New Collection
    Dim arrDimsSpec     As Variant              ' the provided indices, either none or one for each dimension
    Dim lBase           As Long
    Dim lDimsArray      As Long                 ' the provided/input array's number of dimensions - when dimensioned or allocated
    Dim lDimsSpec       As Long
    Dim lDimsOutBound   As Long
    Dim i               As Long
    
    '~~ Get the current array's dimension specifics - when available
    bIsAllocated = ArryIsAllocated(a_arr)
    lDimsArray = ArryDims(a_arr)
    If lDimsArray <> 0 _
    Then aSpecsIn = ArrySpecs(a_arr, lDimsArray) ' The provided array's number of dimesion and their from/to specifics (when dimensioned)
    arrDimsSpec = ArryNdcs(lDimsSpec, a_ndcs) ' the array is only allocated when lDimsSpec > 0
    
    Select Case True
        '~~ Handling of cases when no indices had been provided
        Case lDimsSpec = 0 And lDimsArray > 1
            Err.Raise AppErr(3), ErrSrc(PROC), "For multi-dimensional array indices are obligatory but missing or incomplete!"
            
        Case lDimsSpec = 0 And lDimsArray = 0                           ' No index provided for a yet un-dimensioned array
            lBase = LBound(Array(1))
            ReDim a_arr(lBase To lBase)                                 ' results in a 1-dim array with one Item
            a_arr(lBase) = a_var
            
        Case lDimsSpec = 0 And lDimsArray = 1                           ' No index provided for a 1-dim array
            ReDim Preserve a_arr(LBound(a_arr) To UBound(a_arr) + 1)    ' upper bound extended on the fly
            a_arr(UBound(a_arr)) = a_var
                
        Case lDimsSpec = 0: Stop ' should all have been considered
        
        '~~ Handling cases with provided 1-dim indices
        Case lDimsArray = 0 And lDimsSpec = 1             ' Array yet not dimensioned with an index provided for a 1-dim array
            lBase = LBound(Array(1))
            ReDim a_arr(lBase To arrDimsSpec(1))
            a_arr(arrDimsSpec(1)) = a_var
            
        Case lDimsArray = 0 And lDimsSpec > 1               ' Multi-dim array yet not dimensioned/allocated
            lBase = LBound(Array(1))
            Select Case lDimsSpec                           ' dimensioned wit lower bound = Base Option and upper bound = provided indices
                Case 1: ReDim a_arr(lBase To arrDimsSpec(1))
                Case 2: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2))
                Case 3: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2), lBase To arrDimsSpec(3))
                Case 4: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2), lBase To arrDimsSpec(3), lBase To arrDimsSpec(4))
                Case 5: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2), lBase To arrDimsSpec(3), lBase To arrDimsSpec(4), lBase To arrDimsSpec(5))
                Case 6: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2), lBase To arrDimsSpec(3), lBase To arrDimsSpec(4), lBase To arrDimsSpec(5), lBase To arrDimsSpec(6))
                Case 7: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2), lBase To arrDimsSpec(3), lBase To arrDimsSpec(4), lBase To arrDimsSpec(5), lBase To arrDimsSpec(6), lBase To arrDimsSpec(7))
                Case 8: ReDim a_arr(lBase To arrDimsSpec(1), lBase To arrDimsSpec(2), lBase To arrDimsSpec(3), lBase To arrDimsSpec(4), lBase To arrDimsSpec(5), lBase To arrDimsSpec(6), lBase To arrDimsSpec(7), lBase To arrDimsSpec(8))
            End Select
            ArryItem(a_arr, arrDimsSpec) = a_var
            
        Case bIsAllocated And lDimsArray = 1 _
         And arrDimsSpec(1) >= aSpecsIn(1, 1) _
         And arrDimsSpec(1) <= aSpecsIn(1, 2)               ' Write to a 1-dim array with an index within specified bounds
            a_arr(arrDimsSpec(1)) = a_var
        
        Case lDimsArray = 1 And lDimsSpec = 1 _
         And arrDimsSpec(1) >= aSpecsIn(1, 1) _
         And arrDimsSpec(1) > UBound(a_arr), _
            Not bIsAllocated And lDimsSpec = 1                    ' An Item in a 1-dim array is addressed beyond the current upper bound
            ReDim Preserve a_arr(aSpecsIn(1, 1) To arrDimsSpec(1))  ' The array is redimed/expanded on the fly accordingly
            a_arr(arrDimsSpec(1)) = a_var
                
        Case lDimsArray > 1 _
         And lDimsArray = lDimsSpec           ' The dimensions specified are identical with those of the provided array
                                                    ' The dimensions' index may still differ
            ArryBounds a_arr, arrDimsSpec, cllBoundsOut, cllBounds, lDimsOutBound
            
            If lDimsOutBound = 0 Then
                ArryItem(a_arr, arrDimsSpec) = a_var
                
            ElseIf lDimsOutBound = 1 And IsArray(cllBoundsOut(cllBoundsOut.count)) Then
                '~~ Since the out of bounds dimension is the last one it can be re-dimensioned by VBA.ReDim.
                Select Case lDimsArray
                    Case 1: ReDim a_arr(aSpecsIn(1, 1) To arrDimsSpec(1))
                    Case 2: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To arrDimsSpec(2))
                    Case 3: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To aSpecsIn(2, 2), aSpecsIn(3, 1) To arrDimsSpec(3))
                    Case 4: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To aSpecsIn(2, 2), aSpecsIn(3, 1) To aSpecsIn(3, 2), aSpecsIn(4, 1) To arrDimsSpec(4))
                    Case 5: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To aSpecsIn(2, 2), aSpecsIn(3, 1) To aSpecsIn(3, 2), aSpecsIn(4, 1) To aSpecsIn(4, 2), aSpecsIn(5, 1) To arrDimsSpec(5))
                    Case 6: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To aSpecsIn(2, 2), aSpecsIn(3, 1) To aSpecsIn(3, 2), aSpecsIn(4, 1) To aSpecsIn(4, 2), aSpecsIn(5, 1) To aSpecsIn(5, 2), aSpecsIn(6, 1) To arrDimsSpec(6))
                    Case 7: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To aSpecsIn(2, 2), aSpecsIn(3, 1) To aSpecsIn(3, 2), aSpecsIn(4, 1) To aSpecsIn(4, 2), aSpecsIn(5, 1) To aSpecsIn(5, 2), aSpecsIn(6, 1) To aSpecsIn(6, 2), aSpecsIn(7, 1) To arrDimsSpec(7))
                    Case 8: ReDim a_arr(aSpecsIn(1, 1) To aSpecsIn(1, 2), aSpecsIn(2, 1) To aSpecsIn(2, 2), aSpecsIn(3, 1) To aSpecsIn(3, 2), aSpecsIn(4, 1) To aSpecsIn(4, 2), aSpecsIn(5, 1) To aSpecsIn(5, 2), aSpecsIn(6, 1) To aSpecsIn(6, 2), aSpecsIn(7, 1) To aSpecsIn(7, 2), aSpecsIn(8, 1) To arrDimsSpec(8))
                End Select
                ArryItem(a_arr, arrDimsSpec) = a_var
            
            ElseIf lDimsOutBound >= 1 And Not IsArray(cllBoundsOut(cllBoundsOut.count)) Then
                '~~ Any other bounds than the last dimension's bounds have changed which means that
                '~~ the dimensions redim cannot be provided by VBA's ReDim but requires ArryRedim.
                Select Case cllBounds.count
                    Case 1: ArryReDim a_arr, "1:" & cllBounds(1) & " to " & cllBounds(1)
                    Case 2: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ")
                    Case 3: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ") _
                                           , "3:" & Join(cllBounds(3), " to ")
                    Case 4: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ") _
                                           , "3:" & Join(cllBounds(3), " to ") _
                                           , "4:" & Join(cllBounds(4), " to ")
                    Case 5: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ") _
                                           , "3:" & Join(cllBounds(3), " to ") _
                                           , "4:" & Join(cllBounds(4), " to ") _
                                           , "5:" & Join(cllBounds(5), " to ")
                    Case 6: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ") _
                                           , "3:" & Join(cllBounds(3), " to ") _
                                           , "4:" & Join(cllBounds(4), " to ") _
                                           , "5:" & Join(cllBounds(5), " to ") _
                                           , "6:" & Join(cllBounds(6), " to ")
                    Case 7: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ") _
                                           , "3:" & Join(cllBounds(3), " to ") _
                                           , "4:" & Join(cllBounds(4), " to ") _
                                           , "5:" & Join(cllBounds(5), " to ") _
                                           , "6:" & Join(cllBounds(6), " to ") _
                                           , "7:" & Join(cllBounds(7), " to ")
                    Case 8: ArryReDim a_arr, "1:" & Join(cllBounds(1), " to ") _
                                           , "2:" & Join(cllBounds(2), " to ") _
                                           , "3:" & Join(cllBounds(3), " to ") _
                                           , "4:" & Join(cllBounds(4), " to ") _
                                           , "5:" & Join(cllBounds(5), " to ") _
                                           , "6:" & Join(cllBounds(6), " to ") _
                                           , "7:" & Join(cllBounds(7), " to ") _
                                           , "8:" & Join(cllBounds(8), " to ")
                End Select
                ArryItem(a_arr, arrDimsSpec) = a_var
                            
            Else
                Stop
            End If
        
        Case Not bIsAllocated And lDimsArray = 0
            '~~ For a yet not allocated and not dimensioned array the provided index/indices determine the uper bound
            Select Case lDimsSpec
                Case 1: ReDim a_arr(arrDimsSpec(1))
                Case 2: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2))
                Case 3: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2), arrDimsSpec(3))
                Case 4: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2), arrDimsSpec(3), arrDimsSpec(4))
                Case 5: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2), arrDimsSpec(3), arrDimsSpec(4), arrDimsSpec(5))
                Case 6: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2), arrDimsSpec(3), arrDimsSpec(4), arrDimsSpec(5), arrDimsSpec(6))
                Case 7: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2), arrDimsSpec(3), arrDimsSpec(4), arrDimsSpec(5), arrDimsSpec(6), arrDimsSpec(7))
                Case 8: ReDim a_arr(arrDimsSpec(1), arrDimsSpec(2), arrDimsSpec(3), arrDimsSpec(4), arrDimsSpec(5), arrDimsSpec(6), arrDimsSpec(7), arrDimsSpec(8))
            End Select
            ArryItem(a_arr, arrDimsSpec) = a_var
    End Select
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Arry(Optional ByRef a_arr As Variant, _
                         Optional ByVal a_ndcs As Variant = Nothing) As Variant
' ----------------------------------------------------------------------------
' Common, universal READ from array service supporting up to 8 dimensions.
' The service returns the Item of a provided array (a_arr) with given indices
' (a_ndcs) which might be single integers or an array of indices, a comma
' delimited string of indices/integers or a Collection of integers. when the
' item does not exist an array type specific default is returend.
'
' W. Rauschenberger Berlin, Mar 2025
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
    
    a_ndcs = ArryNdcs(lDims, a_ndcs) ' return an array 1 to n indices with the indices provided a string, array, integer or Collection
    
    On Error Resume Next
    Select Case lDims
        Case 0:
        Case 1: If bObject Then Set Arry = a_arr(a_ndcs(1)) Else Arry = a_arr(a_ndcs(1))
        Case 2: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2))
        Case 3: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3))
        Case 4: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4))
        Case 5: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5))
        Case 6: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6))
        Case 7: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7))
        Case 8: If bObject Then Set Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7), a_ndcs(8)) Else Arry = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7), a_ndcs(8))
    End Select
    
xt:
End Property

Private Property Get ArryItem(Optional ByRef a_arr As Variant, _
                              Optional ByVal a_ndcs As Variant) As Variant
' ---------------------------------------------------------------------------
' Returns from an array (a_arr) the Item addressed by indices (a_ndcs)
' which might be up to 8 dimensions provided as an array or a string with the
' indices delimited by a comma.
' ---------------------------------------------------------------------------
    Dim lDims As Long
    If IsArray(a_ndcs) Then lDims = UBound(a_ndcs) Else lDims = a_ndcs.count
    
    On Error Resume Next
    Select Case lDims
        Case 1: ArryItem = a_arr(a_ndcs(1))
        Case 2: ArryItem = a_arr(a_ndcs(1), a_ndcs(2))
        Case 3: ArryItem = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3))
        Case 4: ArryItem = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4))
        Case 5: ArryItem = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5))
        Case 6: ArryItem = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6))
        Case 7: ArryItem = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7))
        Case 8: ArryItem = a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7), a_ndcs(8))
    End Select

End Property

Private Property Let ArryItem(Optional ByRef a_arr As Variant, _
                              Optional ByVal a_ndcs As Variant, _
                                       ByVal a_Item As Variant)
' ---------------------------------------------------------------------------
' Writes an Item (a_Item) to an array (a_arr) by means of provided indices
' (a_ndcs) which covers up to 8 dimensions.
' ---------------------------------------------------------------------------
    Dim lDims As Long
    If IsArray(a_ndcs) Then lDims = UBound(a_ndcs) Else lDims = a_ndcs.count
    
    On Error Resume Next ' not assignable items are ignored
    Select Case lDims
        Case 1: a_arr(a_ndcs(1)) = a_Item
        Case 2: a_arr(a_ndcs(1), a_ndcs(2)) = a_Item
        Case 3: a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3)) = a_Item
        Case 4: a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4)) = a_Item
        Case 5: a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5)) = a_Item
        Case 6: a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6)) = a_Item
        Case 7: a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7)) = a_Item
        Case 8: a_arr(a_ndcs(1), a_ndcs(2), a_ndcs(3), a_ndcs(4), a_ndcs(5), a_ndcs(6), a_ndcs(7), a_ndcs(8)) = a_Item
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
    
    Dim i As Long
    
    If c_coll Is Nothing Then Set c_coll = New Collection
    Select Case True
        Case c_argmnt = Empty:           Coll = c_dflt
        Case IsInteger(c_argmnt) _
            And c_argmnt > c_coll.count: Coll = c_dflt
        Case IsInteger(c_argmnt) _
            And c_argmnt > 1 _
            And c_argmnt <= c_coll.count
            If IsObject(c_coll(c_argmnt)) _
            Then Set Coll = c_coll(c_argmnt) _
            Else Coll = c_coll(c_argmnt)
        Case Else
            '~~ When an argument is provided which is not an integer, the index of
            '~~ the first Item is returned which is identical with the provided argument
            With c_coll
                If IsObject(c_argmnt) Then
                    For i = 1 To .count
                        If .item(i) Is c_argmnt Then
                            Coll = i
                            GoTo xt
                        End If
                    Next i
                Else
                    For i = 1 To .count
                        If .item(i) = c_argmnt Then
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
                                  
    c_dflt = c_dflt
    If c_coll Is Nothing Then Set c_coll = New Collection
    With c_coll
        Select Case True
            Case c_argmnt = Empty
                '~~ When no index is provided the Item is simply added
                .Add c_var
            Case c_argmnt > .count
                '~~ For any indey beyond count elements = Empty are added up to the provided index - 1
                Do While .count < c_argmnt - 1
                    .Add Empty
                Loop
                .Add c_var
            Case c_argmnt <= .count _
              And Not IsObject(c_coll(c_argmnt))
                '~~ Update!
                .Remove c_argmnt
                .Add c_var, c_argmnt
            Case c_argmnt <= c_coll.count _
              And IsObject(c_coll(c_argmnt))
                '~~ Replace value in object
              
        End Select
    End With

End Property

Public Property Let Dict(Optional ByRef d_dct As Dictionary = Nothing, _
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
                                    ReDim arr(0 To .count - 1)
                                    For i = 0 To .count - 1
                                        arr(i) = .item(i + 1)
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
Public Property Get Dict(Optional ByRef d_dct As Dictionary = Nothing, _
                         Optional ByVal d_key As Variant, _
                         Optional ByVal d_arg As Variant = Nothing) As Variant
                         
    If d_dct Is Nothing Then
        If VarType(d_arg) = vbObject Then Set Dict = d_arg Else Dict = d_arg
    Else
        If d_dct.Exists(d_key) Then
            If VarType(d_dct(d_key)) = vbObject Then Set Dict = d_dct(d_key) Else Dict = d_dct(d_key)
        Else
            If VarType(d_arg) = vbObject _
            Then Set Dict = d_arg _
            Else Dict = d_arg
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
' Adds an Item of an array (a_arr), addressed by indices - derived from the key
' (a_key) - indices delimited by commas - to a Dictionary (a_dct) with the
' provided key (a_key). Empty items in the array (a_arr) are ignored.
' ------------------------------------------------------------------------------
    Const PROC = "ArryAsDictAdd"
    
    On Error GoTo eh
    Dim i       As Long
    Dim sKey    As String
    Dim vItem   As Variant
    Dim l       As Long
    
    With a_dct
        For i = LBound(a_arr, a_dim) To UBound(a_arr, a_dim)
            If a_key = vbNullString _
            Then sKey = i _
            Else sKey = a_key & "," & i
            
            If a_dim < a_dims Then
                ArryAsDictAdd a_arr, a_dct, sKey, a_dim + 1, a_dims
            Else
                vItem = ArryItem(a_arr, ArryNdcs(l, Split(sKey, ",")))
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
                            ByVal a_ndcs As Variant, _
                   Optional ByRef a_bnds_out As Collection, _
                   Optional ByRef a_bnds_in As Collection, _
                   Optional ByRef a_out As Long) As Boolean
' ---------------------------------------------------------------------------
' Returns:
' - TRUE when all dimensions addressed by indices (a_ndcs) are
'   within the bounds of the respective dimension in array (a_arr)
' - FALSE when any of the provided indices (a_ndcs) is out of the bounds
'   of the provided array (a_arr)
' - Returns the dimesions which are out of bounds as Collection with items
'   in-bound empty and out-bound with the new bound
' - Returns the complete dimension specifics which combine the "from" spec
'   of the provided array with the new "to" specs in case they are greater
'   than the present ones
'
' Precondition: The indices are provided (a_ndcs) is either as a single
'               integer - when the array (a_arr) is a 1-dim array - or an
'               array of integers, each specifying the index for one
'               dimension.
'
' Uses: Coll
'
' W. Rauschenberger, Berlin Jan 2025
' ---------------------------------------------------------------------------
    Const PROC = "ArryBounds"
    
    On Error GoTo eh
    Dim aBoundsIn(1 To 2)   As Variant
    Dim aBoundsOut(1 To 2)  As Variant
    Dim cllBoundsIn         As New Collection
    Dim cllBoundsOut        As New Collection
    Dim i                   As Long
    Dim lDimsArry           As Long
    Dim lDimsSpec           As Long
    
    lDimsArry = ArryDims(a_arr)
    lDimsSpec = UBound(a_ndcs) - LBound(a_ndcs) + 1
    a_out = 0
    If lDimsSpec > lDimsArry Then GoTo xt ' the number of specified dimesnions is greater then the number of dimensions of the array
    
    For i = 1 To lDimsSpec
        aBoundsIn(1) = Min(a_ndcs(i), LBound(a_arr, i))
        aBoundsIn(2) = Max(a_ndcs(i), UBound(a_arr, i))
        cllBoundsIn.Add aBoundsIn
        If a_ndcs(i) < LBound(a_arr, i) Or a_ndcs(i) > UBound(a_arr, i) Then
            aBoundsOut(1) = Min(a_ndcs(i), LBound(a_arr, i))
            aBoundsOut(2) = Max(a_ndcs(i), UBound(a_arr, i))
            cllBoundsOut.Add aBoundsOut
            a_out = a_out + 1
        Else
            cllBoundsOut.Add Empty
        End If
    Next i
    
    Set a_bnds_in = cllBoundsIn
    Set a_bnds_out = cllBoundsOut
    ArryBounds = cllBoundsOut.count > 0
    Set cllBoundsIn = Nothing
    Set cllBoundsOut = Nothing

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
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
            Case "Byte()", "Integer()", "Long()", "Single()", "Double()", "Currency()": ArryDefault = 0
            Case "Date()":                                                              ArryDefault = #12:00:00 AM#
            Case "String()":                                                            ArryDefault = vbNullString
            Case "Boolean()":                                                           ArryDefault = False
            Case "Variant()":                                                           ArryDefault = Empty
            Case "Object()":                                                            Set ArryDefault = Nothing
            Case Else:                                                                  ArryDefault = Empty
        End Select
    End If
    
End Function

Public Function ArryDiffers(ByVal a_arr1 As Variant, _
                            ByVal a_arr2 As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the an array (a_arr1) differs from another (a_arr2).
' The arrays are compared via an unload to a corresponding Dictionary which
' keeps all items of a max 8-dimms array each Item with a key equal to the
' indices delimited by a comma.
' ----------------------------------------------------------------------------
    Const PROC = "ArryDiffers"
    
    On Error GoTo eh
    Dim dctArr1 As Dictionary
    Dim dctArr2 As Dictionary
    Dim dctDone As New Dictionary
    Dim v       As Variant
    
    Set dctArr1 = ArryAsDict(a_arr1)
    Set dctArr2 = ArryAsDict(a_arr2)

    ArryDiffers = True
    If dctArr1.count <> dctArr2.count Then GoTo xt
    For Each v In dctArr1
        If Not dctArr2.Exists(v) Then GoTo xt
        If dctArr1(v) <> dctArr2(v) Then GoTo xt
        If Not dctDone.Exists(v) Then dctDone.Add v, vbNullString
    Next v
    
    For Each v In dctArr2
        If Not dctDone.Exists(v) Then GoTo xt
    Next v
    
    ArryDiffers = False
    Set dctDone = Nothing
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function ArryDims(ByVal a_arr As Variant) As Long

    Dim i As Long, j As Long
    On Error Resume Next
    Do
        i = i + 1
        j = LBound(a_arr, i)
    Loop While Err.Number = 0
    
    ArryDims = i - 1
    On Error GoTo 0 ' Reset error handling

End Function

Public Function ArrySpecs(ByVal a_arr As Variant, _
                          ByVal a_dims As Long) As Variant

    Dim aBnds   As Variant
    Dim i       As Long
    
    If a_dims > 0 Then
        ReDim aBnds(1 To a_dims, 1 To 2)
        i = 1
        For i = LBound(aBnds, i) To UBound(aBnds, i)
            aBnds(i, 1) = LBound(a_arr, i)
            aBnds(i, 2) = UBound(a_arr, i)
        Next i
        ArrySpecs = aBnds
    End If
    
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

Public Function ArryNdcs(ByRef a_dims As Long, _
                    ParamArray a_ndcs() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns provided indices (a_ndcs) - if any - as array whereby the indices
' may be provided: - as integers (one for each dimension)
'                  - as an array of integers
'                  - as a string of integers delimited by a , (comma)
'                  - as a collection of indices
' ----------------------------------------------------------------------------
    Const PROC = "ArryNdcs"
    
    Dim arr As Variant
    Dim cll As New Collection
    Dim v2  As Variant
    Dim v1  As Variant
    Dim aTmp    As Variant
    Dim i       As Long
    Dim lBase   As Long
    Dim j       As Long
    
    lBase = LBound(Array(1))
    On Error GoTo xt
    If UBound(a_ndcs) >= LBound(a_ndcs) Then
        For Each v1 In a_ndcs
            Select Case True
                Case IsEmpty(v1)
                    GoTo xt
                Case TypeName(v1) = "Collection"    ' An element in the The first Item is a Collection of incides
                    Set cll = v1
                    With cll
                        ReDim arr(1 To .count)
                        For i = 1 To .count
                            If IsInteger(CInt(Trim(.item(i)))) _
                            Then arr(i) = .item(i) _
                            Else Err.Raise AppErr(1), ErrSrc(PROC), "At least one of the items provided as array is not an integer value!"
                        Next i
                    End With
                Case TypeName(v1) = "String"        ' The first Item is a string of incides delimited by a comma
                    aTmp = Split(v1, ",")
                    ReDim arr(1 To UBound(aTmp) - LBound(aTmp) + 1)
                    For i = LBound(aTmp) To UBound(aTmp)
                        If IsNumeric(Trim(aTmp(i))) Then
                            arr(i + 1 - lBase) = Trim(aTmp(i))
                        Else
                            Err.Raise AppErr(2), ErrSrc(PROC), "At least one of the items provided a string delimited by a comma is not an integer value!"
                        End If
                    Next i
                Case IsInteger(v1)                  ' An element of the ParamArray is an integer
                    i = i + 1
                    If i = 1 Then
                        ReDim arr(1 To i)
                    Else
                        ReDim Preserve arr(1 To i)
                    End If
                    arr(i) = v1
                Case IsArray(v1)                    ' The first Item is an array of indices
                    ReDim arr(1 To UBound(v1) - LBound(v1) + 1)
                    For i = LBound(v1) To UBound(v1)
                        If IsNumeric(Trim(v1(i))) Then
                            j = j + 1
                            arr(j) = Trim(v1(i))
                        Else
                            Err.Raise AppErr(2), ErrSrc(PROC), "At least one of the items provided a string delimited by a comma is not an integer value!"
                        End If
                    Next i
            End Select
        Next v1
        ArryNdcs = arr
        On Error Resume Next
        a_dims = UBound(arr) - LBound(arr) + 1
    End If
    
xt: On Error GoTo 0
    
End Function

'Public Function ArryNdcsCll(ParamArray a_ndcs() As Variant) As Collection
'' ----------------------------------------------------------------------------
'' Returns provided indices (a_ndcs) as Collection whereby the indices may
'' be provided: - as integers (one for each dimension)
''              - as an array of integers
''              - as a string of integers delimited by a , (comma).
'' Note: A collection of indices is the preferred type because it eases
''       handling none.
'' ----------------------------------------------------------------------------
'    Const PROC = "ArryIncides"
'
'    Dim cll As New Collection
'
'    Dim v2  As Variant
'    Dim v1  As Variant
'
'    On Error GoTo xt
'    If UBound(a_ndcs) >= LBound(a_ndcs) Then
'        For Each v1 In a_ndcs
'            Select Case True
'                Case IsEmpty(v1)
'                    GoTo xt
'                Case IsArray(v1)
'                    '~~ The first Item is an array if incides
'                    For Each v2 In a_ndcs(LBound(a_ndcs))
'                        If IsInteger(CInt(Trim(v2))) _
'                        Then cll.Add v2 _
'                        Else Err.Raise AppErr(1), ErrSrc(PROC), "At least one of the items provided as array is not an integer value!"
'                    Next v2
'                Case TypeName(v1) = "Collection"
'                    '~~ The first Item is a Collection of indices
'                    Set cll = a_ndcs(LBound(a_ndcs))
'
'                Case TypeName(v1) = "String"
'                    '~~ The first Item is a string of incides delimited by a comma
'                    For Each v2 In Split(v1, ",")
'                        If IsNumeric(v2) Then
'                            cll.Add CInt(Trim(v2))
'                        Else
'                            Err.Raise AppErr(2), ErrSrc(PROC), "At least one of the items provided a string delimited by a comma is not an integer value!"
'                        End If
'                    Next v2
'                Case IsInteger(v1)
'                    cll.Add v1
'            End Select
'        Next v1
'    End If
'
'xt: On Error GoTo 0
'    Set ArryNdcsCll = cll
'    Set cll = Nothing
'
'End Function

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
' Returns the number of items in a multi-dimensional or "jagged" array (i e.
' an Item which again is an array) by optionally counting only "active"
' (a_default_excluded) items - i.e. items which do neither return an Error nor
' the array type specific default.
'
' Notes:
' - Since an unallocated array returns 0 this function is equivalent to the
'   .Count property of a Collection or Dictionary and also may replace the
'   ArryIsAllocated function/service
' - The function also handles type Object arrays where the default
'   "Is Nothing"
'
' Uses: ArryDims    to get rhe number of dimensions
'       ArryDefault to identify default items when default items are to be
'                   excluded.
'
' W. Rauschenberger, Berlin Mar 2025
'  ----------------------------------------------------------------------------
    Dim lItems  As Long
    Dim lDims   As Long
    Dim v       As Variant
    Dim i       As Long
    Dim vDflt   As Variant
    
    If Not IsArray(a_arr) Then GoTo xt
    
    vDflt = ArryDefault(a_arr)
    For Each v In a_arr
        Select Case True
            Case TypeName(v) = "Error"
                If Not a_default_excluded Then
                    lItems = lItems + 1
                End If
            Case IsArray(v) ' "jagged" array!
                lItems = lItems + ArryItems(v, a_default_excluded)
            Case Else
                If VarType(vDflt) = vbObject Then
                    If a_default_excluded Then
                        If Not v Is vDflt Then lItems = lItems + 1
                    Else
                        lItems = lItems + 1
                    End If
                Else
                    If a_default_excluded Then
                         If v <> vDflt Then lItems = lItems + 1
                    Else
                        lItems = lItems + 1
                    End If
                End If
        End Select
    Next v
    
xt: ArryItems = lItems
    
End Function
Public Function ArryNextIndex(ByVal a_arr As Variant, _
                              ByRef a_ndcs() As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns the logical next index for a multidimensional array (a_arr) based
' on a given index (a_ndcs) which is an array of indices, one for each
' dimension in the provided array (a_arr). When the next index would be one
' above all upper bounds the function returns FALSE, else the indices (a_ndcs)
' are the logically next one.
' Precondition: The indices array (a_ndcs) is specified as 1 to x for each
'               dimension.
' ----------------------------------------------------------------------------
    Const PROC = "ArryNextIndex"

    Dim i       As Long
    Dim lDims   As Long
    Dim lNext   As Long

    On Error GoTo xt
    If LBound(a_ndcs) <= LBound(a_ndcs) Then
        lDims = ArryDims(a_arr)
        If UBound(a_ndcs) <> lDims _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided current index does not correctly consider the number of dimensions in the provided array!"
    
        For i = lDims To 1 Step -1
            '~~ Loop through all dimensins from last to first
            lNext = a_ndcs(i) + 1
            If lNext > UBound(a_arr, i) Then
                a_ndcs(i) = LBound(a_arr, i)
            Else
                a_ndcs(i) = lNext
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
' Returns a provided multi-dimensional (max 8) array (a_arr) with new dimension
' specifics (a_dim) whereby the new dimension specs (a_dim) are provided as
' strings following the format: <dimension>:<from>" to "<to> whereby
' <dimension> is either addressing a dimesion in the current array (a_arr),
' i.e. before the Redim has taken place, or a + for a new dimension. Since only
' new or dimensions with changed from/to specs are provided the information
' will be used to compile the final new target array's dimensions - which must
' not exceed 8!.
'
' Uses: ArryReDimSpecs
'
' Requires References to: - Microsoft Scripting Runtime
'                         - Microsoft VBScript Regular Expressions 5.5
' ------------------------------------------------------------------------------
    Const PROC = "ArryReDim"
    
    Dim arrOut          As Variant
    Dim arrUnloaded     As Variant
    Dim dctDimSpecs     As New Dictionary
    Dim i               As Long
    Dim sIndices        As String
    Dim sIndicesPrfx    As String
    Dim l               As Long
    
    On Error GoTo xt
    If UBound(a_arr) >= LBound(a_arr) Then
        On Error GoTo eh
        
        arrUnloaded = ArryUnload(a_arr)             ' unload the input array in a 2-dim array with Item 1 = indices and Item 2 = the input array's Item
        sIndicesPrfx = ArryReDimSpecs(a_arr, a_dim, dctDimSpecs) ' obtain the new and or changed dimension specifics
        
        '~~ Get confirmed that the total number of dimensions not exieeds the maximum suported
        If dctDimSpecs.count > 8 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), _
                       "The target array's number of dimensions resulting from the provided " & _
                       "array and the provided dimesion specifics exeeds the maximum number of 8 dimensions!"
            
        '~~ Redim the new target array considering a maximum of 8 possible dimensions
        Select Case dctDimSpecs.count
            Case 1: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2))
            Case 2: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2))
            Case 3: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2) _
                               , dctDimSpecs(3)(1) To dctDimSpecs(3)(2))
            Case 4: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2) _
                               , dctDimSpecs(3)(1) To dctDimSpecs(3)(2) _
                               , dctDimSpecs(4)(1) To dctDimSpecs(4)(2))
            Case 5: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2) _
                               , dctDimSpecs(3)(1) To dctDimSpecs(3)(2) _
                               , dctDimSpecs(4)(1) To dctDimSpecs(4)(2) _
                               , dctDimSpecs(5)(1) To dctDimSpecs(5)(2))
            Case 6: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2) _
                               , dctDimSpecs(3)(1) To dctDimSpecs(3)(2) _
                               , dctDimSpecs(4)(1) To dctDimSpecs(4)(2) _
                               , dctDimSpecs(5)(1) To dctDimSpecs(5)(2) _
                               , dctDimSpecs(6)(1) To dctDimSpecs(6)(2))
            Case 7: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2) _
                               , dctDimSpecs(3)(1) To dctDimSpecs(3)(2) _
                               , dctDimSpecs(4)(1) To dctDimSpecs(4)(2) _
                               , dctDimSpecs(5)(1) To dctDimSpecs(5)(2) _
                               , dctDimSpecs(6)(1) To dctDimSpecs(6)(2) _
                               , dctDimSpecs(7)(1) To dctDimSpecs(7)(2))
            Case 8: ReDim arrOut(dctDimSpecs(1)(1) To dctDimSpecs(1)(2) _
                               , dctDimSpecs(2)(1) To dctDimSpecs(2)(2) _
                               , dctDimSpecs(3)(1) To dctDimSpecs(3)(2) _
                               , dctDimSpecs(4)(1) To dctDimSpecs(4)(2) _
                               , dctDimSpecs(5)(1) To dctDimSpecs(5)(2) _
                               , dctDimSpecs(6)(1) To dctDimSpecs(6)(2) _
                               , dctDimSpecs(7)(1) To dctDimSpecs(7)(2) _
                               , dctDimSpecs(8)(1) To dctDimSpecs(8)(2))
        End Select
        
        '~~ Re-load the unloaded array to the new re-dimed array and return it replacing the source array
        For i = LBound(arrUnloaded) To UBound(arrUnloaded)
            sIndices = sIndicesPrfx & arrUnloaded(i, 1)
            On Error Resume Next
            ArryItem(arrOut, ArryNdcs(l, sIndices)) = arrUnloaded(i, 2)
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

Private Function ArryReDimSpecs(ByVal a_arr As Variant, _
                                ByVal a_specs As Variant, _
                                ByRef a_dims As Dictionary) As String
' ------------------------------------------------------------------------------
' Returns
' 1. a Dictionary (a_dims) with the redim specifics by considering the
'    current array's (a_arr) dimension specs plus those provided new (a_specs).
'    The syntax of the new dimension specs is: <dimension>:<from>" to "<to>
'    whereby <dimension> is either "+" for a new one or an integer identifying a
'    dimension of which the specs will be changed.
' 2. The prefix for the re-load which is the new specified dimensions lower
'    bound.
' ------------------------------------------------------------------------------
    Const PROC                  As String = "ArryReDimSpecs"
    '~~ Syntax constants for the a_specs argument
    Const NEW_DIM_INDICATOR     As String = "+"
    Const DIM_SPEC_DELIMITER    As String = ":"
    Const DIM_BOUNDS_DELIMITER  As String = "to"
    
    On Error GoTo eh
    Dim aBounds     As Variant
    Dim lBase       As Long
    Dim lDims       As Long
    Dim lDimsNew    As Long
    Dim i           As Long
    Dim regex       As Object
    Dim sSpec       As String
    Dim v           As Variant
    Dim lDim        As Long
    Dim sNewDims    As String
    
    lBase = LBound(Array("x"))
    lDims = ArryDims(a_arr)
    
    Set regex = CreateObject("VBScript.RegExp")
    If Not ArryIsAllocated(a_specs) Then GoTo xt
    ReDim aBounds(1 To 2)
    
    '~~ 1. Check wether all dimension specifcation conform with the considered syntax
    For Each v In a_specs
        With regex
            ' ------------------------------------------------------------------------------
            ' Syntax for each element in a_specs:
            ' Starts with a "+" (NEW_DIM_INDICATOR) or an integer followed by none or mor spaces,
            ' a ":" (DIM_SPEC_DELIMITER) followed by none or more spaces,
            ' the LBound integer followed by none or more spaces
            ' the "to" delimiter (DIM_BOUNDS_DELIMITER), followed by none or more spaces
            ' the UBound integer
            ' ------------------------------------------------------------------------------
'            .Pattern = "^\s*(\+|\d+)\s*:\s*\d+\s*to\s*\d+\s*$"
            .Pattern = "^\s*(\" & NEW_DIM_INDICATOR & "|\d+)\s*" & DIM_SPEC_DELIMITER & "\s*\d+\s*" & DIM_BOUNDS_DELIMITER & "\s*\d+\s*$"
            .IgnoreCase = False
            .Global = False
            If Not .Test(v) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), _
                           "The provided dimension specification does not conform with expectations!" & vbLf & _
                           "A valid spec starts with a ""+"" (plus) or an integer from 1 to 8 followed " & _
                           "by a "":"" (semicolon), followed by to integers delimited by a ""to"". Any spaces are optional."
        End With
    Next v
    
    With a_dims
        '~~ 2. Collect any new dimensions (those indicated by a +)
        For Each v In a_specs
            If Trim$(Split(v, DIM_SPEC_DELIMITER)(0)) = NEW_DIM_INDICATOR Then
                '~~ A new dimension is indicated by the + sign
                lDimsNew = lDimsNew + 1
                sSpec = Trim$(Split(v, DIM_SPEC_DELIMITER)(1))
                aBounds(1) = CInt(Trim$(Split(sSpec, DIM_BOUNDS_DELIMITER)(0))) ' The new dimension's lower bound
                aBounds(2) = CInt(Trim$(Split(sSpec, DIM_BOUNDS_DELIMITER)(1))) ' The new dimension's upper bound
                .Add lDimsNew, aBounds
                sNewDims = sNewDims & aBounds(1) & ","
            End If
        Next v
            
        '~~ 3. Add any changed dimension specifics/bounds for existing dimensions (those with an integer preceeding the :)
        For Each v In a_specs
            Debug.Print v
            If IsInteger(Trim$(Split(v, DIM_SPEC_DELIMITER)(0))) Then
                '~~ The change of an existing dimension is indicated by an integer representing the dimension
                lDim = Trim$(Split(v, DIM_SPEC_DELIMITER)(0))
                sSpec = Trim(Split(v, DIM_SPEC_DELIMITER)(1))
                aBounds(1) = CInt(Trim$(Split(sSpec, DIM_BOUNDS_DELIMITER)(0)))
                aBounds(2) = CInt(Trim$(Split(sSpec, DIM_BOUNDS_DELIMITER)(1)))
                .Add lDim + lDimsNew, aBounds
            End If
        Next v
    
        '~~ 4. Collect any yet not considered dimension specifics from the source array
        For i = 1 To lDims
            If Not a_dims.Exists(i) Then
                '~~ The dimensions specifics had yet not been collected
                aBounds(1) = LBound(a_arr, i)
                aBounds(2) = UBound(a_arr, i)
                .Add i + lDimsNew, aBounds 'Note: The 1st dim will become the 2nd dim when 1 new had been specified
            End If
        Next i
    End With
    Debug.Print sNewDims
    ArryReDimSpecs = sNewDims
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Sub ArryRemoveItems(ByRef a_arry As Variant, _
                  Optional ByVal a_Item As Long = -99999, _
                  Optional ByVal a_indx As Long = -99999, _
                  Optional ByVal a_items As Long = 1)
' ------------------------------------------------------------------------------
' Returns arr 1-dim array (a_arry) with the number of elements (a_items) removed
' whereby the start element may be indicated by the Item number 1,2,.. (a_Item)
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
            
            Case a_Item = -99999 And a_indx <> -99999
                If a_indx < lLbnd Or a_indx > lUbnd _
                Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided index is out of the array's bounds!"
            
            Case a_Item >= 0 And a_indx = -99999
                '~~ When an Item is provided which is within the number of items in the array
                '~~ it is transformed into an index
                If a_Item < 1 Or a_Item > lItems _
                Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided Item number is 0 or exceeds the number of items in the array!"
                ' lBnd = 1 And Item = 4 > index = Item
                ' lBnd = 4 And Item = 2 > index = 5
                a_indx = lLbnd + a_Item - 1
            
            Case Else
                Err.Raise AppErr(4), ErrSrc(PROC), "Neither an Item nor an index had been provided or both which is conflicting!"
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
    Dim j   As Long
    Dim arr As Variant
    
    On Error GoTo xt
    If UBound(a_arr) >= LBound(a_arr) Then
        j = LBound(a_arr)
        ReDim arr(UBound(a_arr))
        For i = LBound(a_arr) To UBound(a_arr)
            Select Case True
                Case IsError(a_arr(i))
                Case a_arr(i) = Empty
                Case Trim$(a_arr(i)) = vbNullString
                Case Else
                    If Not Len(Trim(a_arr(i))) = 0 Then
                        arr(j) = Trim(a_arr(i))
                        j = j + 1
                    End If
            End Select
        Next i
    End If
    ReDim Preserve arr(j - 1)
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
' into a "flat" array. Any inactive elements/items (IsError or = default) are
' ignored.
' ------------------------------------------------------------------------------
    Const PROC = "ArryUnload"
    
    On Error GoTo eh
    Dim i           As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, p As Long, z As Long
    Dim arr         As Variant
    Dim vDflt       As Variant
    Dim lDims       As Long
    Dim lTmp        As Long
    
    lDims = ArryDims(a_arr)
    vDflt = ArryDefault(a_arr)
    ReDim arr(1 To ArryItems(a_arr, True), 1 To 2) ' ensure the output array has as many itmes as are active!
    
    Select Case lDims
        Case 1: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    If ArryItemIsActive(a_arr, vDflt, i) Then
                        z = z + 1
                        arr(z, 1) = i
                        arr(z, 2) = a_arr(i)
                    End If
                Next i
        Case 2: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        If ArryItemIsActive(a_arr, vDflt, i, j) Then
                            z = z + 1
                            arr(z, 1) = i & "," & j
                            arr(z, 2) = a_arr(i, j)
                        End If
                    Next j
                Next i
        Case 3: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                            If ArryItemIsActive(a_arr, vDflt, i, j, k) Then
                                z = z + 1
                                arr(z, 1) = i & "," & j & "," & k
                                arr(z, 2) = a_arr(i, j, k)
                            End If
                        Next k
                    Next j
                Next i
        Case 4: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                            For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                If ArryItemIsActive(a_arr, vDflt, i, j, k, l) Then
                                    z = z + 1
                                    arr(z, 1) = i & "," & j & "," & k & "," & l
                                    arr(z, 2) = a_arr(i, j, k, l)
                                End If
                            Next l
                        Next k
                    Next j
                Next i
        Case 5: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                            For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                    If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m) Then
                                        z = z + 1
                                        arr(z, 1) = i & "," & j & "," & k & "," & l & "," & m
                                        arr(z, 2) = a_arr(i, j, k, l, m)
                                    End If
                                Next m
                            Next l
                        Next k
                    Next j
                Next i
        Case 6: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                            For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                    For n = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m, n) Then
                                            z = z + 1
                                            Arry(arr, ArryNdcs(l, z, 1)) = i & "," & j & "," & k & "," & l & "," & m & "," & n
                                            Arry(arr, ArryNdcs(l, z, 2)) = a_arr(i, j, k, l, m, n)
                                        End If
                                    Next n
                                Next m
                            Next l
                        Next k
                    Next j
                Next i
        Case 7: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                            For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                    For n = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        For o = LBound(a_arr, 5) To UBound(a_arr, 5)
                                            If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m, n, o) Then
                                                z = z + 1
                                                arr(z, 1) = i & "," & j & "," & k & "," & l & "," & m & "," & n & "," & o
                                                arr(z, 2) = a_arr(i, j, k, l, m, n, o)
                                            End If
                                        Next o
                                    Next n
                                Next m
                            Next l
                        Next k
                    Next j
                Next i
        Case 8: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                    For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                        For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                            For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                    For n = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        For o = LBound(a_arr, 5) To UBound(a_arr, 5)
                                            For p = LBound(a_arr, 5) To UBound(a_arr, 5)
                                                If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m, n, o, p) Then
                                                    z = z + 1
                                                    arr(z, 1) = i & "," & j & "," & k & "," & l & "," & m & "," & n & "," & o & "," & p
                                                    arr(z, 2) = a_arr(i, j, k, l, m, n, o, p)
                                                End If
                                            Next p
                                        Next o
                                    Next n
                                Next m
                            Next l
                        Next k
                    Next j
                Next i
    End Select
    ArryUnload = arr
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function ArryUnloadToDict(ByVal a_arr As Variant, _
                        Optional ByRef a_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Unloads a multi-dimensional (max 8 dims) array (a_arr) to a Dictionary with
' the Item's indices delimited by a comma as key.
' ------------------------------------------------------------------------------
    Const PROC = "ArryUnloadToDict"
    
    On Error GoTo eh
    Dim i       As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, p As Long
    Dim sKey    As String
    Dim lDims   As Long
    Dim vDflt   As Variant

    If a_dct Is Nothing Then Set a_dct = New Dictionary Else a_dct.RemoveAll
    lDims = ArryDims(a_arr)
    vDflt = ArryDefault(a_arr)

    ' Use With...End With for performance
    With a_dct ' Intentionally for performance, i.e. also no recursive solution by intention
        Select Case lDims
            Case 1: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        If ArryItemIsActive(a_arr, vDflt, i) Then
                            sKey = i
                            .Add sKey, a_arr(i)
                        End If
                    Next i
            Case 2: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            If ArryItemIsActive(a_arr, vDflt, i, j) Then
                                sKey = i & "," & j
                                .Add sKey, a_arr(i, j)
                            End If
                        Next j
                    Next i
            Case 3: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                                If ArryItemIsActive(a_arr, vDflt, i, j, k) Then
                                     sKey = i & "," & j & "," & k
                                    .Add sKey, a_arr(i, j, k)
                                End If
                            Next k
                        Next j
                    Next i
            Case 4: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                                For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                    If ArryItemIsActive(a_arr, vDflt, i, j, k, l) Then
                                        sKey = i & "," & j & "," & k & "," & l
                                        .Add sKey, a_arr(i, j, k, l)
                                    End If
                                Next l
                            Next k
                        Next j
                    Next i
            Case 5: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                                For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                    For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m) Then
                                            sKey = i & "," & j & "," & k & "," & l & "," & m
                                            .Add sKey, a_arr(i, j, k, l, m)
                                        End If
                                    Next m
                                Next l
                            Next k
                        Next j
                    Next i
            Case 6: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                                For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                    For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        For n = LBound(a_arr, 5) To UBound(a_arr, 5)
                                            If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m, n) Then
                                                sKey = i & "," & j & "," & k & "," & l & "," & m & "," & n
                                                .Add sKey, a_arr(i, j, k, l, m, n)
                                            End If
                                        Next n
                                    Next m
                                Next l
                            Next k
                        Next j
                    Next i
            Case 7: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                                For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                    For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        For n = LBound(a_arr, 5) To UBound(a_arr, 5)
                                            For o = LBound(a_arr, 5) To UBound(a_arr, 5)
                                                If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m, n, o) Then
                                                    sKey = i & "," & j & "," & k & "," & l & "," & m & "," & n & "," & o
                                                    .Add sKey, a_arr(i, j, k, l, m, n, o)
                                                End If
                                            Next o
                                        Next n
                                    Next m
                                Next l
                            Next k
                        Next j
                    Next i
            Case 8: For i = LBound(a_arr, 1) To UBound(a_arr, 1)
                        For j = LBound(a_arr, 2) To UBound(a_arr, 2)
                            For k = LBound(a_arr, 3) To UBound(a_arr, 3)
                                For l = LBound(a_arr, 4) To UBound(a_arr, 4)
                                    For m = LBound(a_arr, 5) To UBound(a_arr, 5)
                                        For n = LBound(a_arr, 5) To UBound(a_arr, 5)
                                            For o = LBound(a_arr, 5) To UBound(a_arr, 5)
                                                For p = LBound(a_arr, 5) To UBound(a_arr, 5)
                                                    If ArryItemIsActive(a_arr, vDflt, i, j, k, l, m, n, o, p) Then
                                                        sKey = i & "," & j & "," & k & "," & l & "," & m & "," & n & "," & o & "," & p
                                                        .Add sKey, a_arr(i, j, k, l, m, n, o, p)
                                                    End If
                                                Next p
                                            Next o
                                        Next n
                                    Next m
                                Next l
                            Next k
                        Next j
                    Next i
        End Select
    End With

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Function ArryItemIsActive(ByVal a_arr As Variant, _
                                 ByVal a_dflt As Variant, _
                            ParamArray a_indcs() As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Item in an array (a_arr) identified by indices (a_indcs)
' neither is an error nor is it identical with a default (a_dflt).
' ------------------------------------------------------------------------------
    Const PROC = "ArryItemIsActive"
    
    On Error GoTo eh
    Dim vItem   As Variant
    Dim i       As Long
    
    ' Check if the correct number of a_indcs is provided
    If UBound(a_indcs) + 1 > 8 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The number of indices exceeds the maximum supported which is 8."

    '~~ Retrieve the vItem using the provided a_indcs
    On Error GoTo xt ' returns False
    vItem = a_arr
    For i = LBound(a_indcs) To UBound(a_indcs)
        vItem = vItem(a_indcs(i))
    Next i
    On Error GoTo eh
    
    If VarType(a_dflt) <> vbObject Then
        If vItem <> a_dflt Then ArryItemIsActive = True
    Else
        If Not vItem Is a_dflt Then ArryItemIsActive = True
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

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
    Const PROC = "DictAsArray"
    
    On Error GoTo eh
    Dim lDims       As Long
    Dim arrIndices  As Variant
    Dim arr         As Variant
    Dim arrFrom     As Variant
    Dim arrTo       As Variant
    Dim v           As Variant
    Dim l           As Long
    
    If d_dct Is Nothing Then GoTo xt
    If d_dct.count = 0 Then GoTo xt
      
    If d_dimsfrom Is Nothing And d_dimsto Is Nothing Then
        '~~ When no new dimension specs are provided they are retrieved from the dictionar key
        arrFrom = Split(d_dct.Keys()(0), ",")
        lDims = UBound(arrFrom) + 1
        arrTo = Split(d_dct.Keys(d_dct.count - 1), ",")
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
        ArryItem(arr, ArryNdcs(l, arrIndices)) = d_dct(v)
    Next v
    DictAsArry = arr

xt: Exit Function
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DictTest()

    Dim dct As Dictionary
    
    Dict(dct, "X") = 1              ' Add
    Dict(dct, "X") = 2              ' Replace
    Debug.Assert Dict(dct, "X") = 2 ' Assert
    
    Dict(dct, "X") = "+5"           ' Increment
    Debug.Assert Dict(dct, "X") = 7 ' Assert
    
    Dict(dct, "X") = "-3"           ' Decrement
    Debug.Assert Dict(dct, "X") = 4 ' Assert
    
    '~~ Assert defaults when not existing
    Debug.Assert Dict(dct, "B", Nothing) Is Nothing
    Debug.Assert Dict(dct, "B", 0) = 0
    Debug.Assert Dict(dct, "B") = vbNullString      ' No default returns default
    
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
    On Error GoTo xt
    Select Case True
        Case i_value = 0
            IsInteger = True
        Case IsNumeric(i_value)
            IsInteger = (CLng(i_value) = Int(i_value))
    End Select
    
xt: On Error GoTo 0
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
    If s_dct.count = 0 Then GoTo xt
    
    With s_dct
        ReDim arr(0 To .count - 1)
        For i = 0 To .count - 1
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
            .Add key:=vKey, item:=s_dct.item(vKey)
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
'             single Item which again is an Array or a Collection, an error
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
                                    Err.Raise AppErr(1), ErrSrc(PROC), "Single argument array with single Item array (jagged array) not supported!"
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
                Case IsNumeric(v1)
                    If v1 > vMax1 Then vMax1 = v1
                Case VarType(v1) = vbString
                    If Len(v1) > vMax1 Then vMax1 = Len(v1)
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

Public Function RepeatString(ByVal r_str As String, _
                             ByVal r_cnt As Integer) As String
' ----------------------------------------------------------------------------
' Returns a string (r_str) repeated n (r_cnt) times.
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim s As String

    For i = 1 To r_cnt
        s = s & r_str
    Next i
    RepeatString = s
    
End Function

Public Function RngeAsArry(ByVal r_rng As Range) As Variant
' ----------------------------------------------------------------------------
' Returns a range (r_rng) as arry.
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
        If Chr$(a(i)) = " " Then
            Spaced = Spaced & Chr$(160) & Chr$(160)
        Else
            Spaced = Spaced & Chr$(160) & Chr$(a(i))
        End If
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
        Case s_lvl <> 0 And s_lvl > s_stck.count
            GoTo xt

        Case s_lvl = 0 And VarType(s_Item) <> vbObject
            If s_Item = -999999999 Then GoTo xt
            '~~ A specific Item has been provided
            For i = 1 To s_stck.count
                If s_stck(i) = s_Item Then
                    s_lvl = i
                    StackEd = True
                    GoTo xt
                End If
            Next i
        
        Case s_lvl = 0 And VarType(s_Item) = vbObject
            For i = 1 To s_stck.count
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
    If Not StackIsEmpty Then StackIsEmpty = stck.count = 0
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
    
    If IsObject(stck(stck.count)) _
    Then Set StackPop = stck(stck.count) _
    Else StackPop = stck(stck.count)
    stck.Remove stck.count

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Public Sub StackPush(ByRef stck As Collection, _
                     ByVal stck_Item As Variant)
' ----------------------------------------------------------------------------
' Common Stack Push service. Pushes (adds) an Item (stck_Item) to the stack
' (stck). When the provided stack (stck) is Nothing the stack is created.
' ----------------------------------------------------------------------------
    Const PROC = "StckPush"
    
    On Error GoTo eh
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_Item

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
    If IsObject(stck(stck.count)) _
    Then Set StackTop = stck(stck.count) _
    Else StackTop = stck(stck.count)

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

