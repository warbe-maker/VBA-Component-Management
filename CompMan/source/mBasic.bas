Attribute VB_Name = "mBasic"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mBasic: Declarations, procedures, methods and function
' ======================= likely to be required in any VB-Project, optionally
' just being copied.
'
' Note: All provided public services run completely autonomous. I.e. they do
'       not require any other installed module. However, when the Common VBA
'       Message Services (fMsg/mMsg) and or the Common VBA Error Services
'       (mErH) are installed an error message is passed on to their corres-
'       ponding service which provides a much better designed error message.
'
' Public Procedures and Properties (r/w):
' ---------------------------------------
' Align               Returns a provided string in a specified length, with
'                     optional margins (left and right) which defaults to
'                     none, aligned left, right or centered, with an optional
'                     fill string which defaults to spaces.
'                     Specifics:
'                     - When a margin is provided, the final length will be
'                       the specified length plus the length of a left and a
'                       right margin. A margin is typically used when the
'                       string is aligned as an item of serveral items
'                       arranged in columns when the column delimiter is a
'                       vbNullString. When the column delimiter is a | a
'                       marign of a single space is the default *).
'                     - The provided string may contain leading or trailing
'                       spaces. Leading spaces are preserved when the string
'                       is left aligned, trailing spaces are preserved when
'                       the string is aligned right. In any other case
'                       leading and trailing spaces are unstripped.
'                     - The function is also used to align items arragend in
'                       columns.
'
'                     *) Column arranged option = TRUE (defaults to False):
'                        - The provided length is regarded the maximum. I.e.
'                          when the provided string is longer it is truncated
'                          to the right
'                        - The final result string has any specified margin
'                          (left and right) added
'                        - When a fill is specified the final string has at
'                          this one added. For example when the fill string
'                          is " -", the margin is a single space and the
'                          alignment is left, a string "xxx" is returned as
'                          " xxx -------- "
'                          a string "xxxxxxxxxx" is returned as :
'                          " xxxxxxxxxx - "
'                        Column arranged option = FALSE (the default):
'                        - The provided length is the final length returned.
'                        - Any specified margin is ignored
'                        - A pecified fill is added only to end up with the
'                          specidied length
' AlignCntr           Called by Align or directly
' AlignLeft           Called by Align or directly
' AlignRght           Called by Align or directly
' AppErr              Converts a positive error number into a negative to
'                     ensures an error number not conflicting with a VB
'                     run time error or any other system error number.
'                     Returns the origin positive error number when called
'                     with the negative Application Error number. 3)
' AppIsInstalled      Returns TRUE when a named exec is found in the system
'                     path.
' Arry            r/w Universal read/write array service.
'                     r:  Returns a provided default when a given array is
'                         not allocated or a provided index is beyond/
'                         outside the array's number of items (the default
'                         defaults to vbNullString).
'                     w: - Adds an item (c_var) to an array (c_arr) when
'                          no index is provided or adds it with the
'                          provided index
'                        - When an index is provided, the item is
'                          inserted/updated at the given index, even when
'                          the array yet doesn't exist or yet is not
'                          allocated.
' ArryAsRange         Transferres the content of a one- or two-dimensional
'                     array to a range
' ArryBase            Returns the component's actual Base Option.
' ArryCompare         Returns a Dictionary with the provided number of items
'                     (defaults to all) which differ between two one.
'                     dimensional arrays. When no difference is encountered
'                     the returned Dictionary is empty (Count = 0).
' ArryDiffers         Returns TRUE when a provided array differs from
'                     annother.
' ArryDims            Returns the number of dimensions of an array. An
'                     unallocated dynamic array returns 0 dimensions.
' ArryIsAllocated     Returns TRUE when the provided array has at least one
'                     item
' ArryItems           Returns the number of items in a multi-dimensional
'                     array or a nested array. The latter is an array of
'                     which one or more items again are arrays, possibly
'                     multi-dimensional. An unallocated array returns 0.
' ArryRemoveItem      Removes an array's item by its index or element number.
' ArryTrim            Removes any leading or trailing empty items.
' BaseName            Returns the file name of a provided argument without
'                     the extension whereby the  argument may be a file or a
'                     file's name or full name.
' CleanTrim           Clears a string from any unprinable characters.
' DelayedAction       Waits for a specified time and performs a provided
'                     action whereby the action needs to be a fully specified
'                     Application.Run action <workbook-name>!<comp>.<proc>
'                     Specific: The action may be passed with up to 5
'                     arguments. Though optional none preceeding must be
'                     ommitted. I.e. when the second argument is provided the
'                     first one is obligatory.
' Dict            r/w Universal Dictionary service
'                     w: with option add (default), replace, increment,
'                        collect, and collect sorted.
'                     r: Supports a default returned when the provided key
'                        (or the Dictionary does not exist).

' KeySort             Returns a given Dictionary sorted ascending by key.
' README              Displays the Common Component's README in the public
'                     GitHub repo.
' ShellRun            Opens a folder, an email-app, a url, an Access instance,
'                     etc.
' TimedDoEvents       Performs a DoEvent by taking the elapsed time printed
'                     in VBE's immediate window
' TimerBegin          Starts a timer (counting system ticks)
' TimerEnd            Returns the elapsed system ticks converted to
'                     milliseconds
'
' Requires:
' ---------
' Reference to "Microsoft Scripting Runtime"
' Reference to "Microsoft Visual Basic Application Extensibility .."
'
' W. Rauschenberger, Berlin Oct 2024
' See https://github.com/warbe-maker/VBA-Basics (with README servie)
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

' Basic declarations potentially uesefull in any VB-Project
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

' Timer means
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

'Functions to get DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const LOGPIXELSX = 88               ' Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72  ' A point is defined as 1/72 inches
Private Declare PtrSafe Function GetForegroundWindow _
  Lib "User32.dll" () As Long

Private Declare PtrSafe Function GetWindowLongPtr _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As LongPtr, _
     ByVal nIndex As Long) _
  As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As LongPtr, _
     ByVal nIndex As LongPtr, _
     ByVal dwNewLong As LongPtr) _
  As LongPtr

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
Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

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
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Public Property Get Arry(Optional ByRef a_arr As Variant, _
                         Optional ByVal a_index As Long = -1, _
                         Optional ByVal a_default As Variant = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Common array read aervice. Returns the item of a given array (c_arr) at a
' given index. When the array is not allocated or the index is outside the
' array's current boundaries a default is returned - which, when none had been
' specified, defaults to vbNullString.
'
' W. Rauschenberger Berlin, Aug 2024
' ----------------------------------------------------------------------------
    Const PROC = "Arry-Get"
    
    Dim i As Long
    
    If IsArray(a_arr) Then
        On Error Resume Next
        i = LBound(a_arr)
        If Err.Number = 0 Then
            If a_index >= LBound(a_arr) And a_index <= UBound(a_arr) _
            Then Arry = a_arr(a_index)
        Else
            Arry = a_default
        End If
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument (a_arr) is not an array!"
    End If
    
End Property

Public Property Let Arry(Optional ByRef a_arr As Variant, _
                         Optional ByVal a_index As Long = -9999, _
                         Optional ByVal a_default As Variant = vbNullString, _
                                  ByVal a_var As Variant)
' ----------------------------------------------------------------------------
' Common array write service. Returns an array (a_arr) with an item (c_var)
' - either simply added or
' - by having replaced an item at a given index
' - by adding the item at a given index when UBound is less than the provided
'   index.
'
' W. Rauschenberger Berlin, Aug 2024
' ----------------------------------------------------------------------------
    Const PROC = "Arry-Let"
    
    Dim bIsAllocated As Boolean
    Dim s            As String
    
    If IsArray(a_arr) Then
        On Error GoTo -1
        On Error Resume Next
        bIsAllocated = UBound(a_arr) >= LBound(a_arr)
        On Error GoTo eh
    ElseIf VarType(a_arr) <> 0 Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Not a Variant type!"
    End If
    
    If bIsAllocated = True Then
        '~~ The array has at least one item
        If a_index = -9999 Then
            '~~ When for an allocated array no index is provided, the item is added
            ReDim Preserve a_arr(UBound(a_arr) + 1)
            a_arr(UBound(a_arr)) = a_var
        ElseIf a_index >= 0 And a_index <= UBound(a_arr) Then
            '~~ Replace an existing item
            a_arr(a_index) = a_var
        ElseIf a_index > UBound(a_arr) Then
            '~~ New item beyond current UBound
            ReDim Preserve a_arr(a_index)
            a_arr(a_index) = a_var
        ElseIf a_index < LBound(a_arr) Then
            Err.Raise AppErr(2), ErrSrc(PROC), "Index is less than LBound of array!"
        End If
        
    ElseIf bIsAllocated = False Then
        '~~ The array does yet not exist
        If a_index = -9999 Then
            '~~ When no index is provided the item is the first of a new array
            '~~ which by the way considers the Base Option of the component in
            '~~ which the service is used.
            a_arr = Array(a_var)
        ElseIf a_index >= 0 Then
            '~~ Even when an index 0 is provided and the Base Option is 1 the
            '~~ LBound of the returned arry will be 1
            ReDim a_arr(a_index)
            a_arr(a_index) = a_var
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

' ----------------------------------------------------------------------------
' Universal Dictionary add, replace, increase item value service.
' Specific: when the item is a string in the form "-n" or "+n" whereby n is a
' numeric value, the item with the corresponding key is incremented or
' decremented provided it exists or it is added with the negative or positive
' value as item.
' - Get/Read:       Returns a default when an item is not existing
' - Let/Write:      Handles the provided item in correspondence with the
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

Public Property Let Dict(Optional ByRef d_dct As Dictionary = Nothing, _
                         Optional ByVal d_key As Variant, _
                         Optional ByVal d_arg As Variant = enAdd, _
                                  ByVal d_item As Variant)
    Const PROC = "Dict_Let"
    
    Dim arr As Variant
    Dim cll As Collection
    Dim i   As Long
    Dim j   As Long
    Dim tmp As Variant

    If d_dct Is Nothing Then Set d_dct = New Dictionary
    Select Case d_arg
        Case enAdd:         If Not d_dct.Exists(d_key) Then d_dct.Add d_key, d_item
        Case enCollect:
                            If Not d_dct.Exists(d_key) Then
                                Set cll = New Collection
                                cll.Add d_item
                                d_dct.Add d_key, cll
                            Else
                                If Not TypeName(d_dct(d_key)) = "Collection" _
                                Then Err.Raise AppErr(2), ErrSrc(PROC), "The option ""collect"" is not applicable since the existing item is not a Collection!"
                                Set cll = d_dct(d_key)
                                cll.Add d_item
                                d_dct.Remove d_key
                                d_dct.Add d_key, cll
                            End If
        Case enCollectSorted
                            If Not d_dct.Exists(d_key) Then
                                Set cll = New Collection
                                cll.Add d_item
                                d_dct.Add d_key, cll
                            Else
                                If Not TypeName(d_dct(d_key)) = "Collection" _
                                Then Err.Raise AppErr(3), ErrSrc(PROC), "The option ""collect"" is not applicable since the existing item is not a Collection!"
                                Set cll = d_dct(d_key)
                                cll.Add d_item
                                
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
                                d_dct.Add d_key, d_item
                            Else
                                If Not IsNumeric(d_dct(d_key)) _
                                Then Err.Raise AppErr(1), ErrSrc(PROC), "The option ""increment"" is not applicable since the item is not numeric!"
                                tmp = d_dct(d_key)
                                tmp = tmp + d_item
                                d_dct.Remove d_key
                                d_dct.Add d_key, tmp
                            End If
        Case enReplace:     If d_dct.Exists(d_key) Then d_dct.Remove d_key: d_dct.Add d_key, d_item
    End Select
    
End Property

Public Function SysFrequency(Optional ByVal s_f As Currency = 0) As Currency
    If s_f = 0 Then s_f = TimerSystemFrequency
    If s_f = 0 Then getFrequency TimerSystemFrequency
    SysFrequency = TimerSystemFrequency
End Function

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
'   when the string is one item of serveral organized in columns when the
'   column delimiter is a vbNullString. When the column delimiter is a |
'   usually a marign of a single space is used
' - The string (a_strn) may contain already leading or trailing spaces of
'   the left are preserved when the string is left aligned and the right
'   spaces are preserved when the string is right aligned.
'   spaces.
' - The function is also used to align items arragend in columns.
'
' Attention to column arranged alignment (a_col_arranged = True):
' ---------------------------------------------------------------
' - The length is regarded the maximum. I.e. when the string (a_string) is
'   longer it is  truncated to the right
' - The final result string has any margin (left and right) added
' - The final string has of any fill <> vbNullString at least one fill
'   (a_fill) added, in the example below this is " -"
'   Example: A string xxx with given length of 10, a single space margin and
'            a " -" fill, left aligned results in: " xxx -------- " which is
'            a final lenght of 14.
'
' Attention to normal alignment (a_col_arranged = True):
' ------------------------------------------------------
' - The provided length is the final length returned. I.e. when the provided
'   string (a_string) exceeds the length (a_length) the result is truncated
'   to the right
' - Any specified margin is ignored
' - Fills (a_fill) are added to the right to the extent the length of the
'   provided string (a_string) allows it without exceeding the final length
'   (a_lenght).
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
' may be extended to the width of the string (a_steing). The result may not be
' centered exactly.
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
' - enclose the result in left and right margins - which default to a vbNullString
'   when a_col_arranged = True
' Specifics:
' - Left spaces provided with the string (a_string) are preserved
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

Public Sub ArryAsRange(ByVal a_arr As Variant, _
                       ByVal a_rng As Range, _
              Optional ByVal a_one_col As Boolean = False)
' ----------------------------------------------------------------------------
' Transferes the content of an array (vArr) into the range (a_rng).
' ----------------------------------------------------------------------------
    Const PROC = "ArryAsRange"
    
    On Error GoTo eh
    Dim rTarget As Range

    If a_one_col Then
        '~~ One column, n rows
        Set rTarget = a_rng.Cells(1, 1).Resize(UBound(a_arr) + 1, 1)
        rTarget.Value = Application.Transpose(a_arr)
    Else
        '~~ One column, n rows
        Set rTarget = a_rng.Cells(1, 1).Resize(1, UBound(a_arr) + 1)
        rTarget.Value = a_arr
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Function ArryBase() As Long
' ----------------------------------------------------------------------------
' Returns the component's actual Base Option.
' ----------------------------------------------------------------------------
    ArryBase = LBound(Array())
End Function

Public Function ArryCompare(ByVal a_v1 As Variant, _
                             ByVal a_v2 As Variant, _
                    Optional ByVal a_stop_after As Long = 0, _
                    Optional ByVal a_ignore_case As Boolean = True, _
                    Optional ByVal a_ignore_empty As Boolean = True) As Dictionary
' --------------------------------------------------------------------------
' Returns a Dictionary with n (a_stop_after) lines which are different
' between array 1 (a_v1) and array 2 (a_v2) with the line number as the
' key and the two different lines as item in the form: '<line>'vbLf'<line>'
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
    
    If Not mBasic.ArryIsAllocated(a_v1) And mBasic.ArryIsAllocated(a_v2) Then
        If a_ignore_empty Then mBasic.ArryTrimm a_v2
        For i = LBound(a_v2) To UBound(a_v2)
            dct.Add i + 1, "'" & a_v2(i) & "'" & vbLf
        Next i
    ElseIf mBasic.ArryIsAllocated(a_v1) And Not mBasic.ArryIsAllocated(a_v2) Then
        If a_ignore_empty Then mBasic.ArryTrimm a_v1
        For i = LBound(a_v1) To UBound(a_v1)
            dct.Add i + 1, "'" & a_v1(i) & "'" & vbLf
        Next i
    ElseIf Not mBasic.ArryIsAllocated(a_v1) And Not mBasic.ArryIsAllocated(a_v2) Then
        GoTo xt
    End If
    
    If a_ignore_empty Then mBasic.ArryTrimm a_v1
    If a_ignore_empty Then mBasic.ArryTrimm a_v2
    
    l = 0
    For i = LBound(a_v1) To Min(UBound(a_v1), UBound(a_v2))
        If StrComp(a_v1(i), a_v2(i), lMethod) <> 0 Then
            dct.Add i + 1, "'" & a_v1(i) & "'" & vbLf & "'" & a_v2(i) & "'"
            l = l + 1
            If a_stop_after <> 0 And l >= a_stop_after Then
                GoTo xt
            End If
        End If
    Next i
    
    If UBound(a_v1) < UBound(a_v2) Then
        For i = UBound(a_v1) + 1 To UBound(a_v2)
            dct.Add i + 1, "''" & vbLf & " '" & a_v2(i) & "'"
            l = l + 1
            If a_stop_after <> 0 And l >= a_stop_after Then
                GoTo xt
            End If
        Next i
        
    ElseIf UBound(a_v2) < UBound(a_v1) Then
        For i = UBound(a_v2) + 1 To UBound(a_v1)
            dct.Add i + 1, "'" & a_v1(i) & "'" & vbLf & "''"
            l = l + 1
            If a_stop_after <> 0 And l >= a_stop_after Then
                GoTo xt
            End If
        Next i
    End If

xt: Set ArryCompare = dct
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
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

Public Function ArryDims(ByVal a_arr As Variant) As Integer
' ----------------------------------------------------------------------------
' Returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This may as well be tested by means of ArryIsAllocated.
' ----------------------------------------------------------------------------
    On Error Resume Next
    Dim Ndx As Integer
    Dim Res As Integer
    
    '~~ Loop, increasing the dimension index Ndx, until an error occurs.
    '~~ An error will occur when Ndx exceeds the number of dimension
    '~~ in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(a_arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    ArryDims = Ndx - 1

End Function

Public Function ArryIsAllocated(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Retunrs TRUE when the array (a_arr) is allocated, i.e. has at least one item.
' ----------------------------------------------------------------------------
    
    On Error Resume Next
    ArryIsAllocated = UBound(a_arr) >= LBound(a_arr)
    On Error GoTo -1
    
End Function

Public Function ArryItems(ByVal a_arr As Variant) As Long
' ----------------------------------------------------------------------------
' Returns the number of items in a multi-dimensional array or a nested array.
' The latter is an array of which one or more items are again possibly multi-
' dimensional arrays. An unallocated array returns 0.
' ----------------------------------------------------------------------------
    
    Dim lDim    As Long
    Dim lItems  As Long
    Dim lDims   As Long
    Dim v       As Variant
    
    lDims = ArryDims(a_arr)
    
    Select Case lDims
        Case 0: lItems = 0
        Case 1
            lItems = (UBound(a_arr, 1) - LBound(a_arr, 1)) + 1
            For Each v In a_arr
                If IsArray(v) Or TypeName(v) Like "*()" Then
                    lItems = lItems + ArryItems(v)
                End If
            Next v
            
        Case Else
            lItems = 1
            For lDim = 1 To lDims
                lItems = lItems * ((UBound(a_arr, lDim) - LBound(a_arr, lDims)) + 1)
            Next
    End Select
    ArryItems = lItems
    
End Function

Private Sub ArryItemsTest()
        
    Dim arr1 As Variant
    Dim arr2 As Variant
    Dim arr3 As Variant
    
    Arry(arr1, 3) = "X"
    Arry(arr1, 5) = "Y"
    Arry(arr1, 9) = "Z"
    Arry(arr2) = arr1
    
    Debug.Assert ArryDims(arr1) = 1
    Debug.Assert ArryItems(arr1) = 10
    
    Debug.Assert ArryDims(arr2) = 1
    Debug.Assert ArryItems(arr2) = 11
    Debug.Assert ArryItems(arr3) = 0

End Sub

Public Sub ArryRemoveItems(ByRef a_va As Variant, _
                  Optional ByVal a_element As Variant, _
                  Optional ByVal a_index As Variant, _
                  Optional ByVal a_no_of_elements = 1)
' ------------------------------------------------------------------------------
' Returns a 'one dimensional'! array (a_va) with the number of elements
' (a_no_of_elements) removed whereby the start element may be indicated by the
' element number 1,2,... (a_element) or the index (a_index) which must be
' within the array's LBound to Ubound. Any inappropriate provision of arguments
' raises an error message. When the last item in an array is removed the
' returned array is erased (i.e is no longer allocated).
'
' W. Rauschenberger, Berlin Feb 2022
' ------------------------------------------------------------------------------
    Const PROC = "ArryRemoveItems"

    On Error GoTo eh
    Dim a           As Variant
    Dim iElement    As Long
    Dim iIndex      As Long
    Dim iNoOfItems  As Long
    Dim i           As Long
    Dim ia          As Long
    Dim iNewUBound  As Long
    
    If Not mBasic.ArryIsAllocated(a_va) Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Array not provided!"
    Else
        a = a_va
        iNoOfItems = UBound(a) - LBound(a) + 1
    End If
    If Not ArryDims(a) = 1 Then
        Err.Raise AppErr(2), ErrSrc(PROC), "Array must not be multidimensional!"
    End If
    If Not IsNumeric(a_element) And Not IsNumeric(a_index) Then
        Err.Raise AppErr(3), ErrSrc(PROC), "Neither FromElement nor FromIndex is a numeric value!"
    End If
    If IsNumeric(a_element) Then
        iElement = a_element
        If iElement < 1 _
        Or iElement > iNoOfItems Then
            Err.Raise AppErr(4), ErrSrc(PROC), "vFromElement is not between 1 and " & iNoOfItems & " !"
        Else
            iIndex = LBound(a) + iElement - 1
        End If
    End If
    If IsNumeric(a_index) Then
        iIndex = a_index
        If iIndex < LBound(a) Or iIndex > UBound(a) _
        Then Err.Raise AppErr(5), ErrSrc(PROC), "FromIndex is not between " & LBound(a) & " and " & UBound(a) & " !"
        For ia = LBound(a) To iIndex
            iElement = iElement + 1
        Next ia
    End If
    If iElement + a_no_of_elements - 1 > iNoOfItems _
    Then Err.Raise AppErr(6), ErrSrc(PROC), "FromElement (" & iElement & ") plus the number of elements to remove (" & a_no_of_elements & ") is beyond the number of elelemnts in the array (" & iNoOfItems & ")!"
    
    For i = iIndex + a_no_of_elements To UBound(a)
        a(i - a_no_of_elements) = a(i)
    Next i
    
    iNewUBound = UBound(a) - a_no_of_elements
    If iNewUBound < 0 Then Erase a Else ReDim Preserve a(LBound(a) To iNewUBound)
    a_va = a
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub ArryTrimm(ByRef a_arr As Variant)
' ------------------------------------------------------------------------------
' Returns the array (a_arr) with all leading and trailing blank items removed. Any
' vbCr, vbCrLf, vbLf are ignored. When the array contains only blank items the
' returned array is erased.
' ------------------------------------------------------------------------------
    Const PROC  As String = "ArryTrimm"

    On Error GoTo eh
    Dim i As Long
    
    '~~ Eliminate leading blank lines
    If Not mBasic.ArryIsAllocated(a_arr) Then Exit Sub
    
    Do While (Len(Trim$(a_arr(LBound(a_arr)))) = 0 Or Trim$(a_arr(LBound(a_arr))) = " ") And UBound(a_arr) >= 0
        mBasic.ArryRemoveItems a_va:=a_arr, a_index:=i
        If Not mBasic.ArryIsAllocated(a_arr) Then Exit Do
    Loop
    
    If mBasic.ArryIsAllocated(a_arr) Then
        Do While (Len(Trim$(a_arr(UBound(a_arr)))) = 0 Or Trim$(a_arr(LBound(a_arr))) = " ") And UBound(a_arr) >= 0
            If UBound(a_arr) = 0 Then
                Erase a_arr
            Else
                ReDim Preserve a_arr(UBound(a_arr) - 1)
            End If
            If Not mBasic.ArryIsAllocated(a_arr) Then Exit Do
        Loop
    End If

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Function BaseName(ByVal v As Variant) As String
' ----------------------------------------------------------------------------
' Returns the file name (v) without the extension. The argument may be a file
' name a full file name, a file object or a Workbook object.
' ----------------------------------------------------------------------------
    Const PROC  As String = "BaseName"
    
    On Error GoTo eh
    
    With New FileSystemObject
        Select Case TypeName(v)
            Case "String":           BaseName = .GetBaseName(v)
            Case "Workbook", "File": BaseName = .GetBaseName(v.Name)
            Case Else:               Err.Raise AppErr(1), ErrSrc(PROC), _
                                     "The parameter (v) is neither a string nor a File or Workbook object (TypeName = '" & TypeName(v) & "')!"
        End Select
    End With

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

Private Sub DictTest()

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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mBasic." & sProc
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
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add Key:=vKey, Item:=s_dct.Item(vKey)
    Next i
    
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
    Dim hWnd As LongPtr
    Dim RetVal

    hWnd = GetForegroundWindow
    
    lStyle = GetWindowLongPtr(hWnd, GWL_STYLE Or WS_THICKFRAME)
    RetVal = SetWindowLongPtr(hWnd, GWL_STYLE, lStyle)

End Sub

Public Function Max(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the maximum value of all values provided (va).
' --------------------------------------------------------
    
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
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

Public Function RangeAsArray(ByVal r_rng As Range) As Variant
    Const PROC = "RangeAsArray"
    
    Dim arr As Variant
    
    Select Case True
        Case r_rng.Cells.Count = 1:     arr = Array(r_rng.Value)                                            ' single cell
        Case r_rng.Columns.Count = 1:   arr = Application.Transpose(r_rng.Value)                            ' single column
        Case r_rng.Rows.Count = 1:      arr = Application.Transpose(Application.Transpose(r_rng.Value))    ' single row
        Case r_rng.Rows.Count = 2
        Case r_rng.Columns.Count = 2:   arr = r_rng.Value
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "Range cannot be transferred/transposed into an aray!"
    End Select
    RangeAsArray = arr
    
End Function

Public Sub README(Optional ByVal r_base_url As String = vbNullString, _
                  Optional ByVal r_bookmark As String = vbNullString)
' ----------------------------------------------------------------------------
' Displays the given url (r_base_url) with the given bookmark (r_bookmark) in the
' computer's default browser. When no url is provided it defaults to this
' component's README url in the public GitHub repo.
' ----------------------------------------------------------------------------
    If r_base_url = vbNullString _
    Then r_base_url = "https://github.com/warbe-maker/VBA-Basics"
    
    If r_bookmark = vbNullString Then
        mBasic.ShellRun r_base_url
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        mBasic.ShellRun r_base_url & r_bookmark
    End If
        
End Sub

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
               Optional ByRef s_item As Variant = -999999999, _
               Optional ByRef s_lvl As Long = 0) As Variant
' ----------------------------------------------------------------------------
' Returns TRUE when an item (s_item) is stacked at a given level (s_lvl) or
' when no level is provided, when it is stacked at any level. In the latter
' case the level (s_lvl) is returned.
' Returns the stacked item when none is provided and but a level is.
' Restriction: The function works with any kind of object an an item which is
'              not -999999999, which is regarded no item is provided.
' ----------------------------------------------------------------------------
    Const PROC = "StckEd"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim i       As Long
    
    If s_stck Is Nothing Then Set s_stck = New Collection
    
    Select Case True
        Case s_lvl <> 0 And s_lvl > s_stck.Count
            GoTo xt

        Case s_lvl = 0 And VarType(s_item) <> vbObject
            If s_item = -999999999 Then GoTo xt
            '~~ A specific item has been provided
            For i = 1 To s_stck.Count
                If s_stck(i) = s_item Then
                    s_lvl = i
                    StackEd = True
                    GoTo xt
                End If
            Next i
        
        Case s_lvl = 0 And VarType(s_item) = vbObject
            For i = 1 To s_stck.Count
                If s_stck(i) Is s_item Then
                    s_lvl = i
                    StackEd = True
                    GoTo xt
                End If
            Next i
                    
        Case VarType(s_item) <> vbObject
            If s_item = -999999999 Then
                StackEd = s_stck(s_lvl)
            ElseIf s_stck(s_lvl) = s_item Then
                StackEd = True
                GoTo xt
            End If
        
        Case VarType(s_item) = vbObject
            If s_stck(s_lvl) Is s_item Then
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
' Common Stack Pop service. Returns the last item pushed on the stack (stck)
' and removes the item from the stack. When the stack (stck) is empty a
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
' Common Stack Push service. Pushes (adds) an item (stck_item) to the stack
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
' Common Stack Top service. Returns the top item from the stack (stck), i.e.
' the item last pushed. If the stack is empty a vbNullString is returned.
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

