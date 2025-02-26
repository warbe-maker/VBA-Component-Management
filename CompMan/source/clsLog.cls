VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1 ' ensures the index conforms with the column number
' -----------------------------------------------------------------------------------
' Class Module clsLog: Common services for writing a log file including writing
' ==================== entries aligned in columns.
'
' Methods/services:
' -----------------
' - AlignmentHeaders Explicitly specifies the alignment for each column's
'                    header, defaults to "C"entered
' - ColsAlign   Explicitly specifies the alignment for each column's
'                    item, defaults to "L"eft adjusted when not provided.
' - ColsDelimiter    Defaults to | when Headers are specified and defaults to a
'                    single space when no Headers are specified.
' - Dsply            Displays the Log-Files content by means of ShellRun.
' - Entry            Either writes a single string, i.e. one without any any |
'                    and or subsequent string elements of writes an number of
'                    provided strings or one string with vertical bars, both
'                    indicating items to be writen alingned in columns.
' - Headers          Specifies the column headers
' - EntryItemsMaxLenght   Explicitly specifies each columns minimum width. The width of,
'                    defaults to the width of the corresponding column header
'                    string, is adjusted to the width of the corresponding header
'                    string when one is provided.
' - NewFile          Removes an existing file and creates a new one when entries
'                    are written
' - NewLog           Explicit indication that the next Entry is the first of a
'                    new series of log entries. I.e. with the next Entry a
'                    Delimiter line is written (provided its not a new log file),
'                    a title is written (provided one has been specified), and
'                    a header is written (provided one has been specified).
'                    An explicit call of this method is only required, in case
'                    neither a Title nor Headers had been specified but a
'                    delimiter line should indicate a new series of log entries.
' - Title            Specifies the - optionally multiline - title of a new series
'                    of log entries. Triggers the writing of the column headers
'                    provided specified.
' - WithTimeStamp    Prefixes each log entry with a timestamp in the format
'                    yy-mm-dd-hh:mm:ss when True
'
' Properties:
' -----------
' - FileBaseName  r/w The log-file's base name, defaults to the
'                     ActiveWorkbook's BaseName.
' - FileExtension r/w The Log-file's file extension
' - FileFullName  r/w w: Specifies the full name of the log-file, thereby providing
'                        the FileBaseName and the FileExtension.
'                     r: When no FileFullName ever has been provided, the default
'                        log-file's full name is assemled from FileLocation,
'                        FileBaseName and FileExtension.
'                     the ActiveWorkbook's BaseName with a .log file extention.
' - FileLocation  r/w Defaults to the ActiveWorkbook's Path
' - KeepLogs        w Specifies the number of logs, i.e. a set of log entries, is
'                     kept when the log file is reorganized.
'                     before it is deleted and re-created, defaults to 1.
' - LogFile       r   The log-file as file object
' - Path            w Sepcifies the log-file's location/folder, defaults to the
'                     ActiveWorkbook's parent folder
'
' W. Rauschenberger, Berlin Nov 2024
' See: https://github.com/warbe-maker/VBA-Log-Service
' -----------------------------------------------------------------------------------
Private aColsAlign              As Variant
Private aColsFill               As Variant
Private aColsHeadersLines       As Variant
Private aColsWidth              As Variant
Private aFills                  As Variant
Private aHeaders                As Variant
Private aItems                  As Variant
Private aTitleLines             As Variant
Private bInitialized            As Boolean
Private bNewLog                 As Boolean
Private bWithTimeStamp          As Boolean
Private fso                     As New FileSystemObject
Private lBase                   As Long
Private lCols                   As Long
Private lColumns                As Long
Private lKeepLogs               As Long
Private sColsDelimiter          As String
Private sColsMargin             As String
Private sFileExtension          As String
Private sFileBaseName           As String
Private sFileFullName           As String
Private sFileLocation           As String
Private sLastSpec               As String ' Entry, Title, Header, else
Private sTitleAlign             As String
Private sTitleFill              As String

Private Enum enAlign
    enAlignLeft = 1
    enAlignRight = 2
    enAlignCentered = 3
End Enum

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

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

Private Sub Class_Initialize()
' -----------------------------------------------------------------------------------
' Initializing defaults
' -----------------------------------------------------------------------------------
    NewLogInit
    lKeepLogs = 10 ' default
    bWithTimeStamp = False
    bNewLog = True
    lBase = ArrayBase
    ColsDelimiter = " "
    '~~ Log file name defaults
    sFileBaseName = fso.GetBaseName(ActiveWorkbook.FullName)
    sFileExtension = "log"
    sFileLocation = ActiveWorkbook.Path
    sFileFullName = FileFullName

End Sub

Private Sub Class_Terminate()
    '~~ In case the file already existed it may have passed the KeepLogs limit
    If fso.FileExists(sFileFullName) Then
        Reorg sFileFullName, lKeepLogs
    End If

    Set fso = Nothing
    ArraysErase
    sColsDelimiter = vbNullString
End Sub

Private Property Get AlignmentHeaderCol(Optional ByVal c_col As Long) As String
    AlignmentHeaderCol = Arry(aColsAlign, c_col)
End Property

Private Property Let AlignmentHeaderCol(Optional ByVal c_col As Long, _
                                                ByVal c_align As String)
    Arry(aColsAlign, c_col) = c_align
End Property

Private Property Get AlignmentItemCol(Optional ByVal c_col As Long) As String
    AlignmentItemCol = Arry(aColsAlign, c_col)
End Property

Private Property Let AlignmentItemCol(Optional ByVal c_col As Long, _
                                               ByVal c_align As String)
    Arry(aColsAlign, c_col) = c_align
End Property

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

Private Function ArryItems(ByVal a_arr As Variant, _
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

Private Property Get ArryItem(Optional ByRef a_arr As Variant, _
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

Private Property Let ArryItem(Optional ByRef a_arr As Variant, _
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

Private Function ArryIsAllocated(ByVal a_arr As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the array (a_arr) is allocated, i.e. has at least one Item.
' ----------------------------------------------------------------------------
    
    On Error Resume Next
    ArryIsAllocated = UBound(a_arr) >= LBound(a_arr)
    On Error GoTo 0
    Err.Clear
    
End Function

Private Function IsInteger(ByVal i_value As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is numeric, and has no decimals.
' ----------------------------------------------------------------------------
    If Not i_value = Empty Then
        If IsNumeric(i_value) _
        Then IsInteger = (i_value = Int(i_value))
    End If
End Function

Private Function ArryIndices(ParamArray a_indices() As Variant) As Collection
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
Private Property Get Coll(Optional ByRef c_coll As Collection, _
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
            If VBA.IsObject(c_coll(c_argmnt)) _
            Then Set Coll = c_coll(c_argmnt) _
            Else Coll = c_coll(c_argmnt)
        Case Else
            '~~ When an argument is provided which is not an integer, the index of
            '~~ the first item is returned which is identical with the provided argument
            With c_coll
                If VBA.IsObject(c_argmnt) Then
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

Private Property Let Coll(Optional ByRef c_coll As Collection, _
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
              And Not VBA.IsObject(c_coll(c_argmnt))
                '~~ Update!
                .Remove c_argmnt
                .Add c_var, c_argmnt
            Case c_argmnt <= c_coll.Count _
              And VBA.IsObject(c_coll(c_argmnt))
                '~~ Replace value in object
              
        End Select
    End With

End Property

Private Property Get ColAlign(Optional ByVal c_col As Long) As String
' ----------------------------------------------------------------------------
' Return a column's width, whereby an available explicit specification
' supercedes a width iplicitely derived from a string.
' ----------------------------------------------------------------------------
    ColAlign = Arry(aColsAlign, c_col)
End Property

Private Property Let ColAlign(Optional ByVal c_col As Long, _
                                       ByVal c_align As String)
    Arry(aColsAlign, c_col) = c_align
End Property

Private Property Get ColFill(Optional ByVal c_col As Long) As String
    ColFill = Arry(aFills, c_col)
End Property

Private Property Let ColFill(Optional ByVal c_col As Long, _
                                      ByVal c_fill As String)
    Arry(aFills, c_col) = c_fill
End Property

Private Property Get ColItem(Optional ByVal c_col As Long) As String
' -----------------------------------------------------------------------------------
' Returns the item of a column (c_col):
' - AlignedImplicitly as explicitly specified or, when not specified explicitly, aligned as
'   indicated implicit by the very first.
' - In the width explicitly specified and/or specified by the header's width or when
'   no width is specified by the width of the very first row's item width.
' Note 1: Any column(c_col) beyond UBound(aItems) is ignored, i.e. returns a
'         vbNullString.
' -----------------------------------------------------------------------------------
    Const PROC = "ColItem-Get"
    
    Dim lWidth  As Long
    
    On Error GoTo xt
    If UBound(aItems) >= LBound(aItems) Then
        On Error GoTo eh
        If bNewLog Then
            '~~ When this is the first line of a new series of log entries the width
            '~~ may still be expanded when the item's width is greater than the
            '~~ already specified - in case
            lWidth = ColWidth(c_col)
            ColWidth(c_col) = Max(Len(aItems(c_col)), lWidth)
        End If
        
        If c_col < lColumns Then
            ColItem = Align(aItems(c_col), AlignmentItemCol(c_col), ColWidthFinal(c_col), ColFill(c_col), sColsMargin, True)
        ElseIf c_col = lColumns Then
            '~~ The width of the rightmost column is unlimited and left aligned.
            ColItem = Align(aItems(c_col), "L", Len(aItems(c_col)) + Len(sColsMargin), " ", sColsMargin, True)
        End If
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let ColItem(Optional ByVal c_col As Long, _
                                      ByVal c_Item As String)
    Arry(aItems, c_col) = c_Item
End Property

Private Property Get Cols() As Long
    Cols = Max(lCols, ArrayNoOfElements(aColsAlign), ArrayNoOfElements(aColsWidth))
End Property

Private Property Let Cols(ByVal c_cols As Long):    lCols = c_cols: End Property

Friend Property Get ColsDelimiter() As String:  ColsDelimiter = sColsDelimiter: End Property

Friend Property Let ColsDelimiter(ByVal s As String)
    Select Case s
        Case " ": sColsDelimiter = s: sColsMargin = vbNullString
        Case "|": sColsDelimiter = s: sColsMargin = " "
    End Select
End Property

Public Property Let ColsHeader(Optional ByVal c_line As Long, _
                                        ByVal c_cols_header As Variant)
' -----------------------------------------------------------------------------------
' Saves/adds provided column headers (c_cols_header) to an array (aColsHeaderLines)
' where each item is one column header line. A lines column headers (c_cols_header)
' may be provided as array, collection, or as a string delimited any string
' identified by by the function SplitStr.
'
' Precondition:
' The columns delimiter (ColsDelimiter) and the columns size (ColsSpecs) had been
' specified, whereby only the widths matter since headers alignment defaults to
' centered filled with spaces.
' -----------------------------------------------------------------------------------
    Dim v           As Variant
    Dim s           As String
    Dim iCol        As Long
    Dim sColHeader  As String
    Dim sColFill    As String
    Dim sItem       As String
    Dim aHeaders    As Variant
    Dim sDelim      As String
    
    c_line = c_line ' avoid dead code
    aHeaders = StringAsArray(c_cols_header)
    sColFill = " "

    For Each v In aHeaders
        iCol = iCol + 1
        sColHeader = Trim$(v) ' ensure centered regardless any leading or trailing spaces
        ColWidth(iCol) = Len(v)
        sItem = Align(sColHeader, "C", ColWidth(iCol), sColFill, ColsMargin, True)
        s = ConcatenateStrings(s, sItem, ColsDelimiter, sDelim)
    Next v
    Arry(aColsHeadersLines) = s ' add s to the array
    
End Property

Public Property Get ColsMargin() As String
    If sColsDelimiter <> " " Then ColsMargin = " " Else ColsMargin = vbNullString
End Property

Public Property Let ColsSpecs(ByVal c_specs As String)
' ----------------------------------------------------------------------------
' Specifying columns specifications length, alignment, and fill
' ----------------------------------------------------------------------------
    Dim aSpecs  As Variant
    Dim iCol    As Long
    Dim v       As Variant
    
    aSpecs = StringAsArray(c_specs)
    ArrayErase aColsWidth
    ArrayErase aColsAlign
    ArrayErase aColsFill
    
    For Each v In aSpecs
        iCol = iCol + 1
        ColSpec iCol, v
    Next v
    
End Property

Public Property Get ColWidth(Optional ByVal c_col As Long) As Long
' ----------------------------------------------------------------------------
' The (maximized) column width is used for:
' 1. The Title line (when one had been specified)
' 2. The column header lines and the header underline
' 3. The final re-assembling of the very first column arranged Entry line
'    (which triggered the above two)
' ----------------------------------------------------------------------------
    ColWidth = Arry(aColsWidth, c_col)

End Property

Public Property Let ColWidth(Optional ByVal c_col As Long, _
                                      ByVal c_width As Long)
' ----------------------------------------------------------------------------
' Initially a column's width may or may not be specified explicitly by
' 'ColsWidth'. Provided the truncate option is False the width is subsequently
' maximized by the aligned item which may include a possible (implicitly
' specified) filling. The final column width is the one specified explicitly
' or derived implicitly plus the length of the specified filling. I.e.
' an item specified "xxxx .:" will finally end up in a column width of 10.
' ----------------------------------------------------------------------------
    Arry(aColsWidth, c_col) = Max(Arry(aColsWidth, c_col), c_width)
    
End Property

Friend Property Get FileBaseName() As String:           FileBaseName = sFileBaseName:   End Property

Friend Property Let FileBaseName(ByVal s As String):    sFileBaseName = s:              End Property

Friend Property Get FileExtension() As String:          FileExtension = sFileExtension: End Property

Friend Property Let FileExtension(ByVal s As String):   sFileExtension = s:             End Property

Friend Property Get FileFullName() As String
    FileFullName = sFileLocation & "\" & sFileBaseName & "." & sFileExtension
    FileFullName = Replace(FileFullName, "\\", "\")
    FileFullName = Replace(FileFullName, "..", ".")
End Property

Friend Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Specifies the log file's name and location, by the way maintaining the Path
' and the FileBaseName property.
' ----------------------------------------------------------------------------
    
    With fso
        If s <> FileFullName Then
            '~~ The log file's name is not/no longer the one previously used
            If .FileExists(FileFullName) Then .DeleteFile FileFullName
        End If
        sFileExtension = .GetExtensionName(s)
        sFileLocation = .GetParentFolderName(s)
        sFileBaseName = .GetBaseName(s)
    End With
    sFileFullName = s
    
End Property

Friend Property Get FileLocation() As String:           FileLocation = sFileLocation:   End Property

Friend Property Let FileLocation(ByVal s As String):    sFileLocation = s:              End Property

Friend Property Let KeepLogs(ByVal l As Long):          lKeepLogs = l:                  End Property

Friend Property Let Path(ByVal v As Variant)
' -----------------------------------------------------------------------------------
' Specifies the location (folder) for the log file based on the provided information
' which may be a string, a Workbook, or a folder object.
' -----------------------------------------------------------------------------------
    Const PROC = "Path-Let"
    Dim wbk As Workbook
    Dim fld As Folder
    
    Select Case VarType(v)
        Case VarType(v) = vbString
            sFileLocation = v
        Case VarType(v) = vbObject
            If TypeOf v Is Workbook Then
                Set wbk = v
                sFileLocation = wbk.Path
            ElseIf TypeOf v Is Folder Then
                Set fld = v
                sFileLocation = fld.Path
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a string specifying a " & _
                                                   "folder's path, nor a Workbook object, nor a Folder object!"
            End If
    End Select
    
End Property

Private Property Get SplitStr(ByRef s As String)
' ------------------------------------------------------------------------------
' Returns the split string in string (s) used by VBA.Split() to turn the string
' into an array.
' ------------------------------------------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
End Property

Public Function Align(ByVal a_string As String, _
             Optional ByVal a_align As String = "L", _
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
        Case "L":   Align = AlignLeft(a_string, lLength, a_fill, a_margin, a_col_arranged)
        Case "R":   Align = AlignRght(a_string, lLength, a_fill, a_margin, a_col_arranged)
        Case "C":   Align = AlignCntr(a_string, lLength, a_fill, a_margin, a_col_arranged)
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AlignCntr(ByVal a_string As String, _
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
        If Len(s) > lWidth Then s = Left(s, lWidth)
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

Private Function AlignLeft(ByRef a_string As String, _
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
        If Len(s) > lWidth Then s = Left(s, lWidth)
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
    If VBA.Trim$(a_string) = vbNullString Then
        AlignLeft = String(Len(s), " ")
    Else
        AlignLeft = s
    End If
    If a_col_arranged Then AlignLeft = a_margin & AlignLeft & a_margin _
    
End Function

Private Function AlignRght(ByVal a_string As String, _
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
    End If

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

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error number never conflicts
' with VB runtime error. Thr function returns a given positive number
' (app_err_no) with the vbObjectError added - which turns it to negative. When
' the provided number is negative it returns the original positive "application"
' error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ArrayAsFile(ByVal a_arry As Variant, _
                    Optional ByRef a_file As Variant = vbNullString, _
                    Optional ByVal a_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes all items of an array (a_arry) to a file (a_file) which might be a
' file object, a file's full name. When no file (a_file) is provided a
' temporary file is returned, else the provided file (a_file) as object.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    On Error GoTo xt
    If UBound(a_arry) >= LBound(a_arry) Then
        Select Case True
            Case a_file = vbNullString:     a_file = TempFile
            Case TypeName(a_file) = "File": a_file = a_file.Path
        End Select
        
        TxtFile(a_file, a_file_append) = Join(a_arry, vbCrLf)
        Set ArrayAsFile = fso.GetFile(a_file)
    End If
xt:
End Function

Private Function ArrayBase() As Long
    Dim arr As Variant
    arr = Array("a")
    ArrayBase = LBound(arr)
End Function

Private Function ArrayErase(ByRef a_arr As Variant)
    If IsArray(a_arr) Then Erase a_arr
End Function

Private Function ArrayHasElements(ByVal a_arr As Variant, _
                         Optional ByRef a_elements As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the array (a_arr) has at least one element.
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    On Error Resume Next
    v = a_arr(LBound(a_arr))
    ArrayHasElements = Err.Number = 0
    Err.Clear
    If ArrayHasElements Then
        If LBound(a_arr) = 0 _
        Then a_elements = UBound(a_arr) - LBound(a_arr) _
        Else a_elements = UBound(a_arr) - LBound(a_arr) + 1
    End If

End Function

Private Function ArrayIsAllocated(a_arr As Variant) As Boolean
' -----------------------------------------------------------------------------------
' Retruns TRUE when an array (arr) is allocated, i.e. when it has a valid LBound and
' and an LBound which is less or equal its UBound.
' -----------------------------------------------------------------------------------
    On Error Resume Next
    ArrayIsAllocated = UBound(a_arr, 1) >= LBound(a_arr, 1)
    On Error GoTo 0
    Err.Clear
End Function

Private Function ArrayNoOfElements(ByVal a_arr As Variant) As Long
    Dim l As Long
    If ArrayHasElements(a_arr, l) Then
        ArrayNoOfElements = l
    End If
    
End Function

Private Sub ArraysErase()
    
    ArrayErase aColsAlign
    ArrayErase aColsWidth
    ArrayErase aColsHeadersLines
    ArrayErase aItems
    ArrayErase aTitleLines
    
End Sub

Private Function CollectionAsArray(ByVal v_items As Collection) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    Dim v   As Variant
    
    For Each v In v_items
        If IsObject(v, s) _
        Then Arry(arr) = s _
        Else Arry(arr) = v
    Next v
    CollectionAsArray = arr
    
End Function

Private Function ColsHeadersLines(ByRef c_width As Long) As String
' ------------------------------------------------------------------------------
' Returns all header lines (aColsHeadersLines) as one string delimited by vbCrLf.
' When none had been specified the function returns a vbNullString.
' ------------------------------------------------------------------------------
    Const PROC = "LogColsHeaderWrite"
    
    On Error GoTo eh
    Dim sColDelimUl As String
    Dim sHeaderUl   As String
    Dim i           As Long
    Dim sLines      As String
    Dim v           As Variant
    Dim l           As Long
    
    If ArrayHasElements(aColsHeadersLines) Then
        '~~ Prepare cols header top/bottom line
        For i = LBound(aColsWidth) To UBound(aColsWidth)
            l = ColWidthFinal(i) + (Len(ColsMargin) * 2)
            sHeaderUl = ConcatenateStrings(sHeaderUl, String(l, "-"), "+", sColDelimUl)
        Next i
        '~~ Line above
        sLines = sHeaderUl
        c_width = Len(sLines)
        '~~ Header lines
        For Each v In aColsHeadersLines
            sLines = sLines & vbCrLf & CStr(v)
        Next v
        '~~ Line below
        sLines = sLines & vbCrLf & sHeaderUl
        ArrayErase aColsHeadersLines
    
        ColsHeadersLines = sLines
    End If
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ColsHeadersSpecified() As Boolean
    ColsHeadersSpecified = ArrayHasElements(aColsHeadersLines)
End Function

Public Function ColsItems(ParamArray c_items() As Variant) As String
' -----------------------------------------------------------------------------------
' Returns a string with the provided items (c_items) arranged in columns, whereby:
' - column widths provided beforehand by means of ColsWidth are considered,
'   expanded to the actual required width when the truncate option (c_truncate) is
'   False and truncated to the specified width otherwise. For a column no width has
'   been specified beforehand the column width of the first aligned column item is
'   considered even when the truncate option (c_truncate) is True.
' - a column delimiter may be a single space or a | (vertical bar). In the first
'   case the margin for each individual item becomes a vbNullString, in the second
'   case the margin will default to a single space.
'
' Precondition: Column widths had been provided.
' -----------------------------------------------------------------------------------
    Const PROC = "ColsItems"
    
    On Error GoTo eh
    Dim iCol       As Long
    Dim s          As String
    Dim sColDelim  As String
    Dim sColItem   As String
    Dim sColString As String
    Dim sItemAlgnd As String
    Dim v          As Variant
    Dim aItems     As Variant
    
    aItems = c_items
    If Not ArrayHasElements(aItems) Then GoTo xt
    
    For Each v In aItems
        iCol = iCol + 1
        sColItem = v
        If sColItem = vbNullString Then sColItem = " "
        
        Select Case ColAlign(iCol)
            Case "L":   sColString = RTrim$(sColItem)
            Case "C":   sColString = Trim$(sColItem)
            Case "R":   sColString = LTrim$(sColItem)
        End Select
        
        '~~ With the exception of the rightmost columns - which has an unlimited length -
        '~~ each column requires a specified width.
        If ColWidth(iCol) = 0 And iCol < Cols _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "Missing width spec for column " & iCol & " !"
        '~~ The rightmost column is adjusted to the max
        If iCol = Cols _
        Then ColWidth(iCol) = Len(sColString) ' rightmost column width is maximized on the go
        
        sItemAlgnd = Align(sColItem, ColAlign(iCol), ColWidth(iCol), ColFill(iCol), ColsMargin, True)
        s = ConcatenateStrings(s, sItemAlgnd, ColsDelimiter, sColDelim)
    Next v
    
xt: NewLogLines Len(s) ' when nothing else is specified this is the minimum length for title and headers
    WriteLines s
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub ColSpec(ByVal c_col As Long, _
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
        Case ColSpecTry(s, "Left"):     sAlign = "L"
        Case ColSpecTry(s, "Center"):   sAlign = "C"
        Case ColSpecTry(s, "Right"):    sAlign = "R"
        Case ColSpecTry(s, "Lft"):      sAlign = "L"
        Case ColSpecTry(s, "Cntr"):     sAlign = "C"
        Case ColSpecTry(s, "Rght"):     sAlign = "R"
        Case ColSpecTry(s, "L"):        sAlign = "L"
        Case ColSpecTry(s, "C"):        sAlign = "C"
        Case ColSpecTry(s, "R"):        sAlign = "R"
        Case Else:                      sAlign = "L" ' the default
    End Select
    ColAlign(c_col) = sAlign
    
    '~~ Get fill (defaults to a single space when none is specified
    Select Case True
        Case ColSpecTry(s, " .:", sFill)
        Case ColSpecTry(s, ".:", sFill)
        Case ColSpecTry(s, " -", sFill):    ColSpecTry s, "- ", sFill   ' if alignment is Centered
        Case ColSpecTry(s, " =", sFill):    ColSpecTry s, "= ", sFill   ' if alignment is Centered
        Case ColSpecTry(s, ".", sFill)
        Case ColSpecTry(s, "-", sFill):     ColSpecTry s, "-", sFill    ' if alignment is Centered
        Case ColSpecTry(s, " ", sFill)
        Case Else:                          sFill = " "  ' the default
    End Select
    ColFill(c_col) = sFill
    
    '~~ Get column width
    '~~ When alignment and fill specs had been extracted (and eliminated), the
    '~~ remaining (or the only spec) is to be the column width.
    If Not IsNumeric(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "The column specification does not provide a numeric value indicating the columns width." & "||" & _
         "The column width is mandatoryIt may be the only specification in case the fill is a single " & _
         "space (the default) and the alignment is Left (the default)."
    
    ColWidth(c_col) = CLng(s)
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ColSpecTry(ByRef c_strng As String, _
                            ByVal c_try As String, _
                   Optional ByRef c_return As String) As Boolean
                   
    If InStr(c_strng, c_try) <> 0 Then
        c_return = c_try
        ColSpecTry = True
        c_strng = Replace(c_strng, c_try, vbNullString)
    End If

End Function

Private Function ColWidthFinal(ByVal c_col As Long) As Long
' -----------------------------------------------------------------------------------
' Returns the final width of a column by considering the specified width and the fill
' string in case one is specified. This is added, because at least one fill - if
' centered two) is at least added to the string when the fill is not a single space.
' Note: This is the "inner width" providing enough space for the specified width
'       which considers the maximum length of the string only. I.e. a singel space
'       margin may still be added when the columns delimiter is a | (vettical bar).
' -----------------------------------------------------------------------------------
    
    Dim lMult As Long
    If ColAlign(c_col) = "C" Then lMult = 2 Else lMult = 1
    
    If ColFill(c_col) <> " " Then
        ColWidthFinal = ColWidth(c_col) + Len(ColFill(c_col)) * lMult
    Else
        ColWidthFinal = ColWidth(c_col)
    End If
    
End Function

Private Function ConcatenateStrings(ByVal c_string As String, _
                                    ByVal c_Item As String, _
                                    ByVal c_conc As String, _
                                    ByRef c_delim As String) As String
    ConcatenateStrings = c_string & c_delim & c_Item
    c_delim = c_conc
End Function

Private Function DictionaryAsArray(ByVal v_items As Dictionary) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    Dim v   As Variant
    
    For Each v In v_items
        If IsObject(v, s) _
        Then Arry(arr) = s _
        Else Arry(arr) = v
    Next v
    DictionaryAsArray = arr
    
End Function

Public Sub Dsply(Optional ByVal d_file_full_name = vbNullString)
' -----------------------------------------------------------------------------------
' Displays the Log-Files content, whereby the log-file defaults to sFileFullName.
' -----------------------------------------------------------------------------------
    If d_file_full_name = vbNullString Then d_file_full_name = sFileFullName
    ShellRun d_file_full_name
End Sub

Public Sub Entry(ByVal e_strng As String)
' ----------------------------------------------------------------------------
' Writes a single string appended as new log entry to the log file.
' ----------------------------------------------------------------------------
    Const PROC = "Entry"
    
    On Error GoTo eh
    
    bInitialized = False
    NewLogLines Len(e_strng) ' when nothing else is specified this is the minimum length for title and headers
    WriteLines e_strng
    sLastSpec = "Entry"
        
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    ErrSrc = "clsLog" & "." & sProc
End Function

Public Function FileAsArray(ByVal f_file As Variant, _
                   Optional ByVal f_empty_excluded = False) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim s       As String
    Dim sSplit  As String
    Dim v       As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    s = FileAsString(f_file, sSplit)
    If f_empty_excluded Then
        For Each v In Split(s, sSplit)
            If Trim$(v) <> vbNullString _
            Then Arry(arr) = v
        Next v
        FileAsArray = arr
    Else
        FileAsArray = Split(s, sSplit)
    End If
    
End Function

Public Function FileAsString(Optional ByVal f_file_full_name As String, _
                             Optional ByRef f_split As String = vbCrLf, _
                             Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    On Error Resume Next
    FileAsString = TxtFile(f_file_full_name)
    
    Select Case True
        Case InStr(FileAsString, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(FileAsString, vbCr) <> 0:   f_split = vbCr
        Case InStr(FileAsString, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate a trailing eof if any
    If Right(FileAsString, 1) = VBA.Chr(26) Then
        FileAsString = Left(FileAsString, Len(FileAsString) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    Do While Right(FileAsString, Len(f_split)) = f_split
        FileAsString = Left(FileAsString, Len(FileAsString) - Len(f_split))
        If Len(FileAsString) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        FileAsString = FileAsStringEmptyExcluded(FileAsString)
    End If
    
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
    Dim s   As String
    Dim i   As Long
    Dim v   As Variant
    
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

Private Function IsDelimiterLine(ByVal i_line As String) As Boolean
' -----------------------------------------------------------------------------------
' Returns TRUE when a line (i_line) is a delimiter line.
' -----------------------------------------------------------------------------------
    IsDelimiterLine = i_line Like "*===============*"
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
         "is the target format. However, the current item is an object which does " & _
         "not have a Name property!"
    
End Function

Private Function Max(ParamArray va() As Variant) As Variant
' -----------------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' -----------------------------------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' -----------------------------------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' -----------------------------------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Public Sub NewFile()
' -----------------------------------------------------------------------------------
' Explicitely starts a series of log entries with a new file (mainly for test purpose
' only)
' -----------------------------------------------------------------------------------
    With fso
        If .FileExists(sFileFullName) Then .DeleteFile sFileFullName
'        .CreateTextFile sFileFullName
    End With
    NewLog
End Sub

Public Sub NewLog()
' -----------------------------------------------------------------------------------
' Explicitly (by default) starts a new series of log entries even when neither Title
' nor ColsHeader are specified.
' Attention: Even when called with FALSE to indicate that no delimiter line is
'            desired, this still triggers the initialization of a new series of log
'            entried by means of NewlogInit
' -----------------------------------------------------------------------------------
    bNewLog = True
    If Not bInitialized Then NewLogInit
    sColsDelimiter = vbNullString
    ArrayErase aColsAlign
    ArrayErase aHeaders
    ArrayErase aColsHeadersLines

End Sub

Private Sub NewLogInit()
' -----------------------------------------------------------------------------------
' Essential for a new series of log entries. Triggered either by NewLog, Title, or
' ColsHeader. bInitialized is reset to false with the next subsequent Entry method call.
'
' Attention: This is the initialization of a new series of log entries - even when
'            NewLog False had been called in order to suppress the delimiter line.
' -----------------------------------------------------------------------------------
    If Not bInitialized Then
        ArrayErase aColsAlign
        ArrayErase aItems
        sColsDelimiter = vbNullString   ' Default in case no ColsHeader are specified
        sColsMargin = vbNullString      ' Default in case no ColsHeader are specified
        bInitialized = True
    End If
End Sub

Private Sub NewLogLines(ByVal n_width As Long)
' ----------------------------------------------------------------------------
' Writes any extra lines for a new log prior a log entry line (n-line) by
' considering the entry line's length.
' ----------------------------------------------------------------------------
    Const PROC = "NewLogLines"
    
    On Error GoTo eh
    Dim sHeaderLines    As String
    Dim sTitleLines     As String
    Dim lWidth          As Long
    
    If bNewLog Then
    
        '~~ Assemble header lines (width considered by title)
        If ColsHeadersSpecified Then sHeaderLines = ColsHeadersLines(n_width)
        '~~ Assemble title lines
        lWidth = n_width
        If TitlesSpecified Then sTitleLines = TitlesAssembled(lWidth)
        
        If fso.FileExists(sFileFullName) Then WriteLines String(lWidth, "=")  ' log delimiter line
        If sTitleLines <> vbNullString Then
            WriteLines sTitleLines
            If sHeaderLines = vbNullString _
            Then WriteLines String(lWidth, "-")
        End If
        If sHeaderLines <> vbNullString _
        Then WriteLines sHeaderLines
        bNewLog = False ' may be set to True again when a new Title is provided or explicit
    End If
    ArrayErase aTitleLines
    ArrayErase aColsHeadersLines
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Reorg(ByVal r_file_full_name As String, _
        Optional ByVal r_keep As Long = 1)
' ----------------------------------------------------------------------------
' Returns a log file (r_file_full_name) with any logs exceeding a number (r_logs)
' unstripped.
' Note: When only one log is kept (the default) the log file still may have
'       two logs. The one kept and the one added after the Reorg. When the
'       Reorg is done with the termination of the class instance 1 will be 1.
' ----------------------------------------------------------------------------
    Const PROC = "Reorg"
    
    On Error GoTo eh
    Dim aFile   As Variant
    Dim aLog    As Variant
    Dim cllLog  As New Collection
    Dim cllLogs As New Collection
    Dim i       As Long
    Dim lLen    As Long
    Dim v       As Variant
    Dim bTop    As Boolean: bTop = True
    
    If Not fso.FileExists(r_file_full_name) Then GoTo xt
    
    '~~ Collect the logs (i.e. lines delimited by a log delimiter line)
    '~~ into a Collection with each series of log entries/lines as a Collection
    '~~ with the array and the max line length as items
    aFile = FileAsArray(r_file_full_name)
    ArrayErase aLog
    With cllLogs
        For i = LBound(aFile) To UBound(aFile)
            If IsDelimiterLine(aFile(i)) And Not i = LBound(aFile) Then
                '~~ A delimiter line indicates the start of a new series of log entries
                cllLog.Add aLog
                cllLog.Add lLen
                .Add cllLog
                ArrayErase aLog
                Set cllLog = Nothing
                Set cllLog = New Collection
                lLen = 0
            End If
            If Not (i = LBound(aFile) And IsDelimiterLine(aFile(i))) Then
                '~~ A file beginning with a delimiter line the first line is ignored
                lLen = Max(lLen, Len(aFile(i)))
                Arry(aLog) = aFile(i)
            End If
        Next i
        '~~ Collect the last series of log entries
        cllLog.Add aLog
        cllLog.Add lLen
        .Add cllLog
        ArrayErase aLog
        Set cllLog = Nothing
        Set cllLog = New Collection
        
        '~~ Remove all excessive logs
        If .Count < r_keep Then GoTo xt
        Do While .Count > r_keep
            .Remove 1
        Loop
    End With
    
    '~~ Rewrite the remaining logs to the log file
    ArrayErase aFile
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

Private Sub ShellRun(ByVal sr_string As String, _
            Optional ByVal sr_show_how As Long = 1)
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
    Const PROC = "ShellRun"
    Const ERROR_SUCCESS = 32&
    Const ERROR_NO_ASSOC = 31&
    Const ERROR_OUT_OF_MEM = 0&
    Const ERROR_FILE_NOT_FOUND = 2&
    Const ERROR_PATH_NOT_FOUND = 3&
    Const ERROR_BAD_FORMAT = 11&
    
    On Error GoTo eh
    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       Err.Raise lRet, ErrSrc(PROC), "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   Err.Raise lRet, ErrSrc(PROC), "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   Err.Raise lRet, ErrSrc(PROC), "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       Err.Raise lRet, ErrSrc(PROC), "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, 1)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

Public Function StringAsArray(ByVal s_strng As String, _
                     Optional ByVal s_split As String = vbNullString, _
                     Optional ByVal s_trim As Variant = True) As Variant
' ----------------------------------------------------------------------------
' Returns a string (s_strng) split into an array of strings. When no split
' indicator (s_split) is provided it one is found by examination of the
' string (s_strng). When the option (s_trim) is TRUE (the default), "R", or
' "L" the items in the array are returned trimmed accordingly.
' Example 1: arr = StringAsArray("this is a string", " ") is returned as an
'            array with 3 items: "this", "is", "a", "string".
' Example 2: arr = StringAsArray(FileAsString(FileBaseName),sSplit,False) is
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
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    TxtFile(s_file, s_file_append) = s_strng
    Set StringAsFile = fso.GetFile(s_file)
    
End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(fso.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = fso.GetSpecialFolder(2)
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp

End Function

Private Function TimeStamp() As String ' Public for test purpose only
    TimeStamp = " "
    If bWithTimeStamp Then TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & TimeStamp
End Function

Public Sub Title(ParamArray t_lines() As Variant)
' ----------------------------------------------------------------------------
' Collects specified title lines for a following new series of log entries
' whereby multiple lines may be specified as strings (ParamArray items).
' Note: all titles may be specified with one call or one call per title line.
' ----------------------------------------------------------------------------
    Const PROC = "Title"
    
    Dim a       As Variant
    Dim i       As Variant
    
    If sLastSpec = "Entry" Or sLastSpec = "ColsHeader" Then
        ArrayErase aColsAlign
    End If
    
    a = t_lines
    If Not ArrayHasElements(a) Then GoTo xt
    
    If TitleSpecs(a(LBound(a)), sTitleAlign, sTitleFill) Then
        i = LBound(a) + 1
    Else
        i = LBound(a)
        If i < UBound(a) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided title apparently contains nothing more than a titles specs alignment and fill!"
    End If
    
    For i = i To UBound(a)
        Arry(aTitleLines) = a(i) ' collect the specified title lines
    Next i
    sLastSpec = "Title"
    NewLog
    
xt: Exit Sub

End Sub

Private Function TitlesAssembled(Optional ByRef t_width As Long = 0) As String
' ----------------------------------------------------------------------------
' Writes the collected title lines - if any - to the log file by considering
' the width of the first log entry and clears the collection. When no width
' (t_width) is provided the lines are adjusted to the longest line.
' ----------------------------------------------------------------------------
    Dim s           As String
    Dim sSplit      As String
    Dim v           As Variant
    
    On Error GoTo xt
    If Not UBound(aTitleLines) >= LBound(aTitleLines) Then
        For Each v In aTitleLines
            t_width = Max(t_width, Len(v))
        Next v
        
        For Each v In aTitleLines
            s = s & sSplit & Align(CStr(v), sTitleAlign, t_width, sTitleFill)
            sSplit = vbCrLf
        Next v
        TitlesAssembled = s
        ArrayErase aTitleLines
    End If
    
xt: Exit Function

End Function

Private Function TitleSpecs(ByVal t_specs As String, _
                            ByRef t_align As String, _
                            ByRef t_fill As String) As Boolean
' ----------------------------------------------------------------------------
' Provided the string (t_specs) has <= 3 characters the specified alignement
' (t_align) and fill (t_fill) and TRUE is returned, else FALSE with the
' default alignment "C" (dentered) and the fill " -".
' ----------------------------------------------------------------------------
    Dim s As String
    
    If Len(t_specs) > 5 Then
        '~~ This is not a title's specification of alignment and fill
        '~~ thus the defaults are returned
        t_align = "C"
        t_fill = " -"
        GoTo xt
    End If
    
    s = t_specs
    '~~ Get alignment (defaults to Left when none is specified
    Select Case True
        Case ColSpecTry(s, "Left"):     t_align = "L"
        Case ColSpecTry(s, "Center"):   t_align = "C"
        Case ColSpecTry(s, "Right"):    t_align = "R"
        Case ColSpecTry(s, "Lft"):      t_align = "L"
        Case ColSpecTry(s, "Cntr"):     t_align = "C"
        Case ColSpecTry(s, "Rght"):     t_align = "R"
        Case ColSpecTry(s, "L"):        t_align = "L"
        Case ColSpecTry(s, "C"):        t_align = "C"
        Case ColSpecTry(s, "R"):        t_align = "R"
        Case Else:                      t_align = "L" ' the default
    End Select
    
    '~~ Get fill (defaults to a single space when none is specified
    Select Case True
        Case ColSpecTry(s, " -", t_fill):   ColSpecTry s, "- ", t_fill   ' if alignment is Centered
        Case ColSpecTry(s, " =", t_fill):   ColSpecTry s, "= ", t_fill   ' if alignment is Centered
        Case ColSpecTry(s, "-", t_fill):    ColSpecTry s, "-", t_fill    ' if alignment is Centered
        Case ColSpecTry(s, "-", t_fill):    ColSpecTry s, "-", t_fill    ' if alignment is Centered
        Case ColSpecTry(s, " ", t_fill)
        Case Else:                          t_fill = " "  ' the default
    End Select
    TitleSpecs = True
    
xt: Exit Function

End Function

Private Function TitlesSpecified() As Boolean
    On Error Resume Next
    TitlesSpecified = UBound(aTitleLines) >= LBound(aTitleLines)
    On Error GoTo 0
    Err.Clear
End Function

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
        If ColWidth(t_col) = 0 Then ColWidth(t_col) = Len(t_string)
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

'Public Property Get ColsHeaders() As Variant
'' ----------------------------------------------------------------------------
'' Returns all columns arranged header lines, whereby each line is an array
'' of column headers.
'' ----------------------------------------------------------------------------
'    ColsHeaders = aColHeaders
'End Property

Private Function VarAsArray(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsArray"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsArray = CollectionAsArray(v_items)
        Case TypeName(v_items) = "File":        VarAsArray = FileAsArray(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsArray = DictionaryAsArray(v_items)
        Case VarType(v_items) = vbArray:        VarAsArray = v_items
        Case IsArray(v_items):                  VarAsArray = v_items
        Case VarType(v_items) = vbString:       VarAsArray = StringAsArray(v_items)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Public Sub WithTimeStamp(Optional ByVal b As Boolean = True):   bWithTimeStamp = b: End Sub

Private Sub WriteItems()
' -----------------------------------------------------------------------------------
' Add a line to the log file aligned in columns.
' - Any string exceeding the number of provided column headers and column widths is
'   ignored!
' - When no column headers had been provided (method ColsHeader) an error is raised.
' -----------------------------------------------------------------------------------
    Const PROC = "WriteItems"
    
    On Error GoTo eh
    Dim i               As Long
    Dim sItemsLine      As String
    
    '~~ When the ColsDelimiter is yet not specified
    '~~ delimiter defaults to a single space.
    If sColsDelimiter = vbNullString Then
        sColsMargin = ""
        sColsDelimiter = " "
    End If

    '~~ Assemble the entry/items line
    For i = LBound(aItems) To UBound(aItems)
        sItemsLine = sItemsLine & sColsDelimiter & ColItem(i)
    Next i
    
    If bNewLog Then NewLogLines sItemsLine
    WriteLines sItemsLine
 
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub WriteLines(ByVal w_lines As String)
' -----------------------------------------------------------------------------------
' Writes one to n lines (w_lines) to the log file. Lines are delimited by a vbCrLf.
' -----------------------------------------------------------------------------------
    Dim a As Variant
    Dim s As String
    
    a = Split(w_lines, vbCrLf)
    s = Join(a, vbCrLf & TimeStamp)
    StringAsFile TimeStamp & s, FileFullName, True
    
End Sub

