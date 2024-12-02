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
Private FSo                     As New FileSystemObject
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
    sFileBaseName = FSo.GetBaseName(ActiveWorkbook.FullName)
    sFileExtension = "log"
    sFileLocation = ActiveWorkbook.Path
    sFileFullName = FileFullName

End Sub

Private Sub Class_Terminate()
    '~~ In case the file already existed it may have passed the KeepLogs limit
    If FSo.FileExists(sFileFullName) Then
        Reorg sFileFullName, lKeepLogs
    End If

    Set FSo = Nothing
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

Public Property Get Arry(Optional ByRef c_arr As Variant, _
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

Public Property Let Arry(Optional ByRef c_arr As Variant, _
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
    
    On Error GoTo eh
    Dim lWidth  As Long
    
    If ArrayIsAllocated(aItems) Then
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
                                      ByVal c_item As String)
    Arry(aItems, c_col) = c_item
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
    
    With FSo
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
    AlignLeft = s
    If a_col_arranged Then AlignLeft = a_margin & s & a_margin _
    
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
      
    If Not ArrayIsAllocated(a_array) Then Exit Function
    
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
    ArrayIsAllocated = IsArray(a_arr) _
              And Not IsError(LBound(a_arr, 1)) _
              And LBound(a_arr, 1) <= UBound(a_arr, 1)
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

Private Function ColMarginsAdded(ByVal c_col As Long, _
                                 ByVal c_strng As String, _
                                 ByVal c_margin As String, _
                                 ByVal c_cols As Long) As String
' -----------------------------------------------------------------------------------
' Returns a string (c_strng) for a column (c_col) ammended with margins (c_margin),
' whereby the leftmost column (1) has no left margin and the rightmost gets no right
' margin. When the column (c_col) is 0 no margins are added (not a column alignment).
' -----------------------------------------------------------------------------------
                              
    If c_col = 0 Then
        ColMarginsAdded = c_strng
    Else
        If c_col = c_cols Then
            ColMarginsAdded = c_margin & c_strng
        Else
            ColMarginsAdded = c_margin & c_strng & c_margin
        End If
    End If
    
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
    Open f_file_full_name For Input As #1
    If Err.Number <> 0 Then GoTo xt
    FileAsString = Input$(lOf(1), 1)
    Close #1
    
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
    With FSo
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
        
        If FSo.FileExists(sFileFullName) Then WriteLines String(lWidth, "=")  ' log delimiter line
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
    Dim bAppend As Boolean
    Dim cllLog  As New Collection
    Dim cllLogs As New Collection
    Dim i       As Long
    Dim lLen    As Long
    Dim s       As String
    Dim v       As Variant
    Dim bTop    As Boolean: bTop = True
    
    If Not FSo.FileExists(r_file_full_name) Then GoTo xt
    
    '~~ Collect the logs (i.e. lines delimited by a log delimiter line)
    '~~ into a Collection with each series of log entries/lines as a Collection
    '~~ with the array and the max line length as items
    aFile = FileAsArray(r_file_full_name)
    ArrayErase aLog
    For i = LBound(aFile) To UBound(aFile)
        If IsDelimiterLine(aFile(i)) And Not i = LBound(aFile) Then
            '~~ A delimiter line indicates the start of a new series of log entries
            cllLog.Add aLog
            cllLog.Add lLen
            cllLogs.Add cllLog
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
    cllLogs.Add cllLog
    ArrayErase aLog
    Set cllLog = Nothing
    Set cllLog = New Collection
    
    '~~ Remove all excessive logs
    If cllLogs.Count < r_keep Then GoTo xt
    Do While cllLogs.Count > r_keep
        cllLogs.Remove 1
    Loop
    
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
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = FSo.GetFile(s_file)
    
End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = FSo.GetSpecialFolder(2)
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
    
    If Not ArrayIsAllocated(aTitleLines) Then GoTo xt
    For Each v In aTitleLines
        t_width = Max(t_width, Len(v))
    Next v
    
    For Each v In aTitleLines
        s = s & sSplit & Align(CStr(v), sTitleAlign, t_width, sTitleFill)
        sSplit = vbCrLf
    Next v
    TitlesAssembled = s
    ArrayErase aTitleLines
    
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
    TitlesSpecified = ArrayIsAllocated(aTitleLines)
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

