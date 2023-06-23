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
' - AlignmentItems   Explicitly specifies the alignment for each column's
'                    item, defaults to "L"eft adjusted when not provided.
' - ColsDelimiter    Defaults to | when Headers are specified and defaults to a
'                    single space when no Headers are specified.
' - Dsply            Displays the Log-Files content by means of ShellRun.
' - Entry            Either writes a single string, i.e. one without any any |
'                    and or subsequent string elements of writes an number of
'                    provided strings or one string with vertical bars, both
'                    indicating items to be writen alingned in columns.
' - Headers          Specifies the column headers
' - MaxItemLengths   Explicitly specifies each columns minimum width. The width of,
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
' - FileFullName  r/w     Specifies the full name of the log-file. The location
'                         defaults to the ActiveWorkbook's parent folder, the name
'                         defaults to the ActiveWorkbook's BaseName with a .log
'                         file extention.
' - FileName      w       Specifies the log-file's name, defaults to the
'                         ActiveWorkbook's BaseName with a .log file extention.
' - KeepDays      w       Specifies the number of days a new created log-file is
'                         kept before it is deleted and re-created, defaults to 1.
' - LogFile       r       The log-file as file object
' - Path          w       Sepcifies the log-file's location/folder, defaults to the
'                         ActiveWorkbook's parent folder
'
' W. Rauschenberger, Berlin Apr 2023
' See: https://github.com/warbe-maker/VBA-Log-Service
' -----------------------------------------------------------------------------------
Private aAlignmentsHeader   As Variant
Private aAlignmentsItem     As Variant
Private aHeaders            As Variant
Private aHeaderLines        As Variant
Private aItems              As Variant
Private aTitleLines         As Variant
Private aMaxItemLengths     As Variant
Private aFills()            As String
Private bColsAligned        As Boolean
Private bInitialized        As Boolean
Private bNewLog             As Boolean
Private bWithTimeStamp      As Boolean
Private fso                 As New FileSystemObject
Private lColumns            As Long
Private lKeepDays           As Long
Private lLenTitle           As Long
Private sTitleAlignment     As String
Private sColsDelimiter      As String
Private sColsMargin         As String
Private sEntry              As String
Private sFileFullName       As String
Private sFileName           As String
Private sPath               As String
Private sServicedItem       As String
Private sServicedItemName   As String
Private sServicedItemType   As String
Private sTitleFill          As String

#If Not MsgComp = 1 Then
    ' -------------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed which must
    ' be indicated by the Conditional Compile Argument MsgComp = 1.
    ' See https://github.com/warbe-maker/Common-VBA-Message-Service
    ' -------------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
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
    bWithTimeStamp = False
    sFileName = fso.GetBaseName(ActiveWorkbook.Name) & ".log"
    sPath = ActiveWorkbook.Path
    FileFullName = sPath & "\" & sFileName
    NewLog
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
    Set aAlignmentsHeader = Nothing
    Set aAlignmentsItem = Nothing
    Set aHeaders = Nothing
    Set aHeaderLines = Nothing
    Set aMaxItemLengths = Nothing
    Set aItems = Nothing
    Set aTitleLines = Nothing
    sColsDelimiter = vbNullString
End Sub

Private Property Get AlignmentHeaderColumn(Optional ByVal c_col As Long) As String
    If IsAllocated(aAlignmentsHeader) Then
        If c_col <= UBound(aAlignmentsHeader) Then
            AlignmentHeaderColumn = aAlignmentsHeader(c_col)
        End If
    End If
End Property

Private Property Let AlignmentHeaderColumn(Optional ByVal c_col As Long, _
                                                ByVal c_align As String)
    ArryRedim aAlignmentsHeader, c_col
    aAlignmentsHeader(c_col) = c_align
End Property

Public Property Get AlignmentItemColumn(Optional ByVal c_col As Long) As String
    Dim s   As String
    
    If IsAllocated(aAlignmentsItem) Then
        If c_col <= UBound(aAlignmentsItem) _
        Then s = aAlignmentsItem(c_col)
    End If
    
    AlignmentItemColumn = s

End Property

Private Property Let AlignmentItemColumn(Optional ByVal c_col As Long, _
                                               ByVal c_align As String)
    ArryRedim aAlignmentsItem, c_col
    If aAlignmentsItem(c_col) = vbNullString _
    Then aAlignmentsItem(c_col) = c_align

End Property

Private Property Get ColFill(Optional ByVal c_col As Long) As String
    If IsAllocated(aFills) Then
        If c_col <= UBound(aFills) _
        Then ColFill = aFills(c_col)
    End If
End Property

Private Property Let ColFill(Optional ByVal c_col As Long, _
                                      ByVal c_fill As String)
    ArryRedim aFills, c_col
    aFills(c_col) = c_fill
    
End Property

Private Property Get ColHeader(Optional ByVal c_col As Long, _
                               Optional ByVal c_line As Long = 1) As String
    Dim s       As String
    Dim lWidth  As Long
    Dim v       As Variant
    
    If IsAllocated(aHeaderLines) Then
        aHeaders = aHeaderLines(c_line)
        If c_col <= UBound(aHeaders) Then
            s = WithMargins(aHeaders(c_col))
            lWidth = Len(s)
            lWidth = Max(ColWidth(c_col), lWidth, ColWidthFinal(c_col)) ' width with marging may be greater than originally calculated
            ColHeader = AlignString(aHeaders(c_col), lWidth, AlignmentHeaderColumn(c_col), sColsMargin)
        End If
    End If
End Property

Private Property Let ColHeader(Optional ByVal c_col As Long, _
                               Optional ByVal c_line As Long = 1, _
                                        ByVal c_header As String)
    
    ArryRedim aHeaderLines, c_line
    ArryRedim aHeaders, c_col
    aHeaders(c_col) = c_header
    aHeaderLines(c_line) = aHeaders
    
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
    Dim sFill   As String
    
    If IsAllocated(aItems) Then
        If bNewLog Then
            '~~ When this is the first line of a new series of log entries the width
            '~~ may still be expanded when the item's width is greater than the
            '~~ already specified - in case
            lWidth = ColWidth(c_col)
            ColWidth(c_col) = Max(Len(aItems(c_col)), lWidth)
        End If
        
        If c_col < lColumns Then
            ColItem = AlignString(aItems(c_col), ColWidthFinal(c_col), AlignmentItemColumn(c_col), sColsMargin, ColFill(c_col))
        ElseIf c_col = lColumns Then
            '~~ The width of the rightmost column is unlimited and left aligned.
            ColItem = AlignString(aItems(c_col), Len(aItems(c_col)) + Len(sColsMargin), "L", sColsMargin, " ")
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
    ArryRedim aItems, c_col
    aItems(c_col) = c_item
End Property

Friend Property Let ColsDelimiter(ByVal s As String)
    Select Case s
        Case " ": sColsDelimiter = s: sColsMargin = vbNullString
        Case "|": sColsDelimiter = s: sColsMargin = " "
    End Select
End Property

Private Property Get ColWidth(Optional ByVal c_col As Long) As Long
    If IsAllocated(aMaxItemLengths) Then
        If c_col <= UBound(aMaxItemLengths) Then
            ColWidth = aMaxItemLengths(c_col)
        End If
    End If
End Property

Private Property Let ColWidth(Optional ByVal c_col As Long, _
                                       ByVal c_width As Long)
    ArryRedim aMaxItemLengths, c_col
    aMaxItemLengths(c_col) = c_width
End Property

'Private Property Get DefaultLogFileName() As String
'    DefaultLogFileName = fso.GetBaseName(ActiveWorkbook.Name) & ".log"
'End Property

Private Property Get DelimiterLineDue() As Boolean
    DelimiterLineDue = bNewLog _
                    Or TitleSpecified _
                    Or HeadersSpecified
End Property

Friend Property Get FileFullName() As String:   FileFullName = sFileFullName:   End Property

Friend Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Specifies the log file's name and location, by the way maintains the Path
' and the FileName property.
' ----------------------------------------------------------------------------
    Dim lDaysAge    As Long
    
    With fso
        If sFileFullName <> s Then
            '~~ The log file's name is not/no longer the one previously used
            If .FileExists(sFileFullName) Then .DeleteFile sFileFullName
        End If
        sFileFullName = s
        sPath = .GetParentFolderName(s)
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
        sFileName = .GetFileName(s)
        
        '~~ In case the file already existed it may have passed the KeepDays limit
        lDaysAge = VBA.DateDiff("d", .GetFile(sFileFullName).DateLastAccessed, Now())
        If lDaysAge > lKeepDays Then
            .DeleteFile sFileFullName
            .CreateTextFile sFileFullName
        End If
        
    End With

End Property

Friend Property Get FileName() As String:       FileName = sFileName:           End Property

Friend Property Let FileName(ByVal s As String)
    If fso.GetExtensionName(s) = vbNullString Then
        s = s & ".log"
    End If
    sFileName = s
    sFileFullName = Replace(sPath & "\" & sFileName, "\\", "\")
End Property

Private Property Get HeadersSpecified() As Boolean
    HeadersSpecified = IsAllocated(aHeaderLines)
End Property

Private Property Let HeadersSpecified(ByVal b As Boolean)
    If Not b Then
        Set aHeaderLines = Nothing
        Set aHeaders = Nothing
    End If
End Property

Friend Property Let KeepDays(ByVal l As Long):  lKeepDays = l:  End Property

Friend Property Get LogFile() As File
' -----------------------------------------------------------------------------------
' Returns the log file as file object
' -----------------------------------------------------------------------------------
    With New FileSystemObject
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
        Set LogFile = .GetFile(sFileFullName)
    End With

End Property

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
            sPath = v
        Case VarType(v) = vbObject
            If TypeOf v Is Workbook Then
                Set wbk = v
                sPath = wbk.Path
            ElseIf TypeOf v Is Folder Then
                Set fld = v
                sPath = fld.Path
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a string specifying a " & _
                                                   "folder's path, nor a Workbook object, nor a Folder object!"
            End If
    End Select
    
End Property

Private Property Get TitleSpecified() As Boolean
    TitleSpecified = IsAllocated(aTitleLines)
End Property

Private Property Let TitleSpecified(ByVal b As Boolean)
    If Not b Then Set aTitleLines = Nothing
End Property

Public Sub Alignment(ByVal a_string As String, _
            Optional ByRef a_return As String, _
            Optional ByRef a_fill As String, _
            Optional ByRef a_align As String)
' -----------------------------------------------------------------------------------
' Returns the alignment derived from a string (a_string), whereby the string may be
' an item (implicit alignment spec) or an explicit alignment spec.
' When a_align = vbNullString the following is returned:
'      - the alignment specified implicitly (a_align) in accordance with the
'        following rules:
'        - (L)eft adjusted  = Leading spaces > trailing spaces
'        - (C)entered       = Leading spaces = trailing spaces
'        - (R)ight adjusted = Leading spaces < trailing spaces.
'      - the fill string (a_fill) if any
'      - the provided string (a_string) unstripped from any implicit alignment/fill
'        specs whereby leading spaces are preserved with "L"eft adjusted and
'        trailing spaces are preserved with "R"ight adjusted.
'
'      Examples:  |   (L)eft  |(C)entered |  (R)ight  | Fill |
'                 |-----------+-----------+-----------+------|
'                 | "xxx "    | "xxx"     | " xxxx"   |      |
'                 | " xxxx  " | " xxxx "  | "  xxxx " |
'                 | "xxx ."   |           |           | "."
'                 | "xxx .:"  |           |           | ".:"
'                 |           | "- xxx -" |           | "-"
' Else (a_align <> vbNullString)
'      The string (a_align) is returned (a_return) with leading and/or trailing
'      spaces unstripped in accordance with the provided alignment (a_align).
'
' Important to understand:
' In case an alignment (a_align) is already provided only the returned string
' is unstripped from irrelevant leading and/or trailing spaces.
' -----------------------------------------------------------------------------------
    Dim lLeft   As Long
    Dim lRight  As Long
    Dim i       As Long
    
    a_fill = " " ' default
    Select Case True
        Case a_string = "L", a_string = "R", a_string = "C"
            a_align = a_string

        Case Right$(a_string, 3) = " .:"
            If a_align = vbNullString Then
                a_align = "L"
                a_fill = ".:"
            End If
            a_return = Left$(a_string, Len(a_string) - 3)   ' unstrip the alignment/fill spec
        
        Case Right$(a_string, 2) = ".:"
            If a_align = vbNullString Then
                a_align = "L"
                a_fill = ".:"
            End If
            a_return = Left$(a_string, Len(a_string) - 2)   ' unstrip the alignment/fill spec
        
        Case Right$(a_string, 2) = " ."
            If a_align = vbNullString Then
                a_align = "L"
                a_fill = "."
            End If
            a_return = Left$(a_string, Len(a_string) - 2)   ' unstrip the alignment/fill spec
        
        Case Right$(a_string, 1) = "."
            If a_align = vbNullString Then
                a_align = "L"
                a_fill = "."
            End If
            a_return = Left$(a_string, Len(a_string) - 1)   ' unstrip the alignment/fill spec
        
        Case Left$(a_string, 2) = "- " And Right$(a_string, 2) = " -"
            If a_align = vbNullString Then
                a_align = "C"
                a_fill = "-"
            End If
            a_return = Mid$(a_string, 3, Len(a_string) - 4) ' unstrip the alignment/fill spec
        Case Left$(a_string, 1) = "-" And Right$(a_string, 1) = "-"
            If a_align = vbNullString Then
                a_align = "C"
                a_fill = "-"
            End If
            a_return = Mid$(a_string, 2, Len(a_string) - 2) ' unstrip the alignment/fill spec
        
        Case Else
            '~~ No fill string is provided with the implicit specified alignment
            If a_align = vbNullString Then
                '~~ The alignment is still to be returned
                For i = 1 To Len(a_string)
                    '~~ Count leading spaces
                    If i = 1 Then
                        If Left(a_string, 1) <> " " Then Exit For
                    Else
                        If Mid$(a_string, i, 1) <> " " Then Exit For
                    End If
                    lLeft = lLeft + 1
                Next i
                For i = Len(a_string) To 1 Step -1
                    '~~ Count trailing spaces
                    If i = Len(a_string) Then
                        If Right(a_string, 1) <> " " Then Exit For
                    Else
                        If Mid$(a_string, i, 1) <> " " Then Exit For
                    End If
                    lRight = lRight + 1
                Next i
                Select Case True
                    Case lLeft = lRight:    If Not a_align = "L" Then a_align = "C"
                    Case lLeft < lRight:    a_align = "L"
                    Case lLeft > lRight:    a_align = "R"
                End Select
            End If
            
            '~~ Untrip leading and/or trailing spaces in accordance with the alignment (analysed or provide)
            Select Case a_align
                Case "C":    a_return = Trim$(a_string)  ' unstrip any leading and trailing spaces
                Case "L":    a_return = RTrim$(a_string) ' unstrip trailing spaces
                Case "R":    a_return = LTrim$(a_string) ' unstrip leading spaces
            End Select
    
    End Select

End Sub

Public Sub AlignmentHeaders(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' Explicit headers column alignment specification, overwrites any already specified
' alignment (possibly implicitly specified).
' character is neither L, nor C, nor R, the alignment is considered "implicit",
' indicated by the number of left and right spaces.
' -----------------------------------------------------------------------------------
    Const PROC = "AlignmentHeaders"
    
    On Error GoTo eh
    Dim i       As Long
    Dim iCol    As Long
    Dim sAlign  As String
    Dim sReturn As String
    Dim sFill   As String
    Dim s       As String
    Dim v       As Variant
        
    If Not IsError(c_align(LBound(c_align))) Then
        If InStr(c_align(LBound(c_align)), "|") <> 0 Then
            '~~ The | in the string indicates an alternative how column aligned
            '~~ headers may be specified
            s = c_align(LBound(c_align))
            If Left(s, 1) = "|" Then s = Right(s, Len(s) - 1)
            If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
            v = Split(s, "|")
        Else
            '~~ The headers are all elements of the ParamArray
            v = c_align
        End If
    Else
        v = c_align
    End If
    
    For i = LBound(v) To UBound(v)
        If Not IsError(v(i)) Then
            iCol = i + 1
            sAlign = vbNullString
            If Not IsError(v(i)) Then
                '~~ Skippes not specified columns
                Select Case UCase$(v(i))
                    Case "L", "C", "R"
                        AlignmentHeaderColumn(iCol) = UCase(v(i))
                    Case Else
                        Alignment v(i), sReturn, sFill, sAlign
                        Select Case UCase(Left$(sAlign, 1))
                            Case "L", "C", "R"
                                AlignmentHeaderColumn(iCol) = UCase(Left$(sAlign, 1))
                            Case Else
                                Select Case UCase(Left$(sReturn, 1))
                                    Case "L", "C", "R"
                                        AlignmentHeaderColumn(iCol) = UCase(Left$(sAlign, 1))
                                    Case Else
                                        Err.Raise AppErr(1), ErrSrc(PROC), "The alignment for column " & iCol & " is neither L, nor R, nor C!"
                                End Select
                        End Select
                End Select
            End If
        End If
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AlignmentItems(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' Explicit specification of the items column alignment. When no explicit alignment is
' is specified or the first character is neither L, nor C, nor R, the alignment
' default to the implicit specification of the first line's items.
' Valid specifications: Spec        Meaning
'                       ----------- -------------------------------
'                       L           Left adjusted no fill
'                       L.          Left adjusted filled with ....:
'                       R           Right adjusted
'                       C           Centered
'                       -C-, C-, -C Centered with - fills
'
' -----------------------------------------------------------------------------------
    Const PROC = "AlignmentItems"
    Dim i       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim sAlign  As String
    Dim sFill   As String
    Dim v       As Variant
    Dim sReturn As String
    
    On Error GoTo eh
    If Not IsError(c_align(LBound(c_align))) Then
        '~~ IsError indicates a columns alignment 'not explicitly specified'.
        '~~ The corresponding columns will be aligned implicitly specified
        '~~ with the first entry's items specifications
        s = c_align(LBound(c_align))
        If InStr(s, "|") <> 0 Then
            If Left(s, 1) = "|" Then s = Right(s, Len(s) - 1)
            If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
            v = Split(s, "|")
        Else
            v = c_align
        End If
    Else
        v = c_align
    End If
                
    For i = LBound(v) To UBound(v)
        iCol = i + 1
        If Not IsError(v(i)) Then
            sFill = " "
            sAlign = vbNullString
            Alignment v(i), sReturn, sFill, sAlign
            Select Case True
                Case sAlign = "L" And sFill = " "
                Case sAlign = "L" And sFill = "."
                Case sAlign = "L" And sFill = ".:"
                Case sAlign = "R"
                Case sAlign = "C" And sFill = " "
                Case sAlign = "C" And sFill = "-"
                Case Else: Err.Raise AppErr(1), ErrSrc(PROC), "The specified alignment is invalid! Valid specs are:" & vbLf & _
                                                              "L, L., R, C, -C, C-, -C-"
            End Select
            AlignmentItemColumn(iCol) = sAlign
            ColFill(iCol) = sFill
        Else
            AlignmentItemColumn(iCol) = vbNullString
            ColFill(iCol) = " "
        End If
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function AlignString(ByVal a_strng As String, _
                            ByVal a_lngth As Long, _
                   Optional ByVal a_align As String = "L", _
                   Optional ByVal a_margn As String = vbNullString, _
                   Optional ByVal a_fill As String = " ") As String
' ----------------------------------------------------------------------------
' Returns a string (a_strng) with a lenght (a_lngth) aligned (a_align) and
' filled with (a_fill).
' Preconditions:
' - The length (a_lngth) is the final lenght including any specific fillss
' - The string (a_strn) may contain already leading or trailing preserved
'   spaces.
' ----------------------------------------------------------------------------
    Const PROC = "AlignString"
    
    On Error GoTo eh
    Dim sFill   As String
    Dim s       As String

    If Trim$(a_strng) = vbNullString Then
        AlignString = String$(a_lngth, " ")
    Else
        If a_fill = vbNullString Then a_fill = " "
        Select Case Left(UCase(a_align), 1)
            
            Case "L"
                a_strng = WithMargins(a_strng)
                Select Case a_fill
                    Case ".:"
                        If Right$(a_strng, 1) <> " " Then a_strng = a_strng & " "
                        AlignString = Left$(a_strng & String(a_lngth, a_fill), a_lngth - 1) & ":"
                    Case "."
                        If Right$(a_strng, 1) <> " " Then a_strng = a_strng & " "
                        AlignString = Left$(a_strng & String(a_lngth, a_fill), a_lngth)
                    Case Else
                        AlignString = Left$(a_strng & String(a_lngth, a_fill), a_lngth)
                End Select
            
            Case "R"
                AlignString = Right$(String(a_lngth, a_fill) & a_strng & a_margn, a_lngth)
                
            Case "C"
                s = " " & Trim$(a_strng) & " "
                If Len(s) < a_lngth Then
                    sFill = String$(Int((a_lngth - Len(s)) / 2), a_fill)
                End If
                    AlignString = sFill & s & sFill
                If Len(AlignString) < a_lngth Then
                    AlignString = AlignString & VBA.String$(a_lngth - Len(AlignString), a_fill)
                End If
        End Select
    End If

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
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Private Sub ArryRedim(ByRef a_arry As Variant, _
                      ByVal a_items As Long)
' -----------------------------------------------------------------------------------
' Returns the array (a_arry) redimed with the number of required columns (a_items).
' -----------------------------------------------------------------------------------
    If Not IsAllocated(a_arry) _
    Then ReDim a_arry(a_items) _
    Else ReDim Preserve a_arry(Max(UBound(a_arry), a_items))
End Sub

Public Function AsArray(ByVal a_string As String, _
                        ByVal a_split_by As String) As String()
' -----------------------------------------------------------------------------------
' Returns a string (a_string) as an array of strings considering a split character
' (a_split_by). Any leading or trailing split character is removed before the split.
' -----------------------------------------------------------------------------------
    If InStr(a_string, a_split_by) <> 0 Then
        If Left(a_string, 1) = a_split_by Then a_string = Right(a_string, Len(a_string) - 1)
        If Right(a_string, 1) = a_split_by Then a_string = Left(a_string, Len(a_string) - 1)
    End If
    AsArray = Split(a_string, a_split_by)
End Function

Private Function ColWidthFinal(ByVal c_col As Long) As Long
' -----------------------------------------------------------------------------------
' Returns the final width of a column by considering the either implicitly or
' explicitly specified width and the fill string in case one is specified.
' -----------------------------------------------------------------------------------
    Select Case ColFill(c_col)
        Case ".":   ColWidthFinal = ColWidth(c_col) + 2
        Case ".:":  ColWidthFinal = ColWidth(c_col) + 3
        Case Else:  ColWidthFinal = ColWidth(c_col)
    End Select
End Function

Public Sub Dsply(Optional ByVal d_file_full_name = vbNullString)
' -----------------------------------------------------------------------------------
' Displays the Log-Files content, whereby the log-file defaults to sFileFullName.
' -----------------------------------------------------------------------------------
    If d_file_full_name = vbNullString Then d_file_full_name = sFileFullName
    ShellRun d_file_full_name
End Sub

Public Sub Entry(ParamArray e() As Variant)
' ----------------------------------------------------------------------------
' - When a single string without any | is provided the string is appended as a
'   new line to the log file
' - When multiple strings are provided or a string contains vertical bars (|)
'   the indicated items are writen aligned in columns to the log-file
' - When an item is not available (IsError) an item with spaces only is
'   assumed
' ----------------------------------------------------------------------------
    Const PROC = "Entry"
    
    On Error GoTo eh
    Dim i       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim v       As Variant
    Dim lWidth  As Long
    Dim sItem   As String
    Dim sFill   As String
    Dim sAlign  As String
    
    bInitialized = False
    
    If LBound(e) = UBound(e) Then
        '~~ The log-entry is only one single string
        If Not IsError(e(LBound(e))) Then
            s = e(LBound(e))
            If InStr(s, "|") <> 0 Then
                '~~ The | in the string indicates items for being column aligned
                '~~ and additionally the implicit width
                If Left(s, 1) = "|" Then s = Right(s, Len(s) - 1)
                If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
                v = Split(s, "|")
                lColumns = UBound(v)
                For i = LBound(v) To UBound(v)
                    iCol = i + 1
                    ColItem(iCol) = CStr(v(i))
                    If ColWidth(iCol) = 0 Then
                        '~~ Only in case no column width has explicitly been
                        '~~ spcified by means of the MaxItemLengths method
                        ColWidth(iCol) = Len(v(i))
                    End If
                    sAlign = AlignmentItemColumn(iCol)
                    Alignment v(i), sItem, sFill, sAlign
                Next i
                WriteItems
            Else
                NewLogLines s
                WriteLine s
            End If
        End If
    Else
        '~~ The Entry provides more than just one string, thereby
        '~~ indicating that these items are to be aligned in columns
        lColumns = UBound(e) + 1
        For i = LBound(e) To UBound(e)
            iCol = i + 1
            If Not IsError(e(i)) Then
                sAlign = AlignmentItemColumn(iCol)
                sFill = ColFill(iCol)
                If sAlign = vbNullString Then
                    '~~ Not yet (explicitly) specified
                    '~~ Alignment returns:
                    '~~ - sItem with irrelevant leading and/or trailing spaces unstripped
                    '~~ - sAlign implicitly specified
                    '~~ - sFill implicitly specified
                    Alignment e(i), sItem, sFill, sAlign
                    AlignmentItemColumn(iCol) = sAlign
                    ColFill(iCol) = sFill
                Else
                    '~~ Already (explicitly) specified
                    '~~ Alignment (only) returns sItem with irrelevant leading and/or trailing spaces unstripped
                    Alignment e(i), sItem, sFill, sAlign
                End If
                ColItem(iCol) = sItem
                If ColWidth(iCol) = 0 Then
                    '~~ Only in case no column width has explicitly been
                    '~~ spcified by means of the MaxItemLengths method
                    ColWidth(iCol) = Len(sItem)
                End If
            Else
                '~~ Any not provided item becomes one with spaces only
                ColItem(iCol) = String(ColWidth(iCol), " ")
            End If
        Next i
        WriteItems
    End If
        
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
' Universal error message display service which displays a debugging option
' button when the Conditional Compile Argument 'Debugging = 1', displays an
' optional additional "About:" section when the err_dscrptn has an additional
' string concatenated by two vertical bars (||), and displays the error message
' by means of VBA.MsgBox when neither the Common Component mErH (indicated by
' the Conditional Compile Argument "ErHComp = 1", nor the Common Component mMsg
' (idicated by the Conditional Compile Argument "MsgComp = 1") is installed.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'       ErrSrc  To provide an unambiguous procedure name by prefixing is with
'               the module name.
'
' W. Rauschenberger Berlin, Apr 2023
'
' See: https://github.com/warbe-maker/Common-VBA-Error-Services
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
    If err_source = vbNullString Then err_source = Err.source
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
    ErrSrc = "clsLog" & "." & sProc
End Function

Public Sub Headers(ParamArray h_col_headers() As Variant)
' -----------------------------------------------------------------------------------
' Spacifies the column headers for Entry items aligne in columns.
' -----------------------------------------------------------------------------------
    Const PROC = "Headers"
    
    On Error GoTo eh
    Static lLine As Long
    Dim i       As Long
    Dim j       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim v       As Variant
    Dim lWidth  As Long
    Dim sAlign  As String
    Dim sHeader As String
    
    If Not HeadersSpecified Then
        lLine = 0
        NewLog ' Header indicated a newseries of log entries even when no title had been specified
    End If

    '~~ When column headers are specified and the ColsDelimiter is still not
    '~~ specified the delimiter defaults to a vertical bar.
    Select Case sColsDelimiter
        Case vbNullString:      sColsMargin = " ":          sColsDelimiter = "|"
        Case "|":               sColsMargin = " "
        Case " ":               sColsMargin = vbNullString
    End Select
    
    If Not IsError(h_col_headers(LBound(h_col_headers))) Then
        s = h_col_headers(LBound(h_col_headers))
        If InStr(s, "|") <> 0 Then
            '~~ The | in the string indicates an alternative how column aligned
            '~~ headers may be specified
            If Left(s, 1) = "|" Then s = Right(s, Len(s) - 1) ' remove trailing vertical bar - if any
            If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1) ' remove leading vertical bar - if any
            v = Split(s, "|")
        Else
            '~~ The headers are all elements of the ParamArray
            v = h_col_headers
        End If
    Else
        v = h_col_headers
    End If
    
    lLine = lLine + 1
    For i = LBound(v) To UBound(v)
        If Not IsError(v(i)) Then
            iCol = i + 1
            '~~ Skip column without a specified header
            sAlign = vbNullString
            Alignment v(i), sHeader, , sAlign
            If AlignmentHeaderColumn(iCol) = vbNullString Then
                '~~ Implicit alignment when none has yet been specified
                '~~ - e.g. explicitly or with a previous header line
                AlignmentHeaderColumn(iCol) = sAlign
            End If
            ColHeader(iCol, lLine) = sHeader
            lWidth = ColWidth(iCol)
            ColWidth(iCol) = Max(Len(WithMargins(sHeader)), lWidth)
            bNewLog = True ' the header specification indicates the begin of a new series of log entries
        End If
    Next i
    Set aHeaders = Nothing
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HeadersAssembleLines(ByRef h_lines As String)
' -----------------------------------------------------------------------------------
' Assemble column headers, header lines
' -----------------------------------------------------------------------------------
    Dim lLine       As Long
    Dim i           As Long
    Dim sLineBreak  As String
    Dim sHeaderLine As String
    
    sHeaderLine = vbNullString
    sLineBreak = vbNullString
    h_lines = vbNullString
    
    If HeadersSpecified Then
        For lLine = LBound(aHeaderLines) To UBound(aHeaderLines)
            aHeaders = aHeaderLines(lLine)
            For i = LBound(aHeaders) To UBound(aHeaders)
                sHeaderLine = sHeaderLine & sColsDelimiter & ColHeader(i, lLine)
            Next i
            h_lines = h_lines & sLineBreak & sHeaderLine
            sHeaderLine = vbNullString
            sLineBreak = vbLf
        Next lLine
    End If

End Sub

Private Sub HeadersWrite(ByVal h_lines As String, _
                         ByVal h_width_entry As Long)
' -----------------------------------------------------------------------------------
' Write header lines and a header underline
' -----------------------------------------------------------------------------------
    Dim vLine   As Variant
    Dim s       As String
    Dim sDelim  As String
    Dim i       As Long
    
    For Each vLine In Split(h_lines, vbLf)
        WriteLine vLine
    Next vLine
    
    '~~ Write header underline
    If sColsDelimiter = "|" _
    Then sDelim = "+" _
    Else sDelim = " "
    
    aHeaders = aHeaderLines(1)
    For i = LBound(aHeaders) To UBound(aHeaders)
        s = s & sDelim & String(Len(ColHeader(i, 1)), "-")
    Next i
    If Len(s) < h_width_entry Then
        s = s & String(h_width_entry - Len(s), "-")
    End If
    WriteLine s
    Set aHeaders = Nothing

End Sub

Private Function IsAllocated(arr As Variant) As Boolean
' -----------------------------------------------------------------------------------
' Retruns TRUE when an array (arr) is allocated, i.e. when it has a valid LBound and
' and an LBound which is less or equal its UBound.
' -----------------------------------------------------------------------------------
    On Error Resume Next
    IsAllocated = IsArray(arr) _
              And Not IsError(LBound(arr, 1)) _
              And LBound(arr, 1) <= UBound(arr, 1)
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

Public Sub MaxItemLengths(ParamArray c_widths() As Variant)
' -----------------------------------------------------------------------------------
' Specifies the maximum length for the items aligned in columns which will become the
' minimum column width. The corresponding final columns widths may be greater however.
' -----------------------------------------------------------------------------------
    Const PROC = "MaxItemLengths"
    
    On Error GoTo eh
    Dim i       As Long
    Dim iCol    As Long
    
    For i = LBound(c_widths) To UBound(c_widths)
        If Not IsError(c_widths(i)) Then
            If Not IsError(Abs(CInt(c_widths(i)))) Then
                iCol = i + 1
                ColWidth(iCol) = c_widths(i)
            End If
        End If
    Next i

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
        .CreateTextFile sFileFullName
    End With
    NewLog
End Sub

Public Sub NewLog(Optional ByVal b As Boolean = True)
' -----------------------------------------------------------------------------------
' Explicitly (by default) starts a new series of log entries even when neither Title
' nor Headers are specified.
' Attention: Even when called with FALSE to indicate that no delimiter line is
'            desired, this still triggers the initialization of a new series of log
'            entried by means of NewlogInit
' -----------------------------------------------------------------------------------
    bNewLog = b
    If Not bInitialized Then NewLogInit
End Sub

Private Sub NewLogInit()
' -----------------------------------------------------------------------------------
' Essential for a new series of log entries. Triggered either by NewLog, Title, or
' Headers. bInitialized is reset to false with the next subsequent Entry method call.
'
' Attention: This is the initialization of a new series of log entries - even when
'            NewLog False had been called in order to suppress the delimiter line.
' -----------------------------------------------------------------------------------
    If Not bInitialized Then
        aAlignmentsHeader = vbNullString
        aAlignmentsItem = vbNullString
        aItems = vbNullString
        sColsDelimiter = vbNullString   ' Default in case no Headers are specified
        sColsMargin = vbNullString      ' Default in case no Headers are specified
        bInitialized = True
    End If
End Sub

Private Sub NewLogLines(ByVal n_line As String)

    Dim sElement        As String
    Dim lWidthEntry     As Long
    Dim sHeaderLines    As String
    Dim sTitleLines     As String
    Dim vLine           As Variant

    '~~ Assemble the headerlines
    If HeadersSpecified Then
        HeadersAssembleLines sHeaderLines
    End If
    
    lLenTitle = Max(lLenTitle, Len(n_line))
    lWidthEntry = Max(lLenTitle, lWidthEntry, Len(n_line))
    
    '~~ Assemble the title lines
    If TitleSpecified Then
        TitleAssembleLines n_line, lWidthEntry, sTitleLines
    End If
    
    If DelimiterLineDue Then
        WriteDelimiterLine lWidthEntry
    End If
    
    If TitleSpecified Then
        For Each vLine In Split(sTitleLines, vbLf)
            WriteLine vLine
        Next vLine
        WriteLine String(lWidthEntry, "-") ' title underline
        TitleSpecified = False
    End If
      
    If HeadersSpecified Then
        HeadersWrite sHeaderLines, lWidthEntry
        HeadersSpecified = False
    End If
    bNewLog = False ' may be set to True again when a new Title is provided or explicit
    
End Sub

Private Sub ProvideLogFile()
    Dim lDaysAge As Long
    Dim ts       As TextStream
    With fso
        If Not .FileExists(sFileFullName) Then
            Set ts = .CreateTextFile(sFileFullName)
            ts.Close
        Else
            lDaysAge = VBA.DateDiff("d", .GetFile(sFileFullName).DateLastAccessed, Now())
            If lDaysAge > lKeepDays Then
                .DeleteFile sFileFullName
                .CreateTextFile sFileFullName
            End If
        End If
    End With
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
    Dim stRet           As String
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

Private Function TimeStamp() As String ' Public for test purpose only
    TimeStamp = " "
    If bWithTimeStamp Then TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & TimeStamp
End Function

Public Sub Title(ParamArray t_lines() As Variant)
' ----------------------------------------------------------------------------
' Method for specifying an - optionally multi-lined - title for an
' (implicitly specified) new series of log entries.
' Note 1: The first title line implicitly specifies the title's alignment.
' Note 2: Multiple lines may be specified with one method call (with serveral
'         ParamArray items) or with several method calls (each for one line).
' ----------------------------------------------------------------------------
    Static lLine    As Long
    Dim i           As Long
    Dim sTitle      As String
    
    If Not TitleSpecified Then
        lLine = 0
        lLenTitle = 0
        sTitleAlignment = vbNullString
        sTitleFill = " "
        NewLog
    End If
        
    For i = LBound(t_lines) To UBound(t_lines)
        lLine = lLine + 1
        ArryRedim aTitleLines, lLine
        Alignment t_lines(i), sTitle, sTitleFill, sTitleAlignment
        lLenTitle = Max(lLenTitle, Len(sTitle))
        aTitleLines(lLine) = sTitle
    Next i
    aItems = vbNullString ' a new title clears any previous items

End Sub

Private Sub TitleAssembleLines(ByVal t_line As String, _
                               ByVal t_width_entry As Long, _
                               ByRef t_lines As String)
    Dim sLineBreak  As String
    Dim i           As Long
    
    For i = LBound(aTitleLines) To UBound(aTitleLines)
        If aTitleLines(i) <> vbNullString Then
            t_lines = t_lines & sLineBreak & AlignString(aTitleLines(i), t_width_entry, sTitleAlignment, vbNullString, sTitleFill)
            sLineBreak = vbLf
        End If
    Next i

End Sub

Private Function WithMargins(ByVal w_string As String) As String
' -----------------------------------------------------------------------------------
' Returns a string (w_string) with a leading and trailing margin (w_margin) in case
' the string (w_string) not already begins and/or ends with a currently specified
' margin (sColsMargin). A margin = vbNullString is ignored.
' -----------------------------------------------------------------------------------
    Dim s As String
    If sColsMargin <> vbNullString Then
        If Left(w_string, 1) <> sColsMargin _
        Then s = sColsMargin & w_string _
        Else s = w_string
        If Right(s, 1) <> sColsMargin _
        Then s = s & sColsMargin
        WithMargins = s
    Else
        WithMargins = w_string
    End If
    
End Function

Public Sub WithTimeStamp(Optional ByVal b As Boolean = True)
    bWithTimeStamp = b
End Sub

Private Sub WriteDelimiterLine(ByVal w_width_entry As Long)
' -----------------------------------------------------------------------------------
' Write a delimiter line (=====) when it is due and the logfile is not empty
' -----------------------------------------------------------------------------------
    If DelimiterLineDue And fso.GetFile(sFileFullName).Size <> 0 Then
        WriteLine String(w_width_entry, "=")
    End If
End Sub

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
    
    NewLogLines sItemsLine  ' if any
    WriteLine sItemsLine
 
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub WriteLine(ByVal w_string As String)
' ----------------------------------------------------------------------------
' Writes a string (w_string) into the file (sFileFullName).
' ----------------------------------------------------------------------------
    Const PROC = "WriteLine"
    
    On Error GoTo eh
    Dim ts  As TextStream
   
    Set ts = fso.OpenTextFile(FileName:=sFileFullName, IOMode:=ForAppending)
    ts.WriteLine TimeStamp & w_string

xt: ts.Close
    Set ts = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

