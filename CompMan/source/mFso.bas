Attribute VB_Name = "mFso"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard  Module mFso: Common services regarding files and folders. The
' ====================== module makes extensive use of the FileSystemObject
' by extending and supplementing its services. The componet works autonomous
' and does not require any other components.
'
' Public properties and services:
' -------------------------------
' Exists                Universal existence function. Returns TRUE when the
'                       folder, file, section ([xxxx]), or value name exists.
' FileCompareByWinMerge Displays the differences of two files by using
'                       WinMerge
' FileDiffers           Returns a Dictionary with those records/lines which
'                       differ between two provided files, with the options
'                       'ignore case' and 'ignore empty records'
' FileDelete            Deletes a file provided either as object or as full
'                       name
' FileExtension         Returns the extension of a file's name
' Folders
' FileArry          r/w Returns the content of a text file as an array or
'                       writes the content of an array to a text file.
' FileString        r/w Returns the content of a text file as text string
'                       or writes a string to a text file - optionally
'                       appended
' FileTemp
' RenameSubFolders      Renames all folders and sub-folders in a provided
'                       path from a given old to a given new name.
' SortCutTargetPath r   Returns a provided shortcut's target path
'                       Let Creates a shortcut at a provided location with a
'                       provided path
'
' Requires References to:
' - Microsoft Scripting Runtine
' - Windows Script Host Object Model
'
' Uses no other components. Optionally uses mErH, fMsg/mMsg when installed and
' activated (Cond. Comp. Args. `mErH = 1 : mMsg = 1`).
'
' W. Rauschenberger, Berlin Apr 2024
' See also https://github.com/warbe-maker/VBA-File-System-Objects.
' ----------------------------------------------------------------------------
Public FSo                      As New FileSystemObject

Private Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-File-System-Objects"

#If mMsg = 0 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' approach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Conditional Compile Argument mMsg = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Const DQUOTE     As String = """" ' one " character

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

Public Property Get FileArry(Optional ByVal f_file_full_name As String, _
                             Optional ByVal f_exclude_empty As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file) as array. When the file
' (f_file_full_name) does not exist, an error is passed on to the caller.
' ----------------------------------------------------------------------------
    Const PROC  As String = "FileArry"
    
    Dim sSplit  As String
    Dim sFile   As String
    
    If Not FSo.FileExists(f_file_full_name) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named '" & f_file_full_name & "' does not exist!"
    
    sFile = FileString(f_file_full_name)    '~~ Obtain (and return) the used line terminating string (vbCrLf) or character
    sSplit = SplitString(sFile)
    FileArry = Split(sFile, sSplit)

End Property

Public Property Let FileArry(Optional ByVal f_file_full_name As String, _
                             Optional ByVal f_exclude_empty As Boolean = False, _
                                      ByVal f_ar As Variant)
' ----------------------------------------------------------------------------
' Writes array (f_ar) to file (f_file) whereby the array is joined to a text
' string using vbCrLf as line break string.
' ----------------------------------------------------------------------------
    
    f_exclude_empty = f_exclude_empty
    mFso.FileString(f_file_full_name) = Join(f_ar, vbCrLf)
             
End Property

Public Property Get FileDict(ByVal f_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Returns the content of the file (f_file) - which may be provided as file
' object or full file name - as Dictionary by considering any kind of line
' break characters.
' ----------------------------------------------------------------------------
    Const PROC  As String = "FileDict-Get"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim flo     As File
    Dim sFile   As String
    Dim i       As Long
    
    Select Case TypeName(f_file)
        Case "File"
            If Not FSo.FileExists(f_file.Path) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (f_file) does not exist!"
            Set flo = f_file
        Case "String"
            If Not FSo.FileExists(f_file) _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided file string (f_file) is not an existing file!"
            Set flo = FSo.GetFile(f_file)
    End Select
    
    '~~ Unload file into a test stream
    sFile = FileString(flo.Path)
    If sFile = vbNullString Then GoTo xt
    sSplit = SplitString(sFile)
    
    '~~ Test stream to array
    a = Split(sFile, sSplit)
    
    '~~ Remove any leading or trailing empty items
    ArrayTrimm a
    
    For i = LBound(a) To UBound(a)
        dct.Add i + 1, a(i)
    Next i
        
xt: Set FileDict = dct
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get FileString(Optional ByVal f_file_full_name As String, _
                               Optional ByVal f_append As Boolean = False, _
                               Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string with
' a vbCrLf as line terminating string - disregarding the optional split string
' (sSplit) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------
    Dim sSplit  As String
    Dim sFile   As String
    
    Open f_file_full_name For Input As #1
    sFile = Input$(lOf(1), 1)
    Close #1
    
    sSplit = SplitString(sFile)
    
    '~~ Eliminate a trailing eof if any
    If Right(sFile, 1) = VBA.Chr(26) Then
        sFile = Left(sFile, Len(sFile) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    Do While Right(sFile, Len(sSplit)) = sSplit
        DoEvents
        sFile = Left(sFile, Len(sFile) - Len(sSplit))
    Loop
    
    If f_exclude_empty Then
        '~~ Eliminate empty lines
        FileString = StringEmptyExcluded(sFile)
    Else
        FileString = sFile
    End If
        
End Property

Public Property Let FileString(Optional ByVal f_file_full_name As String, _
                               Optional ByVal f_append As Boolean = False, _
                               Optional ByVal f_exclude_empty As Boolean = False, _
                                        ByVal f_s As String)
' ----------------------------------------------------------------------------
' Writes a string (f_s) with multiple records/lines delimited by a vbCrLf to
' a file (f_file_full_name).
' ----------------------------------------------------------------------------
    
    If f_append _
    Then Open f_file_full_name For Append As #1 _
    Else Open f_file_full_name For Output As #1
    Print #1, f_s
    Close #1
        
End Property

Public Property Get FileTemp(Optional ByVal f_path As String = vbNullString, _
                             Optional ByVal f_extension As String = ".tmp") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    FileTemp = sTemp
    
End Property

Private Property Let Push(Optional ByRef s_stk As Collection, _
                                   ByVal v As Variant)
    s_stk.Add v
End Property

Public Property Get ShortCutTargetPath(Optional ByVal s_shortcut As String) As String
' ----------------------------------------------------------------------------
' Returns a provided shortcut's target path.
' ----------------------------------------------------------------------------
    Dim oShell      As IWshShell3
    Dim oShortcut   As IWshShortcut
    
    Set oShell = CreateObject("Wscript.shell")
    Set oShortcut = oShell.CreateShortcut(s_shortcut)
    ShortCutTargetPath = oShortcut.TargetPath
    Set oShell = Nothing
    Set oShortcut = Nothing

End Property

Public Property Let ShortCutTargetPath(Optional ByVal s_shortcut As String, _
                                                ByVal s_target As String)
' ----------------------------------------------------------------------------
' Creates a shortcut at a provided location with a provided path.
' ----------------------------------------------------------------------------
    
    Dim oShell      As IWshShell3
    Dim oShortcut   As IWshShortcut
    
    Set oShell = CreateObject("Wscript.shell")
    Set oShortcut = oShell.CreateShortcut(s_shortcut)
    With oShortcut
        .TargetPath = s_target
        .Save
    End With
    Set oShell = Nothing
    Set oShortcut = Nothing

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

Private Function AppErr(ByVal err_no As Long) As Long
' ----------------------------------------------------------------------------
' Used with Err.Raise AppErr(<l>).
' When the error number <l> is > 0 it is considered an "Application Error
' Number and vbObjectErrror is added to it into a negative number in order not
' to confuse with a VB runtime error. When the error number <l> is negative it
' is considered an Application Error and vbObjectError is added to convert it
' back into its origin positive number.
' ----------------------------------------------------------------------------
    If err_no < 0 Then
        AppErr = err_no - vbObjectError
    Else
        AppErr = vbObjectError + err_no
    End If
End Function

Private Function AppIsInstalled(ByVal sApp As String) As Boolean
    
    Dim i As Long: i = 1
    
    Do Until VBA.Left$(Environ$(i), 5) = "Path="
        i = i + 1
    Loop
    AppIsInstalled = InStr(Environ$(i), sApp) <> 0

End Function

Private Function ArgStrings(Optional ByVal n_v As Variant = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns the strings provided in argument (n_v) as a Collection of string
' items whereby (n_v) may be not provided, a comma delimited string, a
' Dictionary, or a Collection of string items.
' ----------------------------------------------------------------------------
    Const PROC = "ArgStrings"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim dct     As Dictionary
    Dim vString As Variant
    
    Select Case VarType(n_v)
        Case vbObject
            Select Case TypeName(n_v)
                Case "Dictionary"
                    Set dct = n_v
                    With dct
                        For Each vString In dct
                            cll.Add .Item(vString)
                        Next vString
                    End With
                    Set ArgStrings = cll
                Case "Collection"
                    Set ArgStrings = n_v
                Case Else: GoTo xt ' likely Nothing
            End Select
        Case vbString
            If n_v <> vbNullString Then
                With cll
                    For Each vString In Split(n_v, ",")
                        .Add VBA.Trim$(vString)
                    Next vString
                End With
                Set ArgStrings = cll
            End If
        Case Is >= vbArray
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a string, an arry, a Collecton, or a Dictionary!"
    End Select
            
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ArrayIsAllocated(arr As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = _
    IsArray(arr) _
    And Not IsError(LBound(arr, 1)) _
    And LBound(arr, 1) <= UBound(arr, 1)
    
End Function

Private Function ArrayNoOfDims(a_arr As Variant) As Integer
' ----------------------------------------------------------------------------
' Returns the number of dimensions of an array. An un-allocated dynamic array
' has 0 dimensions. This may as well be tested by means of ArrayIsAllocated.
' ----------------------------------------------------------------------------

    On Error Resume Next
    Dim Ndx As Integer
    Dim Res As Integer
    
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(a_arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    ArrayNoOfDims = Ndx - 1

End Function

Private Sub ArrayRemoveItems(ByRef a_va As Variant, _
                    Optional ByVal a_element As Variant, _
                    Optional ByVal a_index As Variant, _
                    Optional ByVal a_no_of_elements = 1)
' ------------------------------------------------------------------------------
' Returns the array (va) with the number of elements (a_no_of_elements) removed
' whereby the start element may be indicated by the element number 1,2,...
' (a_element) or the index (a_index) which must be within the array's LBound to
' Ubound. Any inapropriate provision of arguments results in a clear error
' message. When the last item in an array is removed the returned array is
' erased (no longer allocated).
'
' Restriction: Works only with one dimensional arrays.
'
' W. Rauschenberger, Berlin Jan 2020
' ------------------------------------------------------------------------------
    Const PROC = "ArrayRemoveItems"

    On Error GoTo eh
    Dim a                   As Variant
    Dim iElement            As Long
    Dim iIndex              As Long
    Dim a_no_of_elementsInArray As Long
    Dim i                   As Long
    Dim iNewUBound          As Long
    
    If Not IsArray(a_va) Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Array not provided!"
    Else
        a = a_va
        a_no_of_elementsInArray = UBound(a) - LBound(a) + 1
    End If
    If Not ArrayNoOfDims(a) = 1 Then
        Err.Raise AppErr(2), ErrSrc(PROC), "Array must not be multidimensional!"
    End If
    If Not IsNumeric(a_element) And Not IsNumeric(a_index) Then
        Err.Raise AppErr(3), ErrSrc(PROC), "Neither FromElement nor FromIndex is a numeric value!"
    End If
    If IsNumeric(a_element) Then
        iElement = a_element
        If iElement < 1 _
        Or iElement > a_no_of_elementsInArray Then
            Err.Raise AppErr(4), ErrSrc(PROC), "vFromElement is not between 1 and " & a_no_of_elementsInArray & " !"
        Else
            iIndex = LBound(a) + iElement - 1
        End If
    End If
    If IsNumeric(a_index) Then
        iIndex = a_index
        If iIndex < LBound(a) _
        Or iIndex > UBound(a) Then
            Err.Raise AppErr(5), ErrSrc(PROC), "FromIndex is not between " & LBound(a) & " and " & UBound(a) & " !"
        Else
            iElement = ElementOfIndex(a, iIndex)
        End If
    End If
    If iElement + a_no_of_elements - 1 > a_no_of_elementsInArray Then
        Err.Raise AppErr(6), ErrSrc(PROC), "FromElement (" & iElement & ") plus the number of elements to remove (" & a_no_of_elements & ") is beyond the number of elelemnts in the array (" & a_no_of_elementsInArray & ")!"
    End If
    
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

Private Sub ArrayTrimm(ByRef a As Variant)
' ------------------------------------------------------------------------------
' Returns the array (a) with all leading and trailing blank items removed. Any
' vbCr, vbCrLf, vbLf are ignored. When the array contains only blank items the
' returned array is erased.
' ------------------------------------------------------------------------------
    Const PROC  As String = "ArrayTrimm"

    On Error GoTo eh
    Dim i As Long
    
    '~~ Eliminate leading blank lines
    If Not ArrayIsAllocated(a) Then Exit Sub
    
    Do While (Len(Trim$(a(LBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
        ArrayRemoveItems a, a_index:=i
        If Not ArrayIsAllocated(a) Then Exit Do
    Loop
    
    If ArrayIsAllocated(a) Then
        Do While (Len(Trim$(a(UBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
            If UBound(a) = 0 Then
                Erase a
            Else
                ReDim Preserve a(UBound(a) - 1)
            End If
            If Not ArrayIsAllocated(a) Then Exit Do
        Loop
    End If

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function DiffItem(ByVal d_line As Long, _
                          ByVal d_file_left As String, _
                          ByVal d_file_right As String, _
                 Optional ByVal d_line_left As String = vbNullString, _
                 Optional ByVal d_line_right As String = vbNullString) As String
' --------------------------------------------------------------------
'
' --------------------------------------------------------------------
    Dim sFileLeft   As String
    Dim sFileRight  As String
    Dim i           As Long
    
    For i = 1 To Min(Len(d_file_left), Len(d_file_right))
        If VBA.Mid$(d_file_left, i, 1) <> VBA.Mid$(d_file_right, i, 1) _
        Then Exit For
    Next i
    i = i - 2
    sFileLeft = "..." & VBA.Right$(d_file_left, Len(d_file_left) - i) & "Line " & Format(d_line, "0000") & ": "
    sFileRight = "..." & VBA.Right$(d_file_right, Len(d_file_right) - i) & "Line " & Format(d_line, "0000") & ": "
    
    DiffItem = sFileLeft & "'" & d_line_left & "'" & vbLf & sFileRight & "'" & d_line_right & "'"

End Function

Private Function ElementOfIndex(ByVal a As Variant, _
                                ByVal i As Long) As Long
' ------------------------------------------------------
' Returns the element number of index (i) in array (a).
' ------------------------------------------------------
    Dim ia  As Long
    
    For ia = LBound(a) To i
        ElementOfIndex = ElementOfIndex + 1
    Next ia
    
End Function

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
    ErrSrc = "mFso." & sProc
End Function

Public Function Exists(Optional ByVal x_folder As String = vbNullString, _
                       Optional ByVal x_file As String = vbNullString, _
                       Optional ByRef x_result_folder As Folder = Nothing, _
                       Optional ByRef x_result_files As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Universal File System Objects existence check whereby the existence check
' depend on the provided arguments. The function returns TRUE when:
'
' Argument     | TRUE condition (despite the fact not vbNullString)
' -------------| --------------------------------------------------------
' x_folder     | The folder exists, no x_file provided
' x_file       | When no x_folder had been provided the provided x_file
'              | exists. When an x_folder had been provided at least one
'              | or more x_file meet the LIKE criteria and x_section is
'              | not provided
' x_section    | Exactly one file had been passed the existenc check, the
'              | provided section exists and no x_value_name is provided.
' x_value_name | The provided value-name exists - in the existing section
'              | in the one and only existing file.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Exists"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim fo          As Folder   ' Folder
    Dim sfo         As Folder   ' Sub-Folder
    Dim fl          As File
    Dim queue       As New Collection
    
    Set x_result_files = New Collection

    With FSo
        If Not x_folder = vbNullString Then
            '~~ Folder existence check
            If Not .FolderExists(x_folder) Then GoTo xt
            Set x_result_folder = .GetFolder(x_folder)
            If x_file = vbNullString Then
                '~~ When no x_file is provided, that's it
                Exists = True
                GoTo xt
            End If
        End If
        
        If x_file <> vbNullString And x_folder <> vbNullString Then
            '~~ For the existing folder an x_file argument had been provided
            '~~ This is interpreted as a "Like" existence check is due which
            '~~ by default includes all subfolders
            sFileName = x_file
            Set fo = .GetFolder(x_folder)
            Set queue = New Collection
            queue.Add fo

            Do While queue.Count > 0
                Set fo = queue(queue.Count)
                queue.Remove queue.Count ' dequeue the processed subfolder
                For Each sfo In fo.SubFolders
                    queue.Add sfo ' enqueue (collect) all subfolders
                Next sfo
                For Each fl In fo.Files
                    If VBA.Left$(fl.Name, 1) <> "~" _
                    And fl.Name Like x_file Then
                        '~~ The file in the (sub-)folder meets the search criteria
                        '~~ In case the x_file does not contain any "LIKE"-wise characters
                        '~~ only one file may meet the criteria
                        x_result_files.Add fl
                        Exists = True
                     End If
                Next fl
            Loop
            If x_result_files.Count <> 1 Then
                '~~ None of the files in any (sub-)folder matched with x_file
                '~~ or more than one file matched
                GoTo xt
            End If
        ElseIf x_file <> vbNullString And x_folder = vbNullString Then
            If Not .FileExists(x_file) Then GoTo xt
            x_result_files.Add .GetFile(x_file)
            Exists = True
        End If
        
    End With
        
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileCompareByFc(ByVal f_file1 As String, f_file2 As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "FileCompareByFc"
    
    On Error GoTo eh
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String
    Dim wshShell        As Object
    
    If Not FSo.FileExists(f_file1) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & f_file1 & """ does not exist!"
    
    If Not FSo.FileExists(f_file2) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & f_file2 & """ does not exist!"
    
    sCommand = "Fc /C /W " & _
               """" & f_file1 & """" & " " & _
               """" & f_file2 & """"
    
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        FileCompareByFc = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FileCompareByWinMerge(ByVal f_file_left As String, _
                                      ByVal f_lef_title As String, _
                                      ByVal f_file_right As String, _
                                      ByVal f_right_title As String) As Long
' ----------------------------------------------------------------------------
' Compares two text files by means of WinMerge. An error is raised when
' WinMerge is not installed of one of the two files doesn't exist.
' ----------------------------------------------------------------------------
    Const PROC = "FileCompareByWinMerge"
    
    On Error GoTo eh
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String
    Dim wshShell        As Object
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1), Source:=ErrSrc(PROC), Description:= _
                   "WinMerge is obligatory for the Compare service of this module " & _
                   "but not installed!" & vbLf & vbLf & _
                   "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not FSo.FileExists(f_file_left) _
    Then Err.Raise Number:=AppErr(2), Source:=ErrSrc(PROC), Description:= _
                   "The file """ & f_file_left & """ does not exist!"
    
    If Not FSo.FileExists(f_file_right) _
    Then Err.Raise Number:=AppErr(3), Source:=ErrSrc(PROC), Description:= _
                   "The file """ & f_file_right & """ does not exist!"
    
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & f_lef_title & DQUOTE & _
               " /dr " & DQUOTE & f_right_title & DQUOTE & " " & _
               """" & f_file_left & """" & " " & _
               """" & f_file_right & """"
    
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        FileCompareByWinMerge = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub FileDelete(ByVal v As Variant)

    With FSo
        If TypeName(v) = "File" Then
            .DeleteFile v.Path
        ElseIf TypeName(v) = "String" Then
            If .FileExists(v) Then .DeleteFile v
        End If
    End With
    
End Sub

Public Function FileDiffersFromFile(ByVal f_file_this As String, _
                                    ByVal f_file_from As String, _
                           Optional ByVal f_exclude_empty As Boolean = False, _
                           Optional ByVal f_compare As VbCompareMethod = vbTextCompare) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when this code lines differ from those (d_from_code) code lines
' ----------------------------------------------------------------------------
    FileDiffersFromFile = StrComp(FileString(f_file_this), FileString(f_file_from), f_compare)
End Function

Public Sub FileDiffersFromFileDebug(ByVal f_file_this As String, _
                                    ByVal f_file_from As String, _
                           Optional ByVal f_exclude_empty As Boolean = False)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "FileDiffersFromFileDebug"
    
    Dim i       As Long
    Dim arr     As Variant
    Dim arrFrom As Variant
    
    If f_exclude_empty _
    Then Debug.Print ErrSrc(PROC) & ": " & "changed (empty excluded)   : " & f_file_this _
    Else Debug.Print ErrSrc(PROC) & ": " & "changed (empty included)   : " & f_file_this
    arr = Split(StringTrimmed(f_exclude_empty), vbCrLf)
    arrFrom = Split(StringTrimmed(f_exclude_empty), vbCrLf)
    For i = 1 To Min(UBound(arr), UBound(arrFrom))
        If StrComp(arr(i), arrFrom(i), vbTextCompare) <> 0 Then
            Debug.Print ErrSrc(PROC) & ": " & "                             The first difference has been detected in line " & i & ":"
            Debug.Print ErrSrc(PROC) & ": " & "                             Line " & i & " """ & arr(i) & """"
            Debug.Print ErrSrc(PROC) & ": " & "                             Line " & i & " """ & arrFrom(i) & """"
            Exit For
        End If
    Next i

End Sub

Public Function FileExtension(ByVal fe_file As Variant)

    With FSo
        If TypeName(fe_file) = "File" Then
            FileExtension = .GetExtensionName(fe_file.Path)
        Else
            FileExtension = .GetExtensionName(fe_file)
        End If
    End With

End Function

Public Function FileIsValidName(ivf_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is a valid file name.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileOrFolderName"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim a()         As String
    Dim i           As Long
    Dim v           As Variant
    
    '~~ Check each element of the argument whether it can be created as file
    '~~ !!! this is the brute force method to check valid file names
    a = Split(ivf_name, "\")
    i = UBound(a)
    With FSo
        For Each v In a
            '~~ Check each element of the path (except the drive spec) whether it can be created as a file
            If InStr(v, ":") = 0 Then ' exclude the drive spec
                On Error Resume Next
                sFileName = .GetSpecialFolder(2) & "\" & v
                .CreateTextFile sFileName
                FileIsValidName = Err.Number = 0
                If Not FileIsValidName Then GoTo xt
                On Error GoTo eh
                If .FileExists(sFileName) Then
                    .DeleteFile sFileName
                End If
            End If
        Next v
    End With
    
xt: Exit Function
 
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FilePicked(Optional ByVal p_title As String = "Select a file", _
                           Optional ByVal p_multi As Boolean = False, _
                           Optional ByVal p_init_path As String = "C:\", _
                           Optional ByVal p_filters As String = "All files,*.*", _
                           Optional ByRef p_file As File = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Displays an msoFileDialogFilePicker dialog.
' - When a file has been selected the function returns TRUE with the selected
'   file as object (p_file).
' - When no file is selected the function returns FALSE and a file object
'   (p_file) = None.
' - All arguments have a reasonable default
' - The filters argument (p_filters) defaults to "All files", "*.*" with the
'   foillowing syntax:
'   <title>,<filter>[<title>,<filter>]...
' ----------------------------------------------------------------------------
    Const PROC = "FilePicked"
   
    On Error GoTo eh
    Dim v   As Variant
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = p_multi
        .Title = p_title
        .InitialFileName = p_init_path
        .Filters.Clear
        For Each v In Split(p_filters, ";")
            .Filters.Add Description:=Trim(Split(v, ",")(0)), Extensions:=Trim(Split(v, ",")(1))
        Next v
#If mTrc = 1 Then  ' exclude the time spent for the selection dialog execution
        mTrc.Pause      ' from the trace
#ElseIf clsTrc = 1 Then
        Trc.Pause
#End If                 ' when the execution trace is active
        If .Show = -1 Then
            FilePicked = True
            Set p_file = FSo.GetFile(.SelectedItems(1))
        Else
            Set p_file = Nothing
        End If
#If mTrc = 1 Then
        mTrc.Continue
#ElseIf clsTrc = 1 Then
        Trc.Continue
#End If
     End With
     
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FilesSearch(ByVal f_root As String, _
                   Optional ByVal f_mask As String = "*", _
                   Optional ByVal f_in_subfolders As Boolean = True, _
                   Optional ByVal f_stop_after As Long = 0) As Collection
' ---------------------------------------------------------------------
' Returns a collection of all file names which meet the criteria:
' - in any subfolder of the root (f_root)
' - meeting the wildcard comparison (f_file_mask)
' ---------------------------------------------------------------------
    Const PROC = "FilesSearch"
    
    On Error GoTo eh
    Dim fo      As Folder
    Dim sfo     As Folder
    Dim fl      As File
    Dim queue   As New Collection

    Set FilesSearch = New Collection
    If Right(f_root, 1) = "\" Then f_root = Left(f_root, Len(f_root) - 1)
    If Not FSo.FolderExists(f_root) Then GoTo xt
    queue.Add FSo.GetFolder(f_root)

    Do While queue.Count > 0
        Set fo = queue(queue.Count)
        queue.Remove queue.Count ' dequeue the processed subfolder
        If f_in_subfolders Then
            For Each sfo In fo.SubFolders
                queue.Add sfo ' enqueue (collect) all subfolders
            Next sfo
        End If
        For Each fl In fo.Files
            If VBA.Left$(fl.Name, 1) <> "~" _
            And fl.Name Like f_mask _
            Then FilesSearch.Add fl
        Next fl
    Loop

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FolderIsValidName(ivf_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is a vaslid folder name.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileOrFolderName"
    
    On Error GoTo eh
    Dim a() As String
    Dim v   As Variant
    Dim fo  As String
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "^(?!(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.[^.]*)?$)[^<>:""/\\|?*\x00-\x1F]*[^<>:""/\\|?*\x00-\x1F\ .]$"
        FolderIsValidName = Not .Test(ivf_name)
    End With
    
    If FolderIsValidName Then
        '~~ Let's prove the above by a 'brute force method':
        '~~ Checking each element of the argument whether it can be created as folder
        '~~ is the final assertion.
        a = Split(ivf_name, "\")
        With FSo
            For Each v In a
                '~~ Check each element of the path (except the drive spec) whether it can be created as a file
                If InStr(v, ":") = 0 Then ' exclude the drive spec
                    fo = .GetSpecialFolder(2) & "\" & v
                    On Error Resume Next
                    .CreateFolder fo
                    FolderIsValidName = Err.Number = 0
                    If Not FolderIsValidName Then GoTo xt
                    On Error GoTo eh
                    If .FolderExists(fo) Then
                        .DeleteFolder fo
                    End If
                End If
            Next v
        End With
    End If
    
xt: Exit Function
 
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Folders(Optional ByVal fo_spec As String = vbNullString, _
                        Optional ByVal fo_subfolders As Boolean = False, _
                        Optional ByRef fo_result As String) As Collection
' ----------------------------------------------------------------------------
' Returns all folders in a folder (fo_spec) - optionally including all
' sub-folders (fo_subfolders = True) - as folder objects in ascending order.
' When no folder (fo_spec) is provided a folder selection dialog request one.
' When the provided folder does not exist or no folder is selected the
' the function returns with an empty collection. The provided or selected
' folder is returned (fo_result).
' ----------------------------------------------------------------------------
    Const PROC = "Folders"
    
    Static cll      As Collection
    Static queue    As Collection   ' FiFo queue for folders with sub-folders
    Static Stack    As Collection   ' LiFo stack for recursive calls
    Static foStart  As Folder
    Dim aFolders()  As Variant
    Dim fo1         As Folder
    Dim fo2         As Folder
    Dim i           As Long
    Dim j           As Long
    Dim s           As String
    Dim v           As Variant
    
    If cll Is Nothing Then Set cll = New Collection
    If queue Is Nothing Then Set queue = New Collection
    If Stack Is Nothing Then Set Stack = New Collection
    
    If queue.Count = 0 Then
        '~~ Provide the folder to start with - when not provided by fo_spec via a selection dialog
        If fo_spec <> vbNullString Then
            If Not FSo.FolderExists(fo_spec) Then
                fo_result = fo_spec
                GoTo xt
            End If
            Set fo1 = FSo.GetFolder(fo_spec)
        Else
            Application.DisplayAlerts = False
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "Please select the desired folder!"
                .InitialFileName = CurDir
                .AllowMultiSelect = False
                If .Show <> -1 Then GoTo xt
                Set fo1 = FSo.GetFolder(.SelectedItems(1))
            End With
        End If
        Set foStart = fo1
    Else
        '~~ When recursively called (Queue.Count <> 0) take first sub-folder queued
        Set fo1 = queue(1)
    End If
    
    For Each fo2 In fo1.SubFolders
        cll.Add fo2
        If fo1.SubFolders.Count <> 0 And fo_subfolders Then
            queue.Add fo2
        End If
    Next fo2
    Stack.Add cll ' stack result in preparation for the function being called resursively
    
    If queue.Count > 0 Then
        Debug.Print ErrSrc(PROC) & ": " & "Remove " & queue(1).Path & " from stack (" & queue.Count - 1 & " still in queue)"
        queue.Remove 1
    End If
    If queue.Count > 0 Then
        Folders queue(1).Path ' recursive call for each folder with subfolders
    End If
    
xt: If Stack.Count > 0 Then
        Set cll = Stack(Stack.Count)
        Stack.Remove Stack.Count
    End If
    If Stack.Count = 0 Then
        If cll.Count > 0 Then
            '~~ Unload cll to array, when fo_subfolders = False only those with a ParentFolder foStart
            ReDim aFolders(cll.Count - 1)
            For Each v In cll
                aFolders(i) = v
                i = i + 1
            Next v
            
            '~~ Sort array from A to Z
            For i = LBound(aFolders) To UBound(aFolders)
                For j = i + 1 To UBound(aFolders)
                    If UCase(aFolders(i)) > UCase(aFolders(j)) Then
                        s = aFolders(j)
                        aFolders(j) = aFolders(i)
                        aFolders(i) = s
                    End If
                Next j
            Next i
            
            '~~ Transfer array as folder objects to collection
            Set cll = New Collection
            For i = LBound(aFolders) To UBound(aFolders)
                Set fo1 = FSo.GetFolder(aFolders(i))
                cll.Add fo1
            Next i
        End If
        Set Folders = cll
        If Not foStart Is Nothing Then fo_result = foStart.Path
    End If
    Set cll = Nothing
    
End Function

Private Function IsValidFileFullName(ByVal i_s As String, _
                            Optional ByRef i_fle As File) As Boolean
' ----------------------------------------------------------------------------
' Returns True when a string (i_s) is a valid file full name. If True and the file
' exists it is returned (i_fle).
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileFullName"

    On Error GoTo eh
    Dim bExists As Boolean
    Dim fle     As File

    With FSo
        If InStr(i_s, "\") = 0 Then GoTo xt ' not a valid path
        On Error Resume Next
        bExists = .FileExists(i_s)
        If Err.Number = 0 Then
            On Error GoTo eh
            If bExists Then
                IsValidFileFullName = True
                Set i_fle = .GetFile(i_s)
            Else
                On Error Resume Next
                Set fle = .CreateTextFile(i_s)
                If Err.Number = 0 Then
                    IsValidFileFullName = True
                    .DeleteFile i_s
                    Exit Function
                End If
            End If
        End If
        If IsValidFileFullName Then
            IsValidFileFullName = .GetExtensionName(i_s) <> vbNullString
        End If
    End With
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Private Function Min(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' --------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Private Function Pop(ByRef s_stk As Collection) As Variant
    If s_stk.Count > 0 Then
        If IsObject(s_stk(1)) Then Set Pop = s_stk(1) Else Pop = s_stk(1)
        s_stk.Remove 1
    End If
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Public Sub RenameSubFolders(ByVal r_folder_path As String, _
                            ByVal r_folder_old_name As String, _
                            ByVal r_folder_new_name As String, _
                   Optional ByRef r_renamed As Collection = Nothing)
' ----------------------------------------------------------------------------
' Rename any sub-folder in the provided path (r_folder_path) named (r_folder_old_name)
' to (r_folder_new_name). Any subfolders in renamed folders are ignored. Returns
' all renamed folders as Collection (r_renamed).
' ----------------------------------------------------------------------------
    Dim cll As Collection
    Dim v   As Variant
    Dim fld As Folder
    
    Set cll = SubFolders(r_folder_path)
    For Each v In cll
        Set fld = v
        If fld.Name = r_folder_old_name Then
            fld.Name = r_folder_new_name
            If Not r_renamed Is Nothing Then r_renamed.Add fld
        End If
    Next v
    
xt: Exit Sub

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

Private Function SplitString(ByVal s_s As String) As String
    
    Select Case True
        Case InStr(s_s, vbCrLf) <> 0: SplitString = vbCrLf
        Case InStr(s_s, vbCr) <> 0:   SplitString = vbCr
        Case InStr(s_s, vbLf) <> 0:   SplitString = vbLf
    End Select
    If Len(SplitString) = 0 Then SplitString = vbCrLf
    
End Function

Public Function StringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    Dim sSplit As String
    
    '~~ Obtain (and return) the used line terminating string (vbCrLf) or character
    sSplit = SplitString(s_s)
    
    s_s = StringTrimmed(s_s) ' leading and trailing empty already excluded
    
    Do While InStr(s_s, sSplit & sSplit) <> 0
        s_s = Replace(s_s, sSplit & sSplit, sSplit)
    Loop
    StringEmptyExcluded = s_s
    
End Function

Private Function StringTrimmed(ByVal s_s As String, _
                      Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) provided as a single string, with any leading and
' trailing empty items (record/lines) excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    Dim v As Variant
    
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
    
    StringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Public Function SubFolders(ByVal s_path As String) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection of all folders and sub-folders in a folder (s_path).
' ----------------------------------------------------------------------------
    Const PROC = "SubFolders"
    
    Dim stk As New Collection
    Dim cll As New Collection
    Dim fld As Folder
    Dim sfd As Folder ' Subfolder
        
    Push(stk) = FSo.GetFolder(s_path)  ' push the initial folder onto the queue
    Do While stk.Count > 0
        Set fld = Pop(stk)             ' pop the first dolder pushed to the queue
        cll.Add fld
        Debug.Print ErrSrc(PROC) & ": " & fld.Path
        For Each sfd In fld.SubFolders
            Push(stk) = sfd
        Next sfd
    Loop
    
    Set SubFolders = cll
    Set cll = Nothing
    
End Function

