Attribute VB_Name = "mFso_RnmdByCompMan"
'Option Explicit
'Option Compare Text
'' ----------------------------------------------------------------------------
'' Standard  Module mFso: Common services regarding files and folders. The
'' ====================== module makes extensive use of the FileSystemObject
'' by extending and supplementing its services. The componet works autonomous
'' and does not require any other components.
''
'' Public properties and services:
'' -------------------------------
'' Exists                Universal existence function. Returns TRUE when the
''                       folder, file, section ([xxxx]), or value name exists.
'' FileCompareByWinMerge Displays the differences of two files by using
''                       WinMerge
'' FileDiffers           Returns a Dictionary with those records/lines which
''                       differ between two provided files, with the options
''                       'ignore case' and 'ignore empty records'
'' FileDelete            Deletes a file provided either as object or as full
''                       name
'' FileExtension         Returns the extension of a file's name
'' Folders
'' FileArry          r/w Returns the content of a text file as an array or
''                       writes the content of an array to a text file.
'' FileString        r/w Returns the content of a text file as text string
''                       or writes a string to a text file - optionally
''                       appended
'' FileTemp
'' RenameSubFolders      Renames all folders and sub-folders in a provided
''                       path from a given old to a given new name.
'' SortCutTargetPath r   Returns a provided shortcut's target path
''                       Let Creates a shortcut at a provided location with a
''                       provided path
''
'' Requires References to:
'' - Microsoft Scripting Runtine
'' - Windows Script Host Object Model
''
'' Uses no other components. Optionally uses mErH, fMsg/mMsg when installed and
'' activated (Cond. Comp. Args. `mErH = 1 : mMsg = 1`).
''
'' W. Rauschenberger, Berlin Apr 2024
'' See also https://github.com/warbe-maker/VBA-File-System-Objects.
'' ----------------------------------------------------------------------------
'Public fso                      As New FileSystemObject
'Private dctFilesOpen            As New Dictionary
'Private Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-File-System-Objects"
'
'#If mMsg = 0 Then
'    ' ------------------------------------------------------------------------
'    ' The 'minimum error handling' approach implemented with this module and
'    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
'    ' message which includes a debugging option to resume the error line
'    ' provided the Conditional Compile Argument 'Debugging = 1'.
'    ' This declaration allows the mTrc module to work completely autonomous.
'    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
'    ' be indicated by the Conditional Compile Argument mMsg = 1
'    '
'    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
'    '    how to install an use.
'    ' ------------------------------------------------------------------------
'    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
'#End If
'
'
'Private Declare PtrSafe Function GetFileTime Lib "kernel32" ( _
'    ByVal hFile As LongPtr, _
'    ByRef lpCreationTime As Currency, _
'    ByRef lpLastAccessTime As Currency, _
'    ByRef lpLastWriteTime As Currency) As Long
'
'Private Declare PtrSafe Function CopyFileEx Lib "kernel32" Alias "CopyFileExA" ( _
'    ByVal lpExistingFileName As String, _
'    ByVal lpNewFileName As String, _
'    ByVal lpProgressRoutine As LongPtr, _
'    ByVal lpData As LongPtr, _
'    ByVal pbCancel As LongPtr, _
'    ByVal dwCopyFlags As Long) As Long
'
'Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
'    ByVal lpFileName As String, _
'    ByVal dwDesiredAccess As Long, _
'    ByVal dwShareMode As Long, _
'    ByVal lpSecurityAttributes As LongPtr, _
'    ByVal dwCreationDisposition As Long, _
'    ByVal dwFlagsAndAttributes As Long, _
'    ByVal hTemplateFile As LongPtr) As LongPtr
'
'Private Declare PtrSafe Function SetFileTime Lib "kernel32" ( _
'    ByVal hFile As LongPtr, _
'    ByRef lpCreationTime As Currency, _
'    ByRef lpLastAccessTime As Currency, _
'    ByRef lpLastWriteTime As Currency) As Long
'
'Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
'    ByVal hObject As LongPtr) As Long
'
'Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
'    Alias "ShellExecuteA" _
'    (ByVal hwnd As Long, _
'    ByVal lpOperation As String, _
'    ByVal lpFile As String, _
'    ByVal lpParameters As String, _
'    ByVal lpDirectory As String, _
'    ByVal nShowCmd As Long) _
'    As Long
'
''***App Window Constants***
'Private Const WIN_NORMAL = 1         'Open Normal
'
'Private Const DQUOTE     As String = """" ' one " character
''***Error Codes***
'Private Const ERROR_SUCCESS = 32&
'Private Const ERROR_NO_ASSOC = 31&
'Private Const ERROR_OUT_OF_MEM = 0&
'Private Const ERROR_FILE_NOT_FOUND = 2&
'Private Const ERROR_PATH_NOT_FOUND = 3&
'Private Const ERROR_BAD_FORMAT = 11&
'
'Public Property Get FileDict(ByVal f_file As Variant) As Dictionary
'' ----------------------------------------------------------------------------
'' Returns the content of the file (f_file) - which may be provided as file
'' object or full file name - as Dictionary by considering any kind of line
'' break characters.
'' ----------------------------------------------------------------------------
'    Const PROC  As String = "FileDict-Get"
'
'    On Error GoTo eh
'    Dim a       As Variant
'    Dim dct     As New Dictionary
'    Dim sSplit  As String
'    Dim flo     As File
'    Dim sFile   As String
'    Dim i       As Long
'
'    Select Case TypeName(f_file)
'        Case "File"
'            If Not fso.FileExists(f_file.Path) _
'            Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (f_file) does not exist!"
'            Set flo = f_file
'        Case "String"
'            If Not fso.FileExists(f_file) _
'            Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided file string (f_file) is not an existing file!"
'            Set flo = fso.GetFile(f_file)
'    End Select
'
'    '~~ Unload file into a test stream
'    sFile = FileAsString(flo.Path)
'    If sFile = vbNullString Then GoTo xt
'    sSplit = SplitString(sFile)
'
'    '~~ Test stream to array
'    a = Split(sFile, sSplit)
'
'    '~~ Remove any leading or trailing empty items
'    ArrayTrimm a
'
'    For i = LBound(a) To UBound(a)
'        dct.Add i + 1, a(i)
'    Next i
'
'xt: Set FileDict = dct
'    Exit Property
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Property
'
'Public Property Get FileTemp(Optional ByVal f_path As String = vbNullString, _
'                             Optional ByVal f_extension As String = ".tmp") As String
'' ------------------------------------------------------------------------------
'' Returns the full file name of a temporary randomly named file. When a path
'' (f_path) is omitted in the CurDir path, else in at the provided folder.
'' ------------------------------------------------------------------------------
'    Dim sTemp As String
'
'    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
'    sTemp = Replace(fso.GetTempName, ".tmp", f_extension)
'    If f_path = vbNullString Then f_path = CurDir
'    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
'    FileTemp = sTemp
'
'End Property
'
'Private Property Let Push(Optional ByRef s_stk As Collection, _
'                                   ByVal v As Variant)
'    s_stk.Add v
'End Property
'
'Public Property Get ShortCutTargetPath(Optional ByVal s_shortcut As String) As String
'' ----------------------------------------------------------------------------
'' Returns a provided shortcut's target path.
'' ----------------------------------------------------------------------------
'    Dim oShell      As IWshShell3
'    Dim oShortcut   As IWshShortcut
'
'    Set oShell = CreateObject("Wscript.shell")
'    Set oShortcut = oShell.CreateShortcut(s_shortcut)
'    ShortCutTargetPath = oShortcut.TargetPath
'    Set oShell = Nothing
'    Set oShortcut = Nothing
'
'End Property
'
'Public Property Let ShortCutTargetPath(Optional ByVal s_shortcut As String, _
'                                                ByVal s_target As String)
'' ----------------------------------------------------------------------------
'' Creates a shortcut at a provided location with a provided path.
'' ----------------------------------------------------------------------------
'
'    Dim oShell      As IWshShell3
'    Dim oShortcut   As IWshShortcut
'
'    Set oShell = CreateObject("Wscript.shell")
'    Set oShortcut = oShell.CreateShortcut(s_shortcut)
'    With oShortcut
'        .TargetPath = s_target
'        .Save
'    End With
'    Set oShell = Nothing
'    Set oShortcut = Nothing
'
'End Property
'
'Private Property Get SplitStr(ByRef s As String)
'' ------------------------------------------------------------------------------
'' Returns the split string in string (s) used by VBA.Split() to turn the string
'' into an array.
'' ------------------------------------------------------------------------------
'    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
'    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
'    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
'End Property
'
'Private Function AppErr(ByVal err_no As Long) As Long
'' ----------------------------------------------------------------------------
'' Used with Err.Raise AppErr(<l>).
'' When the error number <l> is > 0 it is considered an "Application Error
'' Number and vbObjectErrror is added to it into a negative number in order not
'' to confuse with a VB runtime error. When the error number <l> is negative it
'' is considered an Application Error and vbObjectError is added to convert it
'' back into its origin positive number.
'' ----------------------------------------------------------------------------
'    If err_no < 0 Then
'        AppErr = err_no - vbObjectError
'    Else
'        AppErr = vbObjectError + err_no
'    End If
'End Function
'
'Private Function AppIsInstalled(ByVal sApp As String) As Boolean
'
'    Dim i As Long: i = 1
'
'    Do Until VBA.Left$(Environ$(i), 5) = "Path="
'        i = i + 1
'    Loop
'    AppIsInstalled = InStr(Environ$(i), sApp) <> 0
'
'End Function
'
'Private Function ArgStrings(Optional ByVal n_v As Variant = Nothing) As Collection
'' ----------------------------------------------------------------------------
'' Returns the strings provided in argument (n_v) as a Collection of string
'' items whereby (n_v) may be not provided, a comma delimited string, a
'' Dictionary, or a Collection of string items.
'' ----------------------------------------------------------------------------
'    Const PROC = "ArgStrings"
'
'    On Error GoTo eh
'    Dim cll     As New Collection
'    Dim dct     As Dictionary
'    Dim vString As Variant
'
'    Select Case VarType(n_v)
'        Case vbObject
'            Select Case TypeName(n_v)
'                Case "Dictionary"
'                    Set dct = n_v
'                    With dct
'                        For Each vString In dct
'                            cll.Add .Item(vString)
'                        Next vString
'                    End With
'                    Set ArgStrings = cll
'                Case "Collection"
'                    Set ArgStrings = n_v
'                Case Else: GoTo xt ' likely Nothing
'            End Select
'        Case vbString
'            If n_v <> vbNullString Then
'                With cll
'                    For Each vString In Split(n_v, ",")
'                        .Add VBA.Trim$(vString)
'                    Next vString
'                End With
'                Set ArgStrings = cll
'            End If
'        Case Is >= vbArray
'        Case Else
'            Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a string, an arry, a Collecton, or a Dictionary!"
'    End Select
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function ArrayAsFile(ByVal a_arry As Variant, _
'                   Optional ByRef a_file As Variant = vbNullString, _
'                   Optional ByVal a_file_append As Boolean = False) As File
'' ----------------------------------------------------------------------------
'' Writes all items of an array (a_arry) to a file (a_file) which might be a
'' file object, a file's full name. When no file (a_file) is provided a
'' temporary file is returned, else the provided file (a_file) as object.
'' Note: In oreder to make sure the array (a_arry) can be joined, any item
''       which returns an error is set to Empty.
'' ----------------------------------------------------------------------------
'    Const PROC = "ArrayAsFile"
'
'    On Error GoTo xt
'    Dim i       As Long
'    Dim sFile   As String
'    Dim s       As String
'
'    If UBound(a_arry) >= LBound(a_arry) Then
'        Select Case True
'            Case a_file = vbNullString:         sFile = TempFileFullName
'            Case TypeName(a_file) = "File":     sFile = a_file.Path
'            Case TypeName(a_file) = "String":   sFile = a_file
'        End Select
'        '~~ Make sure all items do exist by providing non existing items with Empty
'        For i = LBound(a_arry) To UBound(a_arry)
'            If IsError(a_arry(i)) Then
'                a_arry(i) = Empty
'            End If
'        Next i
'        s = Join(a_arry, vbCrLf)
'        FileOutput sFile, s, a_file_append
'        Set ArrayAsFile = fso.GetFile(sFile)
'    End If
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Private Function ArrayIsAllocated(arr As Variant) As Boolean
'
'    On Error Resume Next
'    ArrayIsAllocated = _
'    IsArray(arr) _
'    And Not IsError(LBound(arr, 1)) _
'    And LBound(arr, 1) <= UBound(arr, 1)
'
'End Function
'
'Private Function ArrayNoOfDims(a_arr As Variant) As Integer
'' ----------------------------------------------------------------------------
'' Returns the number of dimensions of an array. An un-allocated dynamic array
'' has 0 dimensions. This may as well be tested by means of ArrayIsAllocated.
'' ----------------------------------------------------------------------------
'
'    On Error Resume Next
'    Dim Ndx As Integer
'    Dim Res As Integer
'
'    ' Loop, increasing the dimension index Ndx, until an error occurs.
'    ' An error will occur when Ndx exceeds the number of dimension
'    ' in the array. Return Ndx - 1.
'    Do
'        Ndx = Ndx + 1
'        Res = UBound(a_arr, Ndx)
'    Loop Until Err.Number <> 0
'    Err.Clear
'    ArrayNoOfDims = Ndx - 1
'
'End Function
'
'Private Sub ArrayRemoveItems(ByRef a_va As Variant, _
'                    Optional ByVal a_element As Variant, _
'                    Optional ByVal a_index As Variant, _
'                    Optional ByVal a_no_of_elements = 1)
'' ------------------------------------------------------------------------------
'' Returns the array (va) with the number of elements (a_no_of_elements) removed
'' whereby the start element may be indicated by the element number 1,2,...
'' (a_element) or the index (a_index) which must be within the array's LBound to
'' Ubound. Any inapropriate provision of arguments results in a clear error
'' message. When the last item in an array is removed the returned array is
'' erased (no longer allocated).
''
'' Restriction: Works only with one dimensional arrays.
''
'' W. Rauschenberger, Berlin Jan 2020
'' ------------------------------------------------------------------------------
'    Const PROC = "ArrayRemoveItems"
'
'    On Error GoTo eh
'    Dim a                   As Variant
'    Dim iElement            As Long
'    Dim iIndex              As Long
'    Dim a_no_of_elementsInArray As Long
'    Dim i                   As Long
'    Dim iNewUBound          As Long
'
'    If Not IsArray(a_va) Then
'        Err.Raise AppErr(1), ErrSrc(PROC), "Array not provided!"
'    Else
'        a = a_va
'        a_no_of_elementsInArray = UBound(a) - LBound(a) + 1
'    End If
'    If Not ArrayNoOfDims(a) = 1 Then
'        Err.Raise AppErr(2), ErrSrc(PROC), "Array must not be multidimensional!"
'    End If
'    If Not IsNumeric(a_element) And Not IsNumeric(a_index) Then
'        Err.Raise AppErr(3), ErrSrc(PROC), "Neither FromElement nor FromIndex is a numeric value!"
'    End If
'    If IsNumeric(a_element) Then
'        iElement = a_element
'        If iElement < 1 _
'        Or iElement > a_no_of_elementsInArray Then
'            Err.Raise AppErr(4), ErrSrc(PROC), "vFromElement is not between 1 and " & a_no_of_elementsInArray & " !"
'        Else
'            iIndex = LBound(a) + iElement - 1
'        End If
'    End If
'    If IsNumeric(a_index) Then
'        iIndex = a_index
'        If iIndex < LBound(a) _
'        Or iIndex > UBound(a) Then
'            Err.Raise AppErr(5), ErrSrc(PROC), "FromIndex is not between " & LBound(a) & " and " & UBound(a) & " !"
'        Else
'            iElement = ElementOfIndex(a, iIndex)
'        End If
'    End If
'    If iElement + a_no_of_elements - 1 > a_no_of_elementsInArray Then
'        Err.Raise AppErr(6), ErrSrc(PROC), "FromElement (" & iElement & ") plus the number of elements to remove (" & a_no_of_elements & ") is beyond the number of elelemnts in the array (" & a_no_of_elementsInArray & ")!"
'    End If
'
'    For i = iIndex + a_no_of_elements To UBound(a)
'        a(i - a_no_of_elements) = a(i)
'    Next i
'
'    iNewUBound = UBound(a) - a_no_of_elements
'    If iNewUBound < 0 Then Erase a Else ReDim Preserve a(LBound(a) To iNewUBound)
'    a_va = a
'
'xt: Exit Sub
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbYes: Stop: Resume
'        Case Else:  GoTo xt
'    End Select
'End Sub
'
'Private Sub ArrayTrimm(ByRef a As Variant)
'' ------------------------------------------------------------------------------
'' Returns the array (a) with all leading and trailing blank items removed. Any
'' vbCr, vbCrLf, vbLf are ignored. When the array contains only blank items the
'' returned array is erased.
'' ------------------------------------------------------------------------------
'    Const PROC  As String = "ArrayTrimm"
'
'    Dim i As Long
'
'    On Error GoTo xt
'    If Not UBound(a) >= LBound(a) Then
'        '~~ Eliminate leading blank lines
'        Do While (Len(Trim$(a(LBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
'            ArrayRemoveItems a, a_index:=i
'            If Not ArrayIsAllocated(a) Then Exit Do
'        Loop
'
'        If ArrayIsAllocated(a) Then
'            Do While (Len(Trim$(a(UBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
'                If UBound(a) = 0 Then
'                    Erase a
'                Else
'                    ReDim Preserve a(UBound(a) - 1)
'                End If
'                If Not ArrayIsAllocated(a) Then Exit Do
'            Loop
'        End If
'    End If
'
'xt: Exit Sub
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub
'
'Private Function ElementOfIndex(ByVal a As Variant, _
'                                ByVal i As Long) As Long
'' ------------------------------------------------------
'' Returns the element number of index (i) in array (a).
'' ------------------------------------------------------
'    Dim ia  As Long
'
'    For ia = LBound(a) To i
'        ElementOfIndex = ElementOfIndex + 1
'    Next ia
'
'End Function
'
'Private Function ErrMsg(ByVal err_source As String, _
'               Optional ByVal err_no As Long = 0, _
'               Optional ByVal err_dscrptn As String = vbNullString, _
'               Optional ByVal err_line As Long = 0) As Variant
'' ------------------------------------------------------------------------------
'' Universal error message display service which displays:
'' - a debugging option button
'' - an "About:" section when the err_dscrptn has an additional string
''   concatenated by two vertical bars (||)
'' - the error message either by means of the Common VBA Message Service
''   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
''   means of the VBA.MsgBox in case not.
''
'' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
''               to turn them into a negative and in the error message back into
''               its origin positive number.
''
'' W. Rauschenberger Berlin, Jan 2024
'' See: https://github.com/warbe-maker/VBA-Error
'' ------------------------------------------------------------------------------
'#If mErH = 1 Then
'    '~~ ------------------------------------------------------------------------
'    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
'    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
'    '~~ is preferred since it provides some enhanced features like a path to the
'    '~~ error.
'    '~~ ------------------------------------------------------------------------
'    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
'    GoTo xt
'#ElseIf mMsg = 1 Then
'    '~~ ------------------------------------------------------------------------
'    '~~ When only the Common Message Services Component (mMsg) is installed but
'    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
'    '~~ provides an enhanced layout and other features.
'    '~~ ------------------------------------------------------------------------
'    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
'    GoTo xt
'#End If
'    '~~ -------------------------------------------------------------------
'    '~~ When neither the mMsg nor the mErH component is installed the error
'    '~~ message is displayed by means of the VBA.MsgBox
'    '~~ -------------------------------------------------------------------
'    Dim ErrBttns    As Variant
'    Dim ErrAtLine   As String
'    Dim ErrDesc     As String
'    Dim ErrLine     As Long
'    Dim ErrNo       As Long
'    Dim ErrSrc      As String
'    Dim ErrText     As String
'    Dim ErrTitle    As String
'    Dim ErrType     As String
'    Dim ErrAbout    As String
'
'    '~~ Obtain error information from the Err object for any argument not provided
'    If err_no = 0 Then err_no = Err.Number
'    If err_line = 0 Then ErrLine = Erl
'    If err_source = vbNullString Then err_source = Err.Source
'    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
'    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
'
'    If InStr(err_dscrptn, "||") <> 0 Then
'        ErrDesc = Split(err_dscrptn, "||")(0)
'        ErrAbout = Split(err_dscrptn, "||")(1)
'    Else
'        ErrDesc = err_dscrptn
'    End If
'
'    '~~ Determine the type of error
'    Select Case err_no
'        Case Is < 0
'            ErrNo = AppErr(err_no)
'            ErrType = "Application Error "
'        Case Else
'            ErrNo = err_no
'            If (InStr(1, err_dscrptn, "DAO") <> 0 _
'            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
'            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
'            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
'            Then ErrType = "Database Error " _
'            Else: ErrType = "VB Runtime Error "
'    End Select
'
'    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
'    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
'    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
'
'    ErrText = "Error: " & vbLf & _
'              ErrDesc & vbLf & vbLf & _
'              "Source: " & vbLf & _
'              err_source & ErrAtLine
'    If ErrAbout <> vbNullString _
'    Then ErrText = ErrText & vbLf & vbLf & _
'                  "About: " & vbLf & _
'                  ErrAbout
'
'    ErrBttns = vbYesNo
'    ErrText = ErrText & vbLf & vbLf & _
'              "Debugging:" & vbLf & _
'              "Yes    = Resume Error Line" & vbLf & _
'              "No     = Terminate"
'    ErrMsg = MsgBox(Title:=ErrTitle _
'                  , Prompt:=ErrText _
'                  , Buttons:=ErrBttns)
'xt: Exit Function
'
'End Function
'
'Private Function ErrSrc(ByVal sProc As String) As String
'    ErrSrc = "mFso." & sProc
'End Function
'
'Public Function Exists(Optional ByVal e_folder As String = vbNullString, _
'                       Optional ByVal e_file As String = vbNullString, _
'                       Optional ByRef e_result_folder As Folder = Nothing, _
'                       Optional ByRef e_result_files As Collection = Nothing) As Boolean
'' ----------------------------------------------------------------------------
'' Universal File System Objects existence check whereby the existence check
'' depends on the provided arguments. The function returns TRUE when:
''
'' Argument     | TRUE condition (despite the fact not vbNullString)
'' -------------| --------------------------------------------------------
'' e_folder     | The folder exists, no e_file provided
'' e_file       | When no e_folder had been provided the provided e_file
''              | exists. When an e_folder had been provided at least one
''              | or more e_file meet the LIKE criteria
'' ----------------------------------------------------------------------------
'    Const PROC  As String = "Exists"
'
'    On Error GoTo eh
'    Dim sFileName   As String
'    Dim fo          As Folder   ' Folder
'    Dim sfo         As Folder   ' Sub-Folder
'    Dim fl          As File
'    Dim queue       As New Collection
'
'    Set e_result_files = New Collection
'
'    With fso
'        Select Case True
'
'            Case e_folder <> vbNullString And e_file = vbNullString
'                '~~ Folder existence check
'                Exists = .FolderExists(e_folder)
'
'            Case e_file <> vbNullString And e_folder = vbNullString
'                If .FileExists(e_file) Then
'                    e_result_files.Add .GetFile(e_file)
'                    Exists = True
'                End If
'
'            Case e_file <> vbNullString And e_folder <> vbNullString
'                '~~ For the existing folder an e_file argument had been provided
'                '~~ This is interpreted as a "Like" existence check is due which
'                '~~ by default includes all subfolders
'                sFileName = e_file
'                Set fo = .GetFolder(e_folder)
'                Set queue = New Collection
'                queue.Add fo
'
'                Do While queue.Count > 0
'                    Set fo = queue(queue.Count)
'                    queue.Remove queue.Count ' dequeue the processed subfolder
'                    For Each sfo In fo.SubFolders
'                        queue.Add sfo ' enqueue (collect) all subfolders
'                    Next sfo
'                    For Each fl In fo.Files
'                        If VBA.Left$(fl.Name, 1) <> "~" _
'                        And fl.Name Like e_file Then
'                            '~~ The file in the (sub-)folder meets the search criteria
'                            '~~ In case the e_file does not contain any "LIKE"-wise characters
'                            '~~ only one file may meet the criteria
'                            e_result_files.Add fl
'                            Exists = True
'                         End If
'                    Next fl
'                Loop
'                If e_result_files.Count <> 1 Then
'                    '~~ None of the files in any (sub-)folder matched with e_file
'                    '~~ or more than one file matched
'                    GoTo xt
'                End If
'
'        End Select
'    End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function FileAsArray(ByVal f_file As Variant, _
'                   Optional ByVal f_empty_excluded = False, _
'                   Optional ByVal f_trim As Variant = False) As Variant
'' ----------------------------------------------------------------------------
'' Returns a file's (f_file) records/lines as array.
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Dim sFile   As String
'    Dim sSplit  As String
'    Dim s       As String
'
'    Select Case True
'        Case TypeName(f_file) = "String":   sFile = f_file
'        Case TypeName(f_file) = "File":     sFile = f_file.Path
'    End Select
'    s = FileAsString(sFile, sSplit, f_empty_excluded)
'    FileAsArray = StringAsArray(s, sSplit, f_trim)
'
'End Function
'
'Public Function FileAsString(ByVal f_file As Variant, _
'                    Optional ByRef f_split As String = vbCrLf, _
'                    Optional ByVal f_exclude_empty As Boolean = False) As String
'' ----------------------------------------------------------------------------
'' Returns a file's (f_file) - provided as full name or object - records/lines
'' as a single string with the records/lines delimited (f_split).
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Const PROC = "FileAsString"
'
'    On Error GoTo eh
'    Dim s       As String
'    Dim sFile   As String
'
'    Select Case True
'        Case TypeName(f_file) = "File":     sFile = f_file.Path
'        Case TypeName(f_file) = "String":   sFile = f_file
'    End Select
'    '~~ An error is passed on to the caller
'    If Not fso.FileExists(sFile) Then Err.Raise AppErr(1), ErrSrc(PROC), _
'                                      "The file, provided by a full name, does not exist!"
'
'    FileInput sFile, s
'
'    f_split = SplitIndctr(s) ' may be vbCrLf or vbLf when file is a downloaded
'
'    '~~ Eliminate any trailing split string
'    Do While Right(s, Len(f_split)) = f_split
'        s = Left(s, Len(s) - Len(f_split))
'        If Len(s) <= Len(f_split) Then Exit Do
'    Loop
'
'    If f_exclude_empty Then
'        s = FileAsStringEmptyExcluded(s)
'    End If
'    FileAsString = s
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Private Function FileAsStringEmptyExcluded(ByVal f_s As String) As String
'' ----------------------------------------------------------------------------
'' Returns a string (f_s) with any empty elements excluded. I.e. the string
'' returned begins and ends with a non vbNullString character and has no
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'
'    Do While InStr(f_s, vbCrLf & vbCrLf) <> 0
'        f_s = Replace(f_s, vbCrLf & vbCrLf, vbCrLf)
'    Loop
'    FileAsStringEmptyExcluded = f_s
'
'End Function
'
'Public Sub FileClone(sourcePath As String, destinationPath As String)
'    Dim hFile As LongPtr
'    Dim creationTime As Currency
'    Dim lastAccessTime As Currency
'    Dim lastWriteTime As Currency
'
'    ' Copy the file
'    If CopyFileEx(sourcePath, destinationPath, 0, 0, 0, 0) = 0 Then
'        MsgBox "Error copying file.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Open the source file to get its attributes
'    hFile = CreateFile(sourcePath, &H80000000, 0, 0, 3, &H80, 0)
'    If hFile = -1 Then
'        MsgBox "Error opening source file.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Get the file times
'    GetFileTime hFile, creationTime, lastAccessTime, lastWriteTime
'    CloseHandle hFile
'
'    ' Open the destination file to set its attributes
'    hFile = CreateFile(destinationPath, &H40000000, 0, 0, 3, &H80, 0)
'    If hFile = -1 Then
'        MsgBox "Error opening destination file.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Set the file times
'    If SetFileTime(hFile, creationTime, lastAccessTime, lastWriteTime) = 0 Then
'        MsgBox "Error setting file attributes.", vbExclamation
'        CloseHandle hFile
'        Exit Sub
'    End If
'
'    ' Close the file handle
'    CloseHandle hFile
'
'    MsgBox "File successfully cloned with attributes."
'End Sub
'
'Private Sub FileClose(ByVal f_file As String)
'
'    If dctFilesOpen.Exists(f_file) Then
'        Close dctFilesOpen(f_file)
'        dctFilesOpen.Remove f_file
'    End If
'
'End Sub
'
'Private Function FileCompareByFc(ByVal f_file1 As String, f_file2 As String)
'' ----------------------------------------------------------------------------
''
'' ----------------------------------------------------------------------------
'    Const PROC = "FileCompareByFc"
'
'    On Error GoTo eh
'    Dim waitOnReturn    As Boolean: waitOnReturn = True
'    Dim windowStyle     As Integer: windowStyle = 1
'    Dim sCommand        As String
'    Dim wshShell        As Object
'
'    If Not fso.FileExists(f_file1) _
'    Then Err.Raise Number:=AppErr(2) _
'                 , Source:=ErrSrc(PROC) _
'                 , Description:="The file """ & f_file1 & """ does not exist!"
'
'    If Not fso.FileExists(f_file2) _
'    Then Err.Raise Number:=AppErr(3) _
'                 , Source:=ErrSrc(PROC) _
'                 , Description:="The file """ & f_file2 & """ does not exist!"
'
'    sCommand = "Fc /C /W " & _
'               """" & f_file1 & """" & " " & _
'               """" & f_file2 & """"
'
'    Set wshShell = CreateObject("WScript.Shell")
'    With wshShell
'        FileCompareByFc = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
'    End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function FileCompareByWinMerge(ByVal f_file_left As String, _
'                                      ByVal f_lef_title As String, _
'                                      ByVal f_file_right As String, _
'                                      ByVal f_right_title As String) As Long
'' ----------------------------------------------------------------------------
'' Compares two text files by means of WinMerge. An error is raised when
'' WinMerge is not installed of one of the two files doesn't exist.
'' ----------------------------------------------------------------------------
'    Const PROC = "FileCompareByWinMerge"
'
'    On Error GoTo eh
'    Dim waitOnReturn    As Boolean: waitOnReturn = True
'    Dim windowStyle     As Integer: windowStyle = 1
'    Dim sCommand        As String
'    Dim wshShell        As Object
'
'    If Not AppIsInstalled("WinMerge") _
'    Then Err.Raise Number:=AppErr(1), Source:=ErrSrc(PROC), Description:= _
'                   "WinMerge is obligatory for the Compare service of this module " & _
'                   "but not installed!" & vbLf & vbLf & _
'                   "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
'
'    If Not fso.FileExists(f_file_left) _
'    Then Err.Raise Number:=AppErr(2), Source:=ErrSrc(PROC), Description:= _
'                   "The file """ & f_file_left & """ does not exist!"
'
'    If Not fso.FileExists(f_file_right) _
'    Then Err.Raise Number:=AppErr(3), Source:=ErrSrc(PROC), Description:= _
'                   "The file """ & f_file_right & """ does not exist!"
'
'    sCommand = "WinMergeU /e" & _
'               " /dl " & DQUOTE & f_lef_title & DQUOTE & _
'               " /dr " & DQUOTE & f_right_title & DQUOTE & " " & _
'               """" & f_file_left & """" & " " & _
'               """" & f_file_right & """"
'
'    Set wshShell = CreateObject("WScript.Shell")
'    With wshShell
'        FileCompareByWinMerge = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
'    End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Sub FileDelete(ByVal v As Variant)
'
'    With fso
'        If TypeName(v) = "File" Then
'            .DeleteFile v.Path
'        ElseIf TypeName(v) = "String" Then
'            If .FileExists(v) Then .DeleteFile v
'        End If
'    End With
'
'End Sub
'
'Public Function FileDiffersFromFile(ByVal f_file_this As String, _
'                                    ByVal f_file_from As String, _
'                           Optional ByVal f_exclude_empty As Boolean = False, _
'                           Optional ByVal f_compare As VbCompareMethod = vbTextCompare) As Boolean
'' ----------------------------------------------------------------------------
'' Returns TRUE when this code lines differ from those (d_from_code) code lines
'' ----------------------------------------------------------------------------
'    FileDiffersFromFile = StrComp(FileAsString(f_file_this, , f_exclude_empty), FileAsString(f_file_from, , f_exclude_empty), f_compare)
'End Function
'
'Public Function FileExtension(ByVal fe_file As Variant)
'
'    With fso
'        If TypeName(fe_file) = "File" Then
'            FileExtension = .GetExtensionName(fe_file.Path)
'        Else
'            FileExtension = .GetExtensionName(fe_file)
'        End If
'    End With
'
'End Function
'
'Public Function FileIsValidName(ivf_name As String) As Boolean
'' ----------------------------------------------------------------------------
'' Returns TRUE when the provided argument is a valid file name.
'' ----------------------------------------------------------------------------
'    Const PROC = "IsValidFileOrFolderName"
'
'    On Error GoTo eh
'    Dim sFileName   As String
'    Dim a()         As String
'    Dim i           As Long
'    Dim v           As Variant
'
'    '~~ Check each element of the argument whether it can be created as file
'    '~~ !!! this is the brute force method to check valid file names
'    a = Split(ivf_name, "\")
'    i = UBound(a)
'    With fso
'        For Each v In a
'            '~~ Check each element of the path (except the drive spec) whether it can be created as a file
'            If InStr(v, ":") = 0 Then ' exclude the drive spec
'                On Error Resume Next
'                sFileName = .GetSpecialFolder(2) & "\" & v
'                .CreateTextFile sFileName
'                FileIsValidName = Err.Number = 0
'                If Not FileIsValidName Then GoTo xt
'                On Error GoTo eh
'                If .FileExists(sFileName) Then
'                    .DeleteFile sFileName
'                End If
'            End If
'        Next v
'    End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Private Sub FileInput(ByVal f_file As String, _
'                      ByRef f_strng As String)
'    Const PROC = "FileInput"
'
'    On Error GoTo eh
'    Dim iFileNumber As Integer
'
'    iFileNumber = FreeFile
'    FileClose f_file
'    Open f_file For Input As #iFileNumber
'    f_strng = Input$(LOF(iFileNumber), iFileNumber)
'    dctFilesOpen.Add f_file, iFileNumber
'    FileClose f_file
'
'xt: Exit Sub
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub
'
'Private Sub FileOutput(ByVal f_file As String, _
'                       ByVal f_strng As String, _
'                       ByVal f_append As Boolean)
'    Const PROC = "FileOutput"
'
'    On Error GoTo eh
'    Dim iFileNumber As Integer
'
'    iFileNumber = FreeFile
'    FileClose f_file
'    If f_append _
'    Then Open f_file For Append As #iFileNumber _
'    Else Open f_file For Output As #iFileNumber
'    dctFilesOpen.Add f_file, iFileNumber
'    Print #iFileNumber, f_strng
'    FileClose f_file
'
'xt: Exit Sub
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub
'
'Public Function FilePicked(Optional ByVal p_title As String = "Select a file", _
'                           Optional ByVal p_multi As Boolean = False, _
'                           Optional ByVal p_init_path As String = "C:\", _
'                           Optional ByVal p_filters As String = "All files,*.*", _
'                           Optional ByRef p_file As File = Nothing) As Boolean
'' ----------------------------------------------------------------------------
'' Displays an msoFileDialogFilePicker dialog.
'' - When a file has been selected the function returns TRUE with the selected
''   file as object (p_file).
'' - When no file is selected the function returns FALSE and a file object
''   (p_file) = None.
'' - All arguments have a reasonable default
'' - The filters argument (p_filters) defaults to "All files", "*.*" with the
''   foillowing syntax:
''   <title>,<filter>[<title>,<filter>]...
'' ----------------------------------------------------------------------------
'    Const PROC = "FilePicked"
'
'    On Error GoTo eh
'    Dim v   As Variant
'
'    With Application.FileDialog(msoFileDialogFilePicker)
'        .AllowMultiSelect = p_multi
'        .Title = p_title
'        .InitialFileName = p_init_path
'        .Filters.Clear
'        For Each v In Split(p_filters, ";")
'            .Filters.Add Description:=Trim(Split(v, ",")(0)), Extensions:=Trim(Split(v, ",")(1))
'        Next v
'#If mTrc = 1 Then  ' exclude the time spent for the selection dialog execution
'        mTrc.Pause      ' from the trace
'#ElseIf clsTrc = 1 Then
'        Trc.Pause
'#End If                 ' when the execution trace is active
'        If .Show = -1 Then
'            FilePicked = True
'            Set p_file = fso.GetFile(.SelectedItems(1))
'        Else
'            Set p_file = Nothing
'        End If
'#If mTrc = 1 Then
'        mTrc.Continue
'#ElseIf clsTrc = 1 Then
'        Trc.Continue
'#End If
'     End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function FilesSearch(ByVal f_root As String, _
'                   Optional ByVal f_mask As String = "*", _
'                   Optional ByVal f_in_subfolders As Boolean = True) As Collection
'' ---------------------------------------------------------------------
'' Returns a collection of all file names which meet the criteria:
'' - in any subfolder of the root (f_root)
'' - meeting the wildcard comparison (f_file_mask)
'' ---------------------------------------------------------------------
'    Const PROC = "FilesSearch"
'
'    On Error GoTo eh
'    Dim fo      As Folder
'    Dim sfo     As Folder
'    Dim fl      As File
'    Dim queue   As New Collection
'
'    Set FilesSearch = New Collection
'    If Right(f_root, 1) = "\" Then f_root = Left(f_root, Len(f_root) - 1)
'    If Not fso.FolderExists(f_root) Then GoTo xt
'    queue.Add fso.GetFolder(f_root)
'
'    Do While queue.Count > 0
'        Set fo = queue(queue.Count)
'        queue.Remove queue.Count ' dequeue the processed subfolder
'        If f_in_subfolders Then
'            For Each sfo In fo.SubFolders
'                queue.Add sfo ' enqueue (collect) all subfolders
'            Next sfo
'        End If
'        For Each fl In fo.Files
'            If VBA.Left$(fl.Name, 1) <> "~" _
'            And fl.Name Like f_mask _
'            Then FilesSearch.Add fl
'        Next fl
'    Loop
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function FolderIsValidName(ivf_name As String) As Boolean
'' ----------------------------------------------------------------------------
'' Returns TRUE when the provided argument is a vaslid folder name.
'' ----------------------------------------------------------------------------
'    Const PROC = "IsValidFileOrFolderName"
'
'    On Error GoTo eh
'    Dim a() As String
'    Dim v   As Variant
'    Dim fo  As String
'
'    With CreateObject("VBScript.RegExp")
'        .Pattern = "^(?!(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.[^.]*)?$)[^<>:""/\\|?*\x00-\x1F]*[^<>:""/\\|?*\x00-\x1F\ .]$"
'        FolderIsValidName = Not .Test(ivf_name)
'    End With
'
'    If FolderIsValidName Then
'        '~~ Let's prove the above by a 'brute force method':
'        '~~ Checking each element of the argument whether it can be created as folder
'        '~~ is the final assertion.
'        a = Split(ivf_name, "\")
'        With fso
'            For Each v In a
'                '~~ Check each element of the path (except the drive spec) whether it can be created as a file
'                If InStr(v, ":") = 0 Then ' exclude the drive spec
'                    fo = .GetSpecialFolder(2) & "\" & v
'                    On Error Resume Next
'                    .CreateFolder fo
'                    FolderIsValidName = Err.Number = 0
'                    If Not FolderIsValidName Then GoTo xt
'                    On Error GoTo eh
'                    If .FolderExists(fo) Then
'                        .DeleteFolder fo
'                    End If
'                End If
'            Next v
'        End With
'    End If
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function Folders(Optional ByVal fo_spec As String = vbNullString, _
'                        Optional ByVal fo_subfolders As Boolean = False, _
'                        Optional ByRef fo_result As String) As Collection
'' ----------------------------------------------------------------------------
'' Returns all folders in a folder (fo_spec) - optionally including all
'' sub-folders (fo_subfolders = True) - as folder objects in ascending order.
'' When no folder (fo_spec) is provided a folder selection dialog request one.
'' When the provided folder does not exist or no folder is selected the
'' the function returns with an empty collection. The provided or selected
'' folder is returned (fo_result).
'' ----------------------------------------------------------------------------
'    Const PROC = "Folders"
'
'    Static cll      As Collection
'    Static queue    As Collection   ' FiFo queue for folders with sub-folders
'    Static Stack    As Collection   ' LiFo stack for recursive calls
'    Static foStart  As Folder
'    Dim aFolders()  As Variant
'    Dim fo1         As Folder
'    Dim fo2         As Folder
'    Dim i           As Long
'    Dim j           As Long
'    Dim s           As String
'    Dim v           As Variant
'
'    If cll Is Nothing Then Set cll = New Collection
'    If queue Is Nothing Then Set queue = New Collection
'    If Stack Is Nothing Then Set Stack = New Collection
'
'    If queue.Count = 0 Then
'        '~~ Provide the folder to start with - when not provided by fo_spec via a selection dialog
'        If fo_spec <> vbNullString Then
'            If Not fso.FolderExists(fo_spec) Then
'                fo_result = fo_spec
'                GoTo xt
'            End If
'            Set fo1 = fso.GetFolder(fo_spec)
'        Else
'            Application.DisplayAlerts = False
'            With Application.FileDialog(msoFileDialogFolderPicker)
'                .Title = "Please select the desired folder!"
'                .InitialFileName = CurDir
'                .AllowMultiSelect = False
'                If .Show <> -1 Then GoTo xt
'                Set fo1 = fso.GetFolder(.SelectedItems(1))
'            End With
'        End If
'        Set foStart = fo1
'    Else
'        '~~ When recursively called (Queue.Count <> 0) take first sub-folder queued
'        Set fo1 = queue(1)
'    End If
'
'    For Each fo2 In fo1.SubFolders
'        cll.Add fo2
'        If fo1.SubFolders.Count <> 0 And fo_subfolders Then
'            queue.Add fo2
'        End If
'    Next fo2
'    Stack.Add cll ' stack result in preparation for the function being called resursively
'
'    If queue.Count > 0 Then
'        Debug.Print ErrSrc(PROC) & ": " & "Remove " & queue(1).Path & " from stack (" & queue.Count - 1 & " still in queue)"
'        queue.Remove 1
'    End If
'    If queue.Count > 0 Then
'        Folders queue(1).Path ' recursive call for each folder with subfolders
'    End If
'
'xt: If Stack.Count > 0 Then
'        Set cll = Stack(Stack.Count)
'        Stack.Remove Stack.Count
'    End If
'    If Stack.Count = 0 Then
'        If cll.Count > 0 Then
'            '~~ Unload cll to array, when fo_subfolders = False only those with a ParentFolder foStart
'            ReDim aFolders(cll.Count - 1)
'            For Each v In cll
'                aFolders(i) = v
'                i = i + 1
'            Next v
'
'            '~~ Sort array from A to Z
'            For i = LBound(aFolders) To UBound(aFolders)
'                For j = i + 1 To UBound(aFolders)
'                    If UCase(aFolders(i)) > UCase(aFolders(j)) Then
'                        s = aFolders(j)
'                        aFolders(j) = aFolders(i)
'                        aFolders(i) = s
'                    End If
'                Next j
'            Next i
'
'            '~~ Transfer array as folder objects to collection
'            Set cll = New Collection
'            For i = LBound(aFolders) To UBound(aFolders)
'                Set fo1 = fso.GetFolder(aFolders(i))
'                cll.Add fo1
'            Next i
'        End If
'        Set Folders = cll
'        If Not foStart Is Nothing Then fo_result = foStart.Path
'    End If
'    Set cll = Nothing
'
'End Function
'
'Public Function IsClone(ByVal i_fle1 As String, _
'                        ByVal i_fle2 As String) As Boolean
'' ----------------------------------------------------------------------------
''
'' ----------------------------------------------------------------------------
'    Dim dtCreate1       As Date
'    Dim dtCreate2       As Date
'    Dim dtLastAccess1   As Date
'    Dim dtLastAccess2   As Date
'    Dim dtLastWrite1    As Date
'    Dim dtLastWrite2    As Date
'    Dim fle1            As File
'    Dim fle2            As File
'    Dim sFile1          As String
'    Dim sFile2          As String
'
'
'    ' Get the file objects
'    Set fle1 = fso.GetFile(fle1)
'    Set fle2 = fso.GetFile(fle2)
'
'    ' Compare file sizes
'    If fle1.Size <> fle2.Size Then
'        IsClone = False
'        Exit Function
'    End If
'
'    ' Compare file times
'    dtCreate1 = fle1.DateCreated
'    dtLastAccess1 = fle1.DateLastAccessed
'    dtLastWrite1 = fle1.DateLastModified
'    dtCreate2 = fle2.DateCreated
'    dtLastAccess2 = fle2.DateLastAccessed
'    dtLastWrite2 = fle2.DateLastModified
'
'    If dtCreate1 <> dtCreate2 Or dtLastAccess1 <> dtLastAccess2 Or dtLastWrite1 <> dtLastWrite2 Then
'        IsClone = False
'        Exit Function
'    End If
'
'    ' Compare file contents
'    Open fle1 For Binary As #1
'    sFile1 = Space(LOF(1))
'    Get #1, , sFile1
'    Close #1
'
'    Open fle2 For Binary As #2
'    sFile2 = Space(LOF(2))
'    Get #2, , sFile2
'    Close #2
'
'    If sFile1 <> sFile2 Then
'        IsClone = False
'        Exit Function
'    End If
'
'    ' If all checks pass, the files are clones
'    IsClone = True
'
'End Function
'
'Private Function IsValidFileFullName(ByVal i_s As String, _
'                            Optional ByRef i_fle As File) As Boolean
'' ----------------------------------------------------------------------------
'' Returns True when a string (i_s) is a valid file full name. If True and the file
'' exists it is returned (i_fle).
'' ----------------------------------------------------------------------------
'    Const PROC = "IsValidFileFullName"
'
'    On Error GoTo eh
'    Dim bExists As Boolean
'    Dim fle     As File
'
'    With fso
'        If InStr(i_s, "\") = 0 Then GoTo xt ' not a valid path
'        On Error Resume Next
'        bExists = .FileExists(i_s)
'        If Err.Number = 0 Then
'            On Error GoTo eh
'            If bExists Then
'                IsValidFileFullName = True
'                Set i_fle = .GetFile(i_s)
'            Else
'                On Error Resume Next
'                Set fle = .CreateTextFile(i_s)
'                If Err.Number = 0 Then
'                    IsValidFileFullName = True
'                    .DeleteFile i_s
'                    Exit Function
'                End If
'            End If
'        End If
'        If IsValidFileFullName Then
'            IsValidFileFullName = .GetExtensionName(i_s) <> vbNullString
'        End If
'    End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Public Function KeySort(ByRef s_dct As Dictionary) As Dictionary
'' ------------------------------------------------------------------------------
'' Returns the items in a Dictionary (s_dct) sorted by key.
'' ------------------------------------------------------------------------------
'    Const PROC  As String = "KeySort"
'
'    On Error GoTo eh
'    Dim dct     As New Dictionary
'    Dim vKey    As Variant
'    Dim arr()   As Variant
'    Dim temp    As Variant
'    Dim i       As Long
'    Dim j       As Long
'
'    If s_dct Is Nothing Then GoTo xt
'    If s_dct.Count = 0 Then GoTo xt
'
'    With s_dct
'        ReDim arr(0 To .Count - 1)
'        For i = 0 To .Count - 1
'            arr(i) = .Keys(i)
'        Next i
'    End With
'
'    '~~ Bubble sort
'    For i = LBound(arr) To UBound(arr) - 1
'        For j = i + 1 To UBound(arr)
'            If arr(i) > arr(j) Then
'                temp = arr(j)
'                arr(j) = arr(i)
'                arr(i) = temp
'            End If
'        Next j
'    Next i
'
'    '~~ Transfer based on sorted keys
'    For i = LBound(arr) To UBound(arr)
'        vKey = arr(i)
'        dct.Add key:=vKey, Item:=s_dct.Item(vKey)
'    Next i
'
'xt: Set s_dct = dct
'    Set KeySort = dct
'    Set dct = Nothing
'    Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Private Function Min(ParamArray va() As Variant) As Variant
'' --------------------------------------------------------
'' Returns the minimum (smallest) of all provided values.
'' --------------------------------------------------------
'    Dim v As Variant
'
'    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
'    For Each v In va
'        If v < Min Then Min = v
'    Next v
'
'End Function
'
'Private Function Pop(ByRef s_stk As Collection) As Variant
'    If s_stk.Count > 0 Then
'        If IsObject(s_stk(1)) Then Set Pop = s_stk(1) Else Pop = s_stk(1)
'        s_stk.Remove 1
'    End If
'End Function
'
'Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
'
'    If r_bookmark = vbNullString Then
'        ShellRun GITHUB_REPO_URL
'    Else
'        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
'        ShellRun GITHUB_REPO_URL & r_bookmark
'    End If
'
'End Sub
'
'Public Sub RenameSubFolders(ByVal r_folder_path As String, _
'                            ByVal r_folder_old_name As String, _
'                            ByVal r_folder_new_name As String, _
'                   Optional ByRef r_renamed As Collection = Nothing)
'' ----------------------------------------------------------------------------
'' Rename any sub-folder in the provided path (r_folder_path) named (r_folder_old_name)
'' to (r_folder_new_name). Any subfolders in renamed folders are ignored. Returns
'' all renamed folders as Collection (r_renamed).
'' ----------------------------------------------------------------------------
'    Dim cll As Collection
'    Dim v   As Variant
'    Dim fld As Folder
'
'    Set cll = SubFolders(r_folder_path)
'    For Each v In cll
'        Set fld = v
'        If fld.Name = r_folder_old_name Then
'            fld.Name = r_folder_new_name
'            If Not r_renamed Is Nothing Then r_renamed.Add fld
'        End If
'    Next v
'
'xt: Exit Sub
'
'End Sub
'
'Private Function ShellRun(ByVal sr_string As String, _
'                 Optional ByVal sr_show_how As Long = WIN_NORMAL) As String
'' ----------------------------------------------------------------------------
'' Opens a folder, email-app, url, or even an Access instance.
''
'' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
''                 - Call Email app: ShellRun("mailto:user@tutanota.com")
''                 - Open URL:       ShellRun("http://.......")
''                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
''                                   "Open With" dialog)
''                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
'' Copyright:      This code was originally written by Dev Ashish. It is not to
''                 be altered or distributed, except as part of an application.
''                 You are free to use it in any application, provided the
''                 copyright notice is left unchanged.
'' Courtesy of:    Dev Ashish
'' ----------------------------------------------------------------------------
'
'    Dim lRet            As Long
'    Dim varTaskID       As Variant
'    Dim stRet           As String
'    Dim hWndAccessApp   As Long
'
'    '~~ First try ShellExecute
'    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
'
'    Select Case True
'        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
'        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
'        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
'        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
'        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
'            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, WIN_NORMAL)
'            lRet = (varTaskID <> 0)
'        Case lRet > ERROR_SUCCESS:          lRet = -1
'    End Select
'
'    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)
'
'End Function
'
'Private Function SplitIndctr(ByVal s_strng As String, _
'                    Optional ByRef s_indctr As String = vbNullString) As String
'' ----------------------------------------------------------------------------
'' Returns the split indicator (s_indctr) of a string (s_strng) as string and
'' as argument. The dedection is bypassed in case one (s_indctr) has already
'' been provided.
'' Note: By intention, the list of possibly identified split indicatory is kept
''       to a minimum in order not to interfere with numeric values as strings,
''       list, sentenses, etc.
'' ----------------------------------------------------------------------------
'    If s_indctr = vbNullString Then
'        Select Case True
'            Case InStr(s_strng, vbCrLf) <> 0: s_indctr = vbCrLf
'            Case InStr(s_strng, vbLf) <> 0:   s_indctr = vbLf      ' e.g. in case of a downloaded file's_strng complete string
'            Case InStr(s_strng, "|&|") <> 0:  s_indctr = "|&|"
'        End Select
'    End If
'    SplitIndctr = s_indctr
'
'End Function
'
'Private Function SplitString(ByVal s_s As String) As String
'
'    Select Case True
'        Case InStr(s_s, vbCrLf) <> 0: SplitString = vbCrLf
'        Case InStr(s_s, vbCr) <> 0:   SplitString = vbCr
'        Case InStr(s_s, vbLf) <> 0:   SplitString = vbLf
'    End Select
'    If Len(SplitString) = 0 Then SplitString = vbCrLf
'
'End Function
'
'Private Function StringAsArray(ByVal s_strng As String, _
'                      Optional ByVal s_split As String = vbNullString, _
'                      Optional ByVal s_trim As Variant = True) As Variant
'' ----------------------------------------------------------------------------
'' Returns a string (s_strng) split into an array of strings. When no split
'' indicator (s_split) is provided it one is found by examination of the
'' string (s_strng). When the option (s_trim) is TRUE (the default), "R", or
'' "L" the items in the array are returned trimmed accordingly.
'' Example 1: arr = StringAsArray("this is a string", " ") is returned as an
''            array with 3 items: "this", "is", "a", "string".
'' Example 2: arr = StringAsArray(FileAsString(FileName),sSplit,False) is
''            returned as any array with records/lines of the provided file,
''            whereby the lines are not trimmed, i.e. leading spaces are
''            preserved.
''            Note: The not provided split indicator has the advantage that it
''                  is provided by the SplitIndctr service, which in that case
''                  returns either vbCrLf or vbLf, the latter when the file is
''                  a download.
'' Example 3: arr = FileAsArray(<file>) return the same as example 2.
'' Note: Split indicators dedected by examination are: vbCrLf, vbLf, "|&|",
''       ", ", "; ", "," or ";". When neither is dedected vbCrLf is returned.
'' ----------------------------------------------------------------------------
'    Dim arr As Variant
'    Dim i   As Long
'
'    If s_split = vbNullString Then s_split = SplitIndctr(s_strng)
'    arr = Split(s_strng, SplitIndctr(s_strng, s_split))
'    If Not s_trim = False Then
'        For i = LBound(arr) To UBound(arr)
'            Select Case s_trim
'                Case True:  arr(i) = VBA.Trim$(arr(i))
'                Case "R":   arr(i) = VBA.RTrim$(arr(i))
'                Case "L":   arr(i) = VBA.Trim$(arr(i))
'            End Select
'        Next i
'    End If
'    StringAsArray = arr
'
'End Function
'
'Public Function StringAsFile(ByVal s_strng As String, _
'                    Optional ByRef s_file As Variant = vbNullString, _
'                    Optional ByVal s_file_append As Boolean = False) As File
'' ----------------------------------------------------------------------------
'' Writes a string (s_strng) to a file (s_file) which might be a file object or
'' a file's full name. When no file (s_file) is provided, a temporary file is
'' returned.
'' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
''         is written a records/lines.
'' Note 2: When the string has the alternate split indicator "|&|" this one is
''         replaced by vbCrLf.
'' Note when copied: Originates in mVarTrans
''                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
'' ----------------------------------------------------------------------------
'    Dim sFile As String
'
'    Select Case True
'        Case s_file = vbNullString:         sFile = TempFileFullName
'        Case TypeName(s_file) = "File":     sFile = s_file.Path
'        Case TypeName(s_file) = "String":   sFile = s_file
'    End Select
'
'    FileOutput sFile, s_strng, s_file_append
'    Set StringAsFile = fso.GetFile(sFile)
'
'End Function
'
'Public Function StringEmptyExcluded(ByVal s_s As String) As String
'' ----------------------------------------------------------------------------
'' Returns a string (s_s) with any empty elements excluded. I.e. the string
'' returned begins and ends with a non vbNullString character and has no
'' ----------------------------------------------------------------------------
'    Dim sSplit As String
'
'    '~~ Obtain (and return) the used line terminating string (vbCrLf) or character
'    sSplit = SplitString(s_s)
'
'    s_s = StringTrimmed(s_s) ' leading and trailing empty already excluded
'
'    Do While InStr(s_s, sSplit & sSplit) <> 0
'        s_s = Replace(s_s, sSplit & sSplit, sSplit)
'    Loop
'    StringEmptyExcluded = s_s
'
'End Function
'
'Private Function StringTrimmed(ByVal s_s As String, _
'                      Optional ByRef s_as_dict As Dictionary = Nothing) As String
'' ----------------------------------------------------------------------------
'' Returns a string (s_s) provided as a single string, with any leading and
'' trailing empty items (record/lines) excluded. When a Dictionary is provided
'' the string is additionally returned as items with the line number as key.
'' ----------------------------------------------------------------------------
'    Dim s As String
'    Dim i As Long
'    Dim v As Variant
'
'    s = s_s
'    '~~ Eliminate any leading empty items
'    Do While Left(s, 2) = vbCrLf
'        s = Right(s, Len(s) - 2)
'    Loop
'    '~~ Eliminate a trailing eof if any
'    If Right(s, 1) = VBA.Chr(26) Then
'        s = Left(s, Len(s) - 1)
'    End If
'    '~~ Eliminate any trailing empty items
'    Do While Right(s, 2) = vbCrLf
'        s = Left(s, Len(s) - 2)
'    Loop
'
'    StringTrimmed = s
'    If Not s_as_dict Is Nothing Then
'        With s_as_dict
'            For Each v In Split(s, vbCrLf)
'                i = i + 1
'                .Add i, v
'            Next v
'        End With
'    End If
'
'End Function
'
'Public Function SubFolders(ByVal s_path As String) As Collection
'' ----------------------------------------------------------------------------
'' Returns a Collection of all folders and sub-folders in a folder (s_path).
'' ----------------------------------------------------------------------------
'    Const PROC = "SubFolders"
'
'    Dim stk As New Collection
'    Dim cll As New Collection
'    Dim fld As Folder
'    Dim sfd As Folder ' Subfolder
'
'    Push(stk) = fso.GetFolder(s_path)  ' push the initial folder onto the queue
'    Do While stk.Count > 0
'        Set fld = Pop(stk)             ' pop the first dolder pushed to the queue
'        cll.Add fld
'        Debug.Print ErrSrc(PROC) & ": " & fld.Path
'        For Each sfd In fld.SubFolders
'            Push(stk) = sfd
'        Next sfd
'    Loop
'
'    Set SubFolders = cll
'    Set cll = Nothing
'
'End Function
'
'Public Function TempFileFullName(Optional ByVal f_path As String = vbNullString, _
'                                 Optional ByVal f_ext As String = "txt", _
'                                 Optional ByVal f_create_as_textstream As Boolean = False) As String
'' ------------------------------------------------------------------------------
'' Returns the full file name of a temporary randomly named file. When a path
'' (f_path) is omitted in the CurDir path, else in at the provided folder.
'' The returned temporary file is registered in cllTempTestItems for being removed
'' either explicitly with CleanUp or implicitly when the class
'' terminates.
'' ------------------------------------------------------------------------------
'    Dim s As String
'
'    If VBA.Left$(f_ext, 1) <> "." Then f_ext = "." & f_ext
'
'    s = Replace(fso.GetTempName, ".tmp", f_ext)
'    If f_path = vbNullString Then f_path = CurDir
'    s = VBA.Replace(f_path & "\" & s, "\\", "\")
'    If f_create_as_textstream Then fso.CreateTextFile s
'    TempFileFullName = s
'
'End Function
'
