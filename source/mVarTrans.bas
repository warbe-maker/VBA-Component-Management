Attribute VB_Name = "mVarTrans"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mVarTrans: Transforms/translates/transposes
' ========================== variant items (String, Array,
' Collection, Dictionary, or File) into a String, Array,
' Collection, Dictionary, or File.
'
' Public procedures:
' ------------------
' ArrayAsCollection      Return an array's items as Collection.
' ArrayAsDictionary      Returns a Dictionary with an array's
'                        items as key. Because the array's
'                        items are returned as Directory keys,
'                        the items will become distinct. I e.
'                        each item will exist only once *).
' ArrayAsFile            Writes all items of an array to a file
'                        - which might be a file object, a file's
'                        full name or not provided - as records
'                        /lines. When no file is provided a
'                        temporary file is written and its name
'                        is returned.
' ArrayAsString          Returns an array (a_array) as string
'                        with the items delimited by a vbCrLf.
' CollectionAsArray      Returns a collection's (c_coll) items as
'                        array.
' CollectionAsDictionary Returns a collection's (c_coll) items as
'                        Dictionary keys. Because the
'                        collection's items are returned as
'                        Directory keys, the items will become
'                        distinct. I. e. each item will exist
'                        only once in the Dictionary *).
' CollectionAsFile       Returns the items of a Collection  as
'                        records/lines in a file, optionally
'                        appended.
'                        Uses: StringAsFile, CollectionAsString.
' CollectionAsString     Returns a collection's items as string
'                        with the items delimited by a vbCrLf.
'                        Itmes are converted into a string, if an
'                        item is an object its Name property is
'                        used (an error is raised when the object
'                        has no Name property.
' DictionaryAs Array     Returns a Dictionary's keys as array. In
'                        case a key is an object the object's
'                        Name property is returned (an error is
'                        raised in case the object has no Name
'                        property).
' DictionaryAsCollection Returns a Collection with a Dictionary's
'                        keys as items.
' DictionaryAsFile       Returns a Dictionary's keys as file
'                        records/lines, optionally appended. Keys
'                        are converted into a string, if an item
'                        is an object its Name property is used.
'                        An error is raised when the object has
'                        no Name property.
'                        Uses: StringAsFile, DictionaryAsString.
' DictionaryAsString     Returns a Dictionary's keys as string
'                        with each key delimited by a vbCrLf. Keys
'                        are converted into a string, if an item
'                        is an object its Name property is used.
'                        An error is raised when the object has no
'                        Name property.
' FileAsArray            Returns a file's records/lines as array.
' FileAsCollection       Returns a file's records/lines as
'                        Collection.
' FileAsDictionary       Returns a file's records/lines as
'                        Dictionary keys. Because the lines
'                        become Directory keys, they will become
'                        distinct. I. e. each line will exist
'                        only once in the Dictionary *).
' FileAsString           Returns a file's records/lines as a
'                        single string with the records/lines
'                        delimited by a vbCrLf.
'
' *) To make this restriction productive, the number of
'    occurrences of each item is returned as item.
'
' Uses:
' -----
' clsTestAid      Common services supporting test including
'                 regression testing.
'
' Requires:       Reference to "Mocrodoft Scripting Runtime"
'
' W. Rauschenberger, Berlin Mar 2024
' ----------------------------------------------------------------
Private Const DQUOTE                As String = """"    ' one " character
Private Const ERR_SPLIT_AMBIGOUS    As String = "The string argument has ambigous delimitiers!"

Private cll          As Collection
Private dct          As Dictionary
Private sDelim       As String
Private fso          As New FileSystemObject
Private v            As Variant

'~~ When the a copy of VarTrans and all its depending Functions is used in a VB-Project,
'~~ e.g. to provide an independency of a module, this will have to be copied in a
'~~ Standard module. VarTrans procedure.
Public Enum VarTransAs
    enAsArray
    enAsCollection
    enAsDictionary
    enAsFile
    enAsString
End Enum
'~~ --------------------------------------------------------------------

Private Function AppErr(ByVal app_err_no As Long) As Long
' ----------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error. Returns a given
' positive 'Application Error' number (app_err_no) as a negative by adding the
' system constant vbObjectError. Returns the original 'Application Error'
' number when called with a negative error number.
' Obligatory copy Private for any VB-Component using the service but not
' having the mBasic common component installed.
' ----------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Sub ArrayAdd(ByRef a_arr As Variant, _
                     ByVal a_str As String)
' ----------------------------------------------------------------------------
' Add an item (a_str) to anarray (a_arr) by extending it by one element,
' considering that the array might not be allocated yet.
' ----------------------------------------------------------------------------
                     
    On Error Resume Next
    ReDim Preserve a_arr(UBound(a_arr) + 1)
    If Err.Number <> 0 Then ReDim a_arr(0)
    a_arr(UBound(a_arr)) = a_str
    
End Sub

Public Function ArrayAsCollection(ByVal a_array As Variant) As Collection
' ----------------------------------------------------------------------------
' Return an array's (a_array) items as Collection.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In a_array
            .Add v
        Next v
    End With
    Set ArrayAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function ArrayAsDictionary(ByVal a_array As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In a_array
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
    Set ArrayAsDictionary = dct
    Set dct = Nothing
    
End Function

Public Function ArrayAsFile(ByVal a_array As Variant, _
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
    Set ArrayAsFile = fso.GetFile(a_file)
    
End Function

Public Function ArrayAsString(ByVal a_array As Variant, _
                     Optional ByVal a_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (a_array) as string with the items delimited (a_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArrayAsString = Join(a_array, a_delim)
End Function

Private Function ArrayIsAllocated(ByVal a_arry As Variant) As Boolean
' ----------------------------------------------------------------------------
' Retunrs TRUE when an array (a_array) is allocated.
' ----------------------------------------------------------------------------
    
    On Error Resume Next
    ArrayIsAllocated = UBound(a_arry) >= LBound(a_arry)
    On Error GoTo -1
    
End Function

Public Function CollectionAsArray(ByVal c_coll As Collection) As Variant
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim sName   As String
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then ArrayAdd arr, sName _
        Else ArrayAdd arr, v
    Next v
    CollectionAsArray = arr
    
End Function

Public Function CollectionAsDictionary(ByVal c_coll As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as Dictionary keys.
' Attention: Because the collection's items are returned as Directory keys,
'            the items will be unified. I e. each item will exist only once.
'            To make this restriction productive, the number of occurrences of
'            each item is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In c_coll
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
    Set CollectionAsDictionary = KeySort(dct)
    Set dct = Nothing

End Function

Public Function CollectionAsFile(ByVal c_coll As Collection, _
                        Optional ByRef c_file_name As String = vbNullString, _
                        Optional ByVal c_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns the items of a Collection (c_coll) as records/lines in a file
' (c_file_name), optionally appended (c_file_append).
' Uses: StringAsFile, CollectionAsString.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------

    If c_file_name = vbNullString Then c_file_name = TempFile
    StringAsFile CollectionAsString(c_coll), c_file_name, c_file_append
    Set CollectionAsFile = fso.GetFile(c_file_name)

End Function

Public Function CollectionAsString(ByVal c_coll As Collection) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an item is an object its
' Name property is used (an error is raised when the object has no Name
' property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "CollectionAsString"
    
    Dim s       As String
    Dim sDelim  As String
    Dim sName   As String
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then s = s & sDelim & sName _
        Else s = s & sDelim & v
        sDelim = vbCrLf
    Next v
    CollectionAsString = s

End Function

Public Function DictionaryAsArray(ByVal d_dict As Dictionary) As Variant
' ----------------------------------------------------------------------------
' Returns a Dictionary's keys as array. In case a key is an object the
' object's Name property is returned (an error is raised in case the object
' has no Name property).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    
    For Each v In d_dict
        If IsObject(v, s) _
        Then ArrayAdd arr, s _
        Else ArrayAdd arr, v
    Next v
    DictionaryAsArray = arr
    
End Function

Public Function DictionaryAsCollection(ByVal d_dict As Dictionary) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection with a Dictionary's (d_dict) keys as items.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In d_dict
            .Add v
        Next v
    End With
    Set DictionaryAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function DictionaryAsDictionary(ByVal d_dict As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns the Dictionary (d_dict) with the keys in ascending order.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    With dct
        For Each v In d_dict
            .Add v, vbNullString
        Next v
    End With
    Set DictionaryAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function DictionaryAsFile(ByVal d_dict As Dictionary, _
                        Optional ByRef d_file_name As String = vbNullString, _
                        Optional ByVal d_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a Dictionary's (d_dict) keys as file (d_file_name) records/lines,
' optionally appended (d_file_append). Keys are converted into a string, if an
' item is an object its Name property is used. An error is raised when the
' object has no Name property.
' Uses: StringAsFile, DictionaryAsString.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    If d_file_name = vbNullString Then d_file_name = TempFile
    StringAsFile DictionaryAsString(d_dict), d_file_name, d_file_append
    Set DictionaryAsFile = fso.GetFile(d_file_name)

End Function

Public Function DictionaryAsString(ByVal v_items As Dictionary) As String
' ----------------------------------------------------------------------------
' Returns a Dictionary's keys as string with each key delimited by a vbCrLf.
' Keys are converted into a string, if an item is an object its Name property
' is used. An error is raised when the object has no Name property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sDelim  As String
    Dim sName   As String
    
    For Each v In v_items
        If IsObject(v, sName) _
        Then s = s & sDelim & sName _
        Else s = s & sDelim & v
        sDelim = vbCrLf
    Next v
    DictionaryAsString = s

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
    ErrSrc = "mVarTrans." & sProc
End Function

Public Function FileAsArray(ByVal f_file As Variant, _
                   Optional ByVal f_empty_excluded) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    For Each v In Split(FileAsString(f_file, vbCrLf), vbCrLf)
        If f_empty_excluded Then
            If Trim$(v) <> vbNullString Then ArrayAdd arr, v
        Else
            ArrayAdd arr, v
        End If
    Next v
    FileAsArray = arr
    
End Function

Public Function FileAsCollection(ByVal f_file As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as Collection.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsCollection"
    
    On Error GoTo eh
    Dim cll As New Collection
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    With cll
        For Each v In Split(FileAsString(f_file), vbCrLf)
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

Public Function FileAsDictionary(ByVal f_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Returns a file's records/lines as Dictionary keys.
' Attention: Because the lines become Directory keys, they will become
'            distinct. I. e. each line will exist only once in the Dictionary.
'            To make this restriction productive, the number of occurrences
'            of each line is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim s   As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    With dct
        For Each v In Split(FileAsString(f_file, vbCrLf), vbCrLf)
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
    Set FileAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function FileAsString(ByVal f_file As Variant, _
                    Optional ByRef f_split As String = vbCrLf, _
                    Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns a file's (f_file) - provided as full name or object - records/lines
' as a single string with the records/lines delimited (f_split).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsString"
    
    On Error GoTo eh
    Dim s   As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    '~~ An error is passed on to the caller
    If Not fso.FileExists(f_file) Then Err.Raise AppErr(1), ErrSrc(PROC), _
                                       "The file, provided by a full name, does not exist!"
    
    Open f_file For Input As #1
    s = Input$(lOf(1), 1)
    
    Select Case True
        Case InStr(s, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(s, vbCr) <> 0:   f_split = vbCr
        Case InStr(s, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate any trailing split string
    Do While Right(s, Len(f_split)) = f_split
        s = Left(s, Len(s) - Len(f_split))
        If Len(s) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        s = FileAsStringEmptyExcluded(s)
    End If
    FileAsString = s

xt: Close #1
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsStringEmptyExcluded(ByVal f_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (f_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Do While InStr(f_s, vbCrLf & vbCrLf) <> 0
        f_s = Replace(f_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = f_s
    
End Function

Private Function FileAsStringTrimmed(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a file as string (s_s) with any leading and trailing empty items,
' i.e. record, lines, excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
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

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim i       As Long
    Dim j       As Long
    Dim sName   As String
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            If IsObject(.Keys(i), sName) _
            Then arr(i) = sName _
            Else arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add Key:=vKey, item:=k_dct.item(vKey)
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

Private Function SplitIndctr(ByVal s As String) As String
' ----------------------------------------------------------------------------
' Returns the split indicator in a string (s), which is either vbCrLf or the
' alternate "|&|". When none is found the split indicator defaults to vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "SplitIndctr"
    
    SplitIndctr = vbCrLf ' The dafault
    
    Select Case True
        Case InStr(s, vbCrLf) <> 0: SplitIndctr = vbCrLf
        Case InStr(s, "|&|") <> 0:  SplitIndctr = "|&|"
        Case InStr(s, ",") <> 0:  SplitIndctr = ","
    End Select

End Function

Public Function StringAsArray(ByVal v_items As String) As Variant
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr As Variant
    
    For Each v In Split(v_items, SplitIndctr(v_items))
        ArrayAdd arr, v
    Next v
    StringAsArray = arr

End Function

Public Function StringAsCollection(ByVal v_items As String) As Collection
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim s   As String: s = SplitIndctr(v_items)
    
    With cll
        For Each v In Split(v_items, s)
            .Add v
        Next v
    End With
    Set StringAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function StringAsDictionary(ByVal v_items As String) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (v_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim s   As String: s = SplitIndctr(v_items)
    
    With dct
        For Each v In Split(v_items, s)
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
    Dim sSplit As String
    
    sSplit = SplitIndctr(s_strng)
    If sSplit <> vbCrLf Then s_strng = Replace(s_strng, sSplit, vbCrLf)
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = fso.GetFile(s_file)
    
End Function

Public Function StringAsString(ByVal v_items As String) As String
' ----------------------------------------------------------------------------
' Returns a string (v_item) with any delimiter string replaced with vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "StringAsString"
    
    On Error GoTo eh
    StringAsString = Replace(v_items, SplitIndctr(v_items), vbCrLf)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt", _
                          Optional ByVal f_create_as_textstream As Boolean = True) As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(fso.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    fso.CreateTextFile sTemp

End Function

Public Function VarTrans(ByVal v_items As Variant, _
                         ByVal v_items_as As VarTransAs, _
                Optional ByVal v_items_split As String = vbCrLf, _
                Optional ByRef v_file_name As String = vbNullString, _
                Optional ByVal v_file_append As Boolean = False, _
                Optional ByVal v_empty_excluded As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Universal - primarily text- - items tranformation service. Variant items (v_items) - which
' may be an Array of items, a Dictionary of items, a Dictionary of keys, a
' TextStream file, or a string with items delimited by: vbCrLf, vbLf, ||, |,
' or a , (comma)as Array (v_arr) - are provided as Array, Dictionary,
' Dictionary (with the item as the key), TextStream file, or as a String
' with the items delimited by a vbCrLf.
' Note: When the provision "as Array" (v_items_as) is requested and an item is
'       an object which does not have a Name property, an error is raised and
'       the function is terminated.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "VarTrans"
    
    Dim sSplit  As String
        
    Select Case TypeName(v_items)
        Case "Collection"
            Select Case v_items_as
                Case enAsArray:        VarTrans = CollectionAsArray(v_items)
                Case enAsCollection:   Set VarTrans = v_items
                Case enAsDictionary:   Set VarTrans = CollectionAsDictionary(v_items)
                Case enAsFile:         Set VarTrans = CollectionAsFile(v_items, v_file_name, v_file_append)
                Case enAsString:       VarTrans = CollectionAsString(v_items)
            End Select

        Case "Dictionary"
            Select Case v_items_as
                Case enAsArray:        VarTrans = DictionaryAsArray(v_items)
                Case enAsCollection:   Set VarTrans = DictionaryAsCollection(v_items)
                Case enAsDictionary:   Set VarTrans = v_items
                Case enAsFile:         Set VarTrans = DictionaryAsFile(v_items, v_file_name, v_file_append)
                Case enAsString:       VarTrans = DictionaryAsString(v_items)
            End Select
        
        Case "File":
            Select Case v_items_as
                Case enAsArray:        VarTrans = FileAsArray(v_items, v_empty_excluded)
                Case enAsCollection:   Set VarTrans = FileAsCollection(v_items)
                Case enAsDictionary:   Set VarTrans = FileAsDictionary(v_items)
                Case enAsFile:         Set VarTrans = v_items
                Case enAsString:       VarTrans = FileAsString(v_items)
            End Select
        
        Case "Nothing"
            Select Case v_items_as
                Case enAsArray:        VarTrans = vbNull
                Case enAsCollection:   Set VarTrans = New Collection
                Case enAsDictionary:   Set VarTrans = New Dictionary
                Case enAsFile:         Set VarTrans = Nothing
                Case enAsString:       VarTrans = "Nothing"
            End Select
        
        Case Else
            If IsArray(v_items) Then
                Select Case v_items_as
                    Case enAsArray:        VarTrans = v_items
                    Case enAsCollection:   Set VarTrans = ArrayAsCollection(v_items)
                    Case enAsDictionary:   Set VarTrans = ArrayAsDictionary(v_items)
                    Case enAsFile:         Set VarTrans = ArrayAsFile(v_items, v_file_name, v_file_append)
                    Case enAsString:       VarTrans = ArrayAsString(v_items)
                End Select
            
            Else
                Select Case v_items_as
                    Case enAsArray:        VarTrans = StringAsArray(CStr(v_items))
                    Case enAsCollection:   Set VarTrans = StringAsCollection(CStr(v_items))
                    Case enAsDictionary:   Set VarTrans = StringAsDictionary(CStr(v_items))
                    Case enAsFile:         Set VarTrans = StringAsFile(CStr(v_items), v_file_name, v_file_append)
                    Case enAsString
                        If TypeName(v_items) = "Boolean" Then
                            If v_items = True Then VarTrans = "TRUE" Else VarTrans = "FALSE"
                        Else
                            VarTrans = StringAsString(CStr(v_items))
                        End If
                End Select
            End If
    End Select
            
xt: Exit Function

End Function

