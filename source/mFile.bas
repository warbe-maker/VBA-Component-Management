Attribute VB_Name = "mFile"
Option Explicit
Option Compare Text
Option Private Module
' ----------------------------------------------------------------------------
' Standard  Module mFile
'           Common methods and functions regarding file objects.
'
' Public common file services:
' ----------------------------
' - Exists          Returns TRUE when the provided file exists
' - Compare         Displays the differences of two files by using WinMerge
' - Differs         Returns a Dictionary with those records/lines which
'                   differ between two provided files, with the options
'                   'ignore case' and 'ignore empty records'
' - Delete          Deletes a file provided either as object or as full name
' - Extension       Returns the extension of a file's name
' - GetFile         Returns a file object for a given name
' - Arry        Get Returns the content of a text file as an array
'               Let Writes a text file from the content of an array
'                   of records(lines, Write an array of text to a file.
' - Txt         Get Returns the content of a text file as text string
'               Let Writes a string to a file - optionally appended
'
' Public PrivateProfile file services:
' ------------------------------------
' - SectionExists   Returns TRUE when a given section exists in a given
'                   PrivateProfile file
' - SectionNames    Returns a Dictionary of all section names
'                   [.....] in a file.
' - Sections    Get Returns named - or if no names are provideds all -
'                   sections as Dictionary with the section name as the key
'                   and the Values Dictionary as item
'               Let Writes all sections provided as a Dictionary (as above)
' - RemoveSections  Removes the sections provided via their name. When no
'                   section names are provided (pp_sections) none are
'                   removed.
' - ReArrange   Reorganizes all sections and their value-names in a
'                   PrivateProfile file by ordering everything in ascending
'                   sequence.
' - Value       Get Reads a named value from a PrivateProfile file
'               Let Writes a named value to a PrivateProfile file
' - ValueNameExists Returns TRUE when a value-name exists in a PrivatProfile
'                   file within a given section
' - ValueNames      Returns a Dictionary of all value-names within given
'                   sections in a PrivateProfile file with the value-name
'                   and the section name as key (<name>[section]) and the
'                   value as item, the names in ascending order. Section
'                   names may be provided as a comma delimited string. a
'                   Dictionary or Collection. When no section names are
'                   provided all names of all sections are returned
' - Values          Returns the value-names and values of a given section
'                   in a ProvateProfile file as Dictionary with the
'                   value-name as the key (in ascending order) and the value
'                   as item.
'
' Uses:             No other components. mErH, mMsg, fMsg are optional and
'                   only used when the Conditional Compile Arguments
'                   ErHComp = 1 : MsgComp = 1
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' This 'Common Component' is developed, maintained in the public Github repo:
' https://github.com/warbe-maker/Common-VBA-File-Services.
' Contribution in whichever form is welcome.
'
' W. Rauschenberger, Berlin Jan 2022
' ----------------------------------------------------------------------------
Private Const DQUOTE     As String = """" ' one " character

Private Declare PtrSafe Function WritePrivateProfileString _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal lpw_ApplicationName As String, _
                ByVal lpw_KeyName As String, _
                ByVal lpw_String As String, _
                ByVal lpw_FileName As String) As Long
                
Private Declare PtrSafe Function GetPrivateProfileString _
                Lib "kernel32" Alias "GetPrivateProfileStringA" _
               (ByVal lpg_ApplicationName As String, _
                ByVal lpg_KeyName As String, _
                ByVal lpg_Default As String, _
                ByVal lpg_ReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpg_FileName As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileSection _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal NoKey As Long, _
                ByVal NoSetting As Long, _
                ByVal Name As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileKey _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal Key As String, _
                ByVal Setting As Long, _
                ByVal Name As String) As Long
                 
Private Declare PtrSafe Function GetPrivateProfileSectionNames _
                Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" _
               (ByVal lpszReturnBuffer As String, _
                ByVal nSize As Long, _
                ByVal lpName As String) As Long
                 
'Private Declare PtrSafe Function GetPrivateProfileSection _
'                Lib "kernel32" Alias "GetPrivateProfileSectionA" _
'               (ByVal Section As String, _
'                ByVal Buffer As String, _
'                ByVal Size As Long, _
'                ByVal Name As String) As Long

Public Property Let Arry( _
           Optional ByVal fa_file As String, _
           Optional ByVal fa_excl_empty_lines As Boolean = False, _
           Optional ByRef fa_split As String = vbCrLf, _
           Optional ByVal fa_append As Boolean = False, _
                    ByVal fa_ar As Variant)
' ----------------------------------------------------------------------------
' Writes array (fa_ar) to file (fa_file) whereby the array is joined to a text
' string using the line break string (fa_split) which defaults to vbCrLf and
' is optionally returned by Arry-Get.
' ----------------------------------------------------------------------------
                    
    mFile.Txt(ft_file:=fa_file _
            , ft_append:=fa_append _
            , ft_split:=fa_split _
             ) = Join(fa_ar, fa_split)
             
End Property

Public Property Get Arry( _
           Optional ByVal fa_file As String, _
           Optional ByVal fa_excl_empty_lines As Boolean = False, _
           Optional ByRef fa_split As String, _
           Optional ByVal fa_append As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Returns the content of the file (fa_file) - a files full name - as array,
' with the used line break string returned in (fa_split).
' ----------------------------------------------------------------------------
    Const PROC  As String = "Arry"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim a       As Variant
    Dim a1()    As String
    Dim sSplit  As String
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim i       As Long
    Dim j       As Long
    Dim v       As Variant
    
    If Not fso.FileExists(fa_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named '" & fa_file & "' does not exist!"
    
    '~~ Unload file to a string
    sFile = mFile.Txt(ft_file:=fa_file _
                    , ft_split:=sSplit _
                     )
    If sFile = vbNullString Then GoTo xt
    a = Split(sFile, sSplit)
    
    If Not fa_excl_empty_lines Then
        a1 = a
    Else
        '~~ Extract non-empty items
        For i = LBound(a) To UBound(a)
            If Len(Trim$(a(i))) <> 0 Then cll.Add a(i)
        Next i
        ReDim a1(cll.Count - 1)
        j = 0
        For Each v In cll
            a1(j) = v:  j = j + 1
        Next v
    End If
    
xt: Arry = a1
    fa_split = sSplit
    Set cll = Nothing
    Set fso = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Dict(ByVal fd_file As Variant) As Dictionary
' --------------------------------------------------------------
' Returns the content of the file (fd_file) - which may be
' provided as file object or full file name - as Dictionary by
' considering any kind of line break characters.
' --------------------------------------------------------------
    Const PROC  As String = "Dict-Get"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim a       As Variant
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim fso     As File
    Dim sFile   As String
    Dim i       As Long
    
    If Not Exists(fd_file, fso) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (fd_file) does not exist!"
    
    '~~ Unload file into a test stream
    With New FileSystemObject
        Set ts = .OpenTextFile(fso.Path, 1)
        With ts
            On Error Resume Next ' may be empty
            sFile = .ReadAll
            .Close
        End With
    End With
    
    If sFile = vbNullString Then GoTo xt
    
    '~~ Get the kind of line break used
    If InStr(sFile, vbCr) <> 0 Then sSplit = vbCr
    If InStr(sFile, vbLf) <> 0 Then sSplit = sSplit & vbLf
    
    '~~ Test stream to array
    a = Split(sFile, sSplit)
    
    '~~ Remove any leading or trailing empty items
    ArrayTrimm a
    
    For i = LBound(a) To UBound(a)
        dct.Add i + 1, a(i)
    Next i
        
xt: Set Dict = dct
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

                            
Public Property Get Sections(Optional ByVal pp_file As Variant, _
                             Optional ByVal pp_sections As Variant = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Returns a Dictionary with named sections (pp_sections) in file (pp_file) -
' provided as full name string or as file object - whereby each section is a
' Dictionary with the section name as the key - in ascending
' order - and a Dictionary of the section's values as item with the value name
' as key and the value as item.
' ----------------------------------------------------------------------------
    Const PROC = "Sections-Get"
    
    On Error GoTo eh
    Dim vName   As Variant
    Dim fl      As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    Set Sections = New Dictionary
    
    If pp_sections = vbNullString Then
        '~~ Return all sections
        For Each vName In mFile.SectionNames(fl)
            AddAscByKey add_dct:=Sections _
                      , add_key:=vName _
                      , add_item:=mFile.Values(pp_file:=fl, pp_section:=vName)
        Next vName
    Else
        '~~ Return named sections
        For Each vName In NamesInArg(pp_sections)
            If mFile.SectionExists(pp_file, vName) Then
                AddAscByKey add_dct:=Sections _
                          , add_key:=vName _
                          , add_item:=mFile.Values(pp_file, vName)
            End If
        Next vName
    End If

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Sections( _
                    Optional ByVal pp_file As Variant, _
                    Optional ByVal pp_sections As Variant, _
                             ByVal pp_dct As Dictionary)
' ------------------------------------------------------------------------
' PrivateProfile file service.
' Writes the sections provided as Dictionary (pp_dct) to file (pp_file) -
' provided as full name string or as file object. Existing sections/values
' are overwritten new sections/values are added.
' ------------------------------------------------------------------------
    Const PROC = "Sections-Get"
    
    On Error GoTo eh
    Dim vN          As Variant
    Dim vS          As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim sName       As String
    Dim fl          As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    
    For Each vS In pp_dct
        sSection = vS
        Set dctValues = pp_dct(vS)
        For Each vN In dctValues
            sName = vN
            mFile.Value(pp_file:=fl _
                      , pp_section:=sSection _
                      , pp_value_name:=sName _
                       ) = dctValues.Item(vN)
        Next vN
    Next vS
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get SplitStr(ByRef s As String)
' ----------------------------------------------
' Returns the split string in string (s) used by
' VBA.Split() to turn the string into an array.
' ----------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
End Property

Public Property Get Temp( _
          Optional ByVal tmp_path As String = vbNullString, _
          Optional ByVal tmp_extension As String = ".tmp") As String
' ------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file,
' when tmp_path is omitted in the CurDir path.
' ------------------------------------------------------------------
    
    Dim fso     As New FileSystemObject
    Dim sTemp   As String
    
    If VBA.Left$(tmp_extension, 1) <> "." Then tmp_extension = "." & tmp_extension
    sTemp = Replace(fso.GetTempName, ".tmp", tmp_extension)
    If tmp_path = vbNullString Then tmp_path = CurDir
    sTemp = VBA.Replace(tmp_path & "\" & sTemp, "\\", "\")
    Temp = sTemp
    
    Set fso = Nothing
End Property

Public Property Get Txt( _
         Optional ByVal ft_file As Variant, _
         Optional ByVal ft_append As Boolean = True, _
         Optional ByRef ft_split As String) As String
' ----------------------------------------------------------------------------
' Returns the text file's (ft_file) content as string with VBA.Split() string
' in (ft_split). When the file doesn't exist a vbNullString is returned.
' Note: ft_append is not used but specified to comply with Property Let.
' ----------------------------------------------------------------------------
    Const PROC = "Txt-Get"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim ts      As TextStream
    Dim s       As String
    Dim sFl As String
   
    ft_split = ft_split  ' not used! for declaration compliance and dead code check only
    ft_append = ft_append ' not used! for declaration compliance and dead code check only
    
    With fso
        If TypeName(ft_file) = "File" Then
            sFl = ft_file.Path
        Else
            '~~ ft_file is regarded a file's full name, created if not existing
            sFl = ft_file
            If Not .FileExists(sFl) Then GoTo xt
        End If
        Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForReading)
    End With
    
    If Not ts.AtEndOfStream Then
        s = ts.ReadAll
        ft_split = SplitStr(s)
        If VBA.Right$(s, 2) = vbCrLf Then
            s = VBA.Left$(s, Len(s) - 2)
        End If
    Else
        Txt = vbNullString
    End If
    If Txt = vbCrLf Then Txt = vbNullString Else Txt = s

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Txt( _
         Optional ByVal ft_file As Variant, _
         Optional ByVal ft_append As Boolean = True, _
         Optional ByRef ft_split As String, _
                  ByVal ft_string As String)
' ----------------------------------------------------------------------------
' Writes the string (ft_string) into the file (ft_file) which might be a file
' object or a file's full name.
' Note: ft_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "Txt-Let"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim sFl As String
   
    ft_split = ft_split ' not used! just for coincidence with Get
    With fso
        If TypeName(ft_file) = "File" Then
            sFl = ft_file.Path
        Else
            '~~ ft_file is regarded a file's full name, created if not existing
            sFl = ft_file
            If Not .FileExists(sFl) Then .CreateTextFile sFl
        End If
        
        If ft_append _
        Then Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForWriting)
    End With
    
    ts.WriteLine ft_string

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Value( _
           Optional ByVal pp_file As Variant, _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' PrivateProfile file service. Reads a value with a specific name from a
' section from a file organized: [section]
'                                <value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC  As String = "ValueGet"
    
    On Error GoTo eh
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant
    Dim fl      As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    
    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString( _
                                      lpg_ApplicationName:=pp_section _
                                    , lpg_KeyName:=pp_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=fl _
                                     )
    vValue = Left$(sRetVal, lResult)
    Value = vValue
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Value( _
           Optional ByVal pp_file As Variant, _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' PrivateProfile file service. Writes a value under a given name
' (pp_value_name) into a section (pp_section) in a file (pp_file) organized:
' [section]
' <value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC = "ValueLet"
        
    On Error GoTo eh
    Dim lChars  As Long
    Dim sValue  As String
    Dim fl      As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    
    Select Case VarType(pp_value)
        Case vbBoolean: sValue = VBA.CStr(VBA.CLng(pp_value))
        Case Else:      sValue = pp_value
    End Select
    
    lChars = WritePrivateProfileString(lpw_ApplicationName:=pp_section _
                                     , lpw_KeyName:=pp_value_name _
                                     , lpw_String:=sValue _
                                     , lpw_FileName:=fl)
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & pp_section & "'" & vbLf & _
               "Value name = '" & pp_value_name & "'" & vbLf & _
               "Value      = '" & CStr(pp_value) & "'" & vbLf & _
               "Value file = '" & pp_file & "'"
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Sub AddAscByKey(ByRef add_dct As Dictionary, _
                           ByVal add_key As Variant, _
                           ByVal add_item As Variant)
' ----------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by
' the key (add_key). When the key is an object with no Name property an error
' is raised.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be
'       copied into any VBProject's module in order to have it independant
'       from this Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ----------------------------------------------------------------------------
    Const PROC = "DctAdd"
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItem           As Variant
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim vValueTarget    As Variant ' the add before/after add_key/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    On Error GoTo eh
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    bOrderByItem = False
    bOrderByKey = True
    bSeqAscending = True
    bCaseIgnored = False
    bCaseSensitive = True
    bStayWithFirst = True
    bEntrySequence = False
    
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not bStayWithFirst Then
            If .Exists(add_key) Then
                If VarType(add_item) = vbObject Then Set .Item(add_key) = add_item Else .Item(add_key) = add_item
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If VarType(add_key) = vbObject Then
            On Error Resume Next
            add_key.Name = add_key.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If VarType(add_item) = vbObject Then
            On Error Resume Next
            add_item.Name = add_item.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = AddAscByKeyValue(add_key)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = AddAscByKeyValue(.Keys()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If VarType(add_dct.Item(vKeyExisting)) = vbObject _
        Then Set vItemExisting = add_dct.Item(vKeyExisting) _
        Else vItemExisting = add_dct.Item(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = AddAscByKeyValue(vKeyExisting)
            
                If vValueExisting = vValueNew And bOrderByItem And bSeqAscending And Not .Exists(add_key) Then
                    If bStayWithFirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and bStayWithFirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AddAscByKeyValue(ByVal add_key As Variant) As Variant
' ----------------------------------------------------------------------------
' When add_key is an object its name becomes the sort order value else the
' the value is returned as is.
' ----------------------------------------------------------------------------
    If VarType(add_key) = vbObject Then
        On Error Resume Next ' the object may not have a Name property
        AddAscByKeyValue = add_key.Name
        If Err.Number <> 0 Then Set AddAscByKeyValue = add_key
    Else
        AddAscByKeyValue = add_key
    End If
End Function

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

Private Function ArrayIsAllocated(arr As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = _
    IsArray(arr) _
    And Not IsError(LBound(arr, 1)) _
    And LBound(arr, 1) <= UBound(arr, 1)
    
End Function

Private Function ArrayNoOfDims(arr As Variant) As Integer
' ------------------------------------------------------
' Returns the number of dimensions of an array. An un-
' allocated dynamic array has 0 dimensions. This may as
' as well be tested by means of ArrayIsAllocated.
' ------------------------------------------------------

    On Error Resume Next
    Dim Ndx As Integer
    Dim Res As Integer
    
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    ArrayNoOfDims = Ndx - 1

End Function

Private Sub ArrayRemoveItems(ByRef va As Variant, _
                    Optional ByVal Element As Variant, _
                    Optional ByVal Index As Variant, _
                    Optional ByVal NoOfElements = 1)
' ------------------------------------------------------------------------------
' Returns the array (va) with the number of elements (NoOfElements) removed
' whereby the start element may be indicated by the element number 1,2,...
' (vElement) or the index (Index) which must be within the array's LBound to
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
    Dim NoOfElementsInArray As Long
    Dim i                   As Long
    Dim iNewUBound          As Long
    
    If Not IsArray(va) Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Array not provided!"
    Else
        a = va
        NoOfElementsInArray = UBound(a) - LBound(a) + 1
    End If
    If Not ArrayNoOfDims(a) = 1 Then
        Err.Raise AppErr(2), ErrSrc(PROC), "Array must not be multidimensional!"
    End If
    If Not IsNumeric(Element) And Not IsNumeric(Index) Then
        Err.Raise AppErr(3), ErrSrc(PROC), "Neither FromElement nor FromIndex is a numeric value!"
    End If
    If IsNumeric(Element) Then
        iElement = Element
        If iElement < 1 _
        Or iElement > NoOfElementsInArray Then
            Err.Raise AppErr(4), ErrSrc(PROC), "vFromElement is not between 1 and " & NoOfElementsInArray & " !"
        Else
            iIndex = LBound(a) + iElement - 1
        End If
    End If
    If IsNumeric(Index) Then
        iIndex = Index
        If iIndex < LBound(a) _
        Or iIndex > UBound(a) Then
            Err.Raise AppErr(5), ErrSrc(PROC), "FromIndex is not between " & LBound(a) & " and " & UBound(a) & " !"
        Else
            iElement = ElementOfIndex(a, iIndex)
        End If
    End If
    If iElement + NoOfElements - 1 > NoOfElementsInArray Then
        Err.Raise AppErr(6), ErrSrc(PROC), "FromElement (" & iElement & ") plus the number of elements to remove (" & NoOfElements & ") is beyond the number of elelemnts in the array (" & NoOfElementsInArray & ")!"
    End If
    
    For i = iIndex + NoOfElements To UBound(a)
        a(i - NoOfElements) = a(i)
    Next i
    
    iNewUBound = UBound(a) - NoOfElements
    If iNewUBound < 0 Then Erase a Else ReDim Preserve a(LBound(a) To iNewUBound)
    va = a
    
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
        ArrayRemoveItems a, Index:=i
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

Public Function Compare(ByVal fc_file_left As String, _
                        ByVal fc_left_title As String, _
                        ByVal fc_file_right As String, _
                        ByVal fc_right_title As String) As Long
' ----------------------------------------------------------------------------
' Compares two text files by means of WinMerge. An error is raised when
' WinMerge is not installed of one of the two files doesn't exist.
' ----------------------------------------------------------------------------
    Const PROC = "Compare"
    
    On Error GoTo eh
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String
    Dim fso             As New FileSystemObject
    Dim wshShell        As Object
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , source:=ErrSrc(PROC) _
                 , Description:="WinMerge is obligatory for the Compare service of this module but not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(fc_file_left) _
    Then Err.Raise Number:=AppErr(2) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & fc_file_left & """ does not exist!"
    
    If Not fso.FileExists(fc_file_right) _
    Then Err.Raise Number:=AppErr(3) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & fc_file_right & """ does not exist!"
    
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & fc_left_title & DQUOTE & _
               " /dr " & DQUOTE & fc_right_title & DQUOTE & " " & _
               """" & fc_file_left & """" & " " & _
               """" & fc_file_right & """"
    
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        Compare = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Delete(ByVal v As Variant)

    With New FileSystemObject
        If TypeName(v) = "File" Then
            .DeleteFile v.Path
        ElseIf TypeName(v) = "String" Then
            If .FileExists(v) Then .DeleteFile v
        End If
    End With
    
End Sub

Public Function Differs( _
                  ByVal fd_file1 As File, _
                  ByVal fd_file2 As File, _
         Optional ByVal fd_stop_after As Long = 0, _
         Optional ByVal fd_ignore_empty_records As Boolean = False, _
         Optional ByVal fd_compare As VbCompareMethod = vbTextCompare) As Dictionary
' -----------------------------------------------------------------------------
' Returns TRUE when the content of file (fd_file1) differs from the content in
' file (fd_file2). The comparison stops after (fd_stop_after) detected
' differences. The detected different lines are optionally returned (vResult).
' ------------------------------------------------------------------------------
    Const PROC = "Differs"
    
    On Error GoTo eh
    Dim a1          As Variant
    Dim a2          As Variant
    Dim dct1        As New Dictionary
    Dim dct2        As New Dictionary
    Dim dctDif      As Dictionary
    Dim dctF1       As New Dictionary
    Dim dctF2       As New Dictionary
    Dim i           As Long
    Dim lDiffLine   As Long
    Dim s1          As String
    Dim s2          As String
    Dim sFile1      As String
    Dim sFile2      As String
    Dim sSplit      As String
    Dim sTest1      As String
    Dim sTest2      As String
    Dim v           As Variant
    
    Set dctDif = New Dictionary
    sFile1 = fd_file1.Path
    sFile2 = fd_file2.Path
    
    s1 = mFile.Txt(ft_file:=sFile1, ft_split:=sSplit)
    If fd_ignore_empty_records Then
        '~~ Eliminate empty records
        sTest1 = VBA.Replace$(s1, sSplit & sSplit, sSplit)
    End If
    
    s2 = mFile.Txt(ft_file:=sFile2, ft_split:=sSplit)
    If fd_ignore_empty_records Then
        '~~ Eliminate empty records
        sTest2 = VBA.Replace$(s2, sSplit & sSplit, sSplit)
    End If
    
    If VBA.StrComp(s1, s2, fd_compare) = 0 Then GoTo xt

     
    a1 = Split(s1, sSplit)
    For i = LBound(a1) To UBound(a1)
        dctF1.Add i + 1, a1(i)
        If fd_ignore_empty_records Then
            If VBA.Trim$(a1(i)) <> vbNullString Then
                dct1.Add i + 1, a1(i)
            End If
        Else
            dct1.Add i + 1, a1(i)
        End If
    Next i
    
    a2 = Split(s2, sSplit)
    For i = LBound(a2) To UBound(a2)
        dctF2.Add i + 1, a2(i)
        If fd_ignore_empty_records Then
            If VBA.Trim$(a2(i)) <> vbNullString Then
                dct2.Add i + 1, a2(i)
            End If
        Else
            dct2.Add i + 1, a2(i)
        End If
    Next i
    If VBA.StrComp(Join(dct1.Items(), sSplit), Join(dct2.Items(), sSplit), fd_compare) = 0 Then GoTo xt
    
    '~~ Get and detect the difference by comparing the items one by one
    '~~ and optaining the line number from the Dictionary when different
    If dct1.Count <= dct2.Count Then
        For Each v In dct1 ' v - 1 = array index
            If VBA.StrComp(a1(v - 1), a2(v - 1), fd_compare) <> 0 Then
                lDiffLine = v
                dctDif.Add lDiffLine, DiffItem(di_line:=lDiffLine _
                                             , di_file_left:=sFile1 _
                                             , di_file_right:=sFile2 _
                                             , di_line_left:=a1(v - 1) _
                                             , di_line_right:=a2(v - 1) _
                                              )
                If fd_stop_after > 0 And dctDif.Count >= fd_stop_after Then GoTo xt
            End If
        Next v
        
        For i = dct1.Count + 1 To dct2.Count
            lDiffLine = dct2.Keys()(i - 1)
            dctDif.Add lDiffLine, DiffItem(di_line:=lDiffLine _
                                         , di_file_left:=sFile1 _
                                         , di_file_right:=sFile2 _
                                         , di_line_right:=a2(i - 1) _
                                          )
        Next i

    ElseIf dct2.Count < dct1.Count Then
        For Each v In dct2 ' v - 1 = array index
            If VBA.StrComp(a1(v - 1), a2(v - 1), fd_compare) <> 0 Then
                lDiffLine = v
                dctDif.Add lDiffLine, DiffItem(di_line:=lDiffLine _
                                             , di_file_left:=sFile1 _
                                             , di_file_right:=sFile2 _
                                             , di_line_left:=a1(v - 1) _
                                             , di_line_right:=a2(v - 1) _
                                              )
                If fd_stop_after > 0 And dctDif.Count >= fd_stop_after Then GoTo xt
            End If
        Next v
        For i = dct2.Count + 1 To dct1.Count
            lDiffLine = dct1.Keys()(i - 1)
            dctDif.Add lDiffLine, DiffItem(di_line:=lDiffLine _
                                         , di_file_left:=sFile1 _
                                         , di_file_right:=sFile2 _
                                         , di_line_left:=a1(i - 1) _
                                          )
        Next i
    End If
        
xt: Set Differs = dctDif
    Set dct1 = Nothing
    Set dct2 = Nothing
    Set dctF1 = Nothing
    Set dctF2 = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function DiffItem( _
                    ByVal di_line As Long, _
                    ByVal di_file_left As String, _
                    ByVal di_file_right As String, _
           Optional ByVal di_line_left As String = vbNullString, _
           Optional ByVal di_line_right As String = vbNullString) As String
' --------------------------------------------------------------------
'
' --------------------------------------------------------------------
    Dim sFileLeft   As String
    Dim sFileRight  As String
    Dim i           As Long
    
    For i = 1 To Min(Len(di_file_left), Len(di_file_right))
        If VBA.Mid$(di_file_left, i, 1) <> VBA.Mid$(di_file_right, i, 1) _
        Then Exit For
    Next i
    i = i - 2
    sFileLeft = "..." & VBA.Right$(di_file_left, Len(di_file_left) - i) & "Line " & Format(di_line, "0000") & ": "
    sFileRight = "..." & VBA.Right$(di_file_right, Len(di_file_right) - i) & "Line " & Format(di_line, "0000") & ": "
    
    DiffItem = sFileLeft & "'" & di_line_left & "'" & vbLf & sFileRight & "'" & di_line_right & "'"

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
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
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
    If err_source = vbNullString Then err_source = Err.source
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
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mFile." & sProc
End Function

Public Function Exists(Optional ByVal ex_folder As String = vbNullString, _
                       Optional ByVal ex_file As String = vbNullString, _
                       Optional ByVal ex_section As String = vbNullString, _
                       Optional ByVal ex_value_name As String = vbNullString, _
                       Optional ByRef ex_result_folder As Folder = Nothing, _
                       Optional ByRef ex_result_files As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Universal File existence check whereby the existence checks depend on the
' provided arguments. The function returns TRUE when:
'
' Argument      | TRUE condition (despite the fact not vbNullString)
' --------------| ------------------------------------------------------------
' ex_folder     | The folder exists, no ex_file provided
' ex_file       | When no ex_folder had been provided the provided ex_file
'               | exists. When an ex_folder had been provided at least one
'               | or more ex_file meet the LIKE criteria and ex_section is
'               | not provided
' ex_section    | Exactly one file had been passed the existenc check, the
'               | provided section exists and no ex_value_name is provided.
' ex_value_name | The provided value-name exists - in the existing section
'               | in the one and only existing file.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Exists"
    
    On Error GoTo eh
    Dim sTest           As String
    Dim sFile           As String
    Dim fo              As Folder   ' Folder
    Dim sfo             As Folder   ' Sub-Folder
    Dim fl              As File
    Dim queue           As Collection
    Dim fso             As New FileSystemObject
    Dim FolderExists    As Boolean
    Dim FileExists      As Boolean
    Dim SectionExists   As Boolean
    Dim ValueNameExists As Boolean
    Dim sFolder         As String
    
    Set ex_result_files = New Collection

    With fso
        If Not ex_folder = vbNullString Then
            '~~ Folder existence check
            If Not .FolderExists(ex_folder) Then GoTo xt
            Set ex_result_folder = .GetFolder(ex_folder)
            If ex_file = vbNullString Then
                '~~ When no ex_file is provided, that's it
                Exists = True
                GoTo xt
            End If
        End If
        
        If ex_file <> vbNullString And ex_folder <> vbNullString Then
            '~~ For the existing folder an ex_file argument had been provided
            '~~ This is interpreted as a "Like" existence check is due which
            '~~ by default includes all subfolders
            sFile = ex_file
            Set fo = .GetFolder(ex_folder)
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
                    And fl.Name Like ex_file Then
                        '~~ The file in the (sub-)folder meets the search criteria
                        '~~ In case the ex_file does not contain any "LIKE"-wise characters
                        '~~ only one file may meet the criteria
                        ex_result_files.Add fl
                        Exists = True
                     End If
                Next fl
            Loop
            If ex_result_files.Count <> 1 Then
                '~~ None of the files in any (sub-)folder matched with ex_file
                '~~ or more than one file matched
                GoTo xt
            End If
        ElseIf ex_file <> vbNullString And ex_folder = vbNullString Then
            If Not .FileExists(ex_file) Then GoTo xt
            ex_result_files.Add .GetFile(ex_file)
            If ex_section = vbNullString Then
                '~~ When no section is provided, that's it
                Exists = True
                GoTo xt
            End If
        End If
        
        '~~ At this point either a provided folder together with a provided file matched exactly one existing file
        '~~ or a specified file's existence had been proved
        If ex_section <> vbNullString Then
            If Not mFile.SectionExists(pp_file:=ex_file, pp_section:=ex_section) Then GoTo xt
            If ex_value_name = vbNullString Then
                '~~ When no ex_value_name is provided, that's it
                Exists = True
            Else
                Exists = mFile.ValueExists(pp_file:=ex_file, pp_section:=ex_section, pp_value_name:=ex_value_name)
            End If
        End If
    End With
        
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Extension(ByVal fe_file As Variant)

    With New FileSystemObject
        If TypeName(fe_file) = "File" Then
            Extension = .GetExtensionName(fe_file.Path)
        Else
            Extension = .GetExtensionName(fe_file)
        End If
    End With

End Function

Private Function Fc(ByVal fc_file1 As String, fc_file2 As String)
    Const PROC = "Fc"
    
    On Error GoTo eh
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String
    Dim fso             As New FileSystemObject
    Dim wshShell        As Object
    
    If Not fso.FileExists(fc_file1) _
    Then Err.Raise Number:=AppErr(2) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & fc_file1 & """ does not exist!"
    
    If Not fso.FileExists(fc_file2) _
    Then Err.Raise Number:=AppErr(3) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & fc_file2 & """ does not exist!"
    
    sCommand = "Fc /C /W " & _
               """" & fc_file1 & """" & " " & _
               """" & fc_file2 & """"
    
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        Fc = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function GetFile(ByVal fg_path As String) As File
    With New FileSystemObject
        Set GetFile = .GetFile(fg_path)
    End With
End Function

Public Function IsValidFileName(ivf_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is a vaslid file name.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileOrFolderName"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim fso     As New FileSystemObject
    Dim a()     As String
    Dim i       As Long
    Dim v       As Variant
    Dim IsValidFolderName   As Boolean
    
    '~~ Check each element of the argument whether it can be created as file
    '~~ !!! this is the brute force method to check valid file names
    a = Split(ivf_name, "\")
    i = UBound(a)
    With fso
        For Each v In a
            '~~ Check each element of the path (except the drive spec) whether it can be created as a file
            If InStr(v, ":") = 0 Then ' exclude the drive spec
                On Error Resume Next
                sFile = .GetSpecialFolder(2) & "\" & v
                .CreateTextFile sFile
                IsValidFileName = Err.Number = 0
                If Not IsValidFileName Then GoTo xt
                On Error GoTo eh
                If .FileExists(sFile) Then
                    .DeleteFile sFile
                End If
            End If
        Next v
    End With
    
xt: Set fso = Nothing
    Exit Function
 
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsValidFolderName(ivf_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided argument is a vaslid folder name.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileOrFolderName"
    
    On Error GoTo eh
    Dim a() As String
    Dim i   As Long
    Dim v   As Variant
    Dim fo  As String
    Dim fso As New FileSystemObject
    
    With CreateObject("VBScript.RegExp")
        .PATTERN = "^(?!(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.[^.]*)?$)[^<>:""/\\|?*\x00-\x1F]*[^<>:""/\\|?*\x00-\x1F\ .]$"
        IsValidFolderName = Not .Test(ivf_name)
    End With
    
    If IsValidFolderName Then
        '~~ Let's prove the above by a 'brute force method':
        '~~ Checking each element of the argument whether it can be created as folder
        '~~ is the final assertion.
        a = Split(ivf_name, "\")
        With fso
            For Each v In a
                '~~ Check each element of the path (except the drive spec) whether it can be created as a file
                If InStr(v, ":") = 0 Then ' exclude the drive spec
                    fo = .GetSpecialFolder(2) & "\" & v
                    On Error Resume Next
                    .CreateFolder fo
                    IsValidFolderName = Err.Number = 0
                    If Not IsValidFolderName Then GoTo xt
                    On Error GoTo eh
                    If .FolderExists(fo) Then
                        .DeleteFolder fo
                    End If
                End If
            Next v
        End With
    End If
    
xt: Set fso = Nothing
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

Private Function NamesInArg( _
            Optional ByVal v As Variant = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns (v) as Collection of string items whereby (v) may be: not provided,
' a comma delimited string, a Dictionary, or a Collection of string items.
' ----------------------------------------------------------------------------
    Const PROC = "NamesInArg"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim dct     As Dictionary
    Dim vName   As Variant
    
    Select Case VarType(v)
        Case vbObject
            Select Case TypeName(v)
                Case "Dictionary"
                    Set dct = v
                    For Each v In dct
                        cll.Add dct.Item(v)
                    Next v
                Case "Collection"
                    Set cll = v
                Case Else: GoTo xt ' likely Nothing
            End Select
        Case vbString
            If v <> vbNullString Then
                For Each vName In Split(v, ",")
                    cll.Add VBA.Trim$(v)
                Next vName
            End If
        Case Is >= vbArray
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a string, an arry, a Collecton, or a Dictionary!"
    End Select
            
xt: Set NamesInArg = cll
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub PathFileSplit(ByRef spf_path As String, _
                Optional ByRef spf_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Returns the provided string (spf_path) split in path and file name provided
' the argument is either an existing folder's path or an existing file's path.
' - When sp_path is a folder only string sp_file = vbNullString,
' - When sp_path only contains a file name sp_folder = vbNullString and
'   sp_file = the file's name
' When the provided path is neither an existing folder's path nor an existing
' file's full name both arguments are returned a vbNullString.
' Note: For a non existing folder or file it is impossible to decide whether
'       the string ends with the name of a folder or the name of a file
'       both have exactly the the same restrictions regardin the name. I.e.
'       a file name xxxx.dat may as well be the name of a folder.
' ----------------------------------------------------------------------------
    Const PROC = "PathFileSplit"
    
    On Error GoTo eh
    Dim aPath() As String
    Dim fso     As New FileSystemObject
    Dim i       As Long
    Dim sFolder As String
    Dim sFile   As String
    
    aPath() = Split(spf_path, "\")  'Put the Parts of our path into an array
    i = UBound(aPath)
    sFile = aPath(i)                ' A valid name only when the argument named an existing of a nonn existing but valid file's name
    aPath(i) = ""                   ' Remove the "maybe-a-file" from the array
    sFolder = Join(aPath, "\")      ' Rebuild the path from the remaing array
        
    With fso
        If .FileExists(spf_path) Then
            '~~ The provided argument ends with a valid file name and thus is a full file name
            spf_path = sFolder
            spf_file = sFile
        ElseIf .FolderExists(spf_file) Then
            '~~ The provided argument ends with a valid folder name and thus is a folder's path
            spf_file = vbNullString
        Else
            Debug.Print ErrSrc(PROC) & ": Neither a folder nor a file exists with the provided string '" & spf_path & "'!"
            spf_path = vbNullString
            spf_file = vbNullString
        End If
    End With
            
xt: Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Picked( _
              Optional ByVal p_title As String = "Select a file", _
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
    Const PROC = "Picked"
   
    On Error GoTo eh
    Dim fd  As FileDialog
    Dim v   As Variant
    Dim fso As New FileSystemObject
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = p_multi
        .Title = p_title
        .InitialFileName = p_init_path
        .Filters.Clear
        For Each v In Split(p_filters, ";")
            .Filters.Add Description:=Trim(Split(v, ",")(0)), Extensions:=Trim(Split(v, ",")(1))
        Next v
#If ExecTrace = 1 Then  ' exclude the time spent for the selection dialog execution
        mTrc.Pause      ' from the trace
#End If                 ' when the execution trace is active
        If .Show = -1 Then
            Picked = True
            Set p_file = fso.GetFile(.SelectedItems(1))
        Else
            Set p_file = Nothing
        End If
#If ExecTrace = 1 Then
        mTrc.Continue
#End If
     End With
     
xt: Set fso = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function PPFile(ByVal pp_file As Variant, _
                        ByVal pp_service As String, _
                        ByRef pp_file_name As String) As String
' ----------------------------------------------------------------------------
' PrivateProfile file function. Returns the PrivateProfile file (pp_file) as
' full file name string (pp_file_name). When no file (pp_file) is provided,
' i.e. is a vbNullString, a file selection dialog is evoked to select one.
' When finally there's still no file provided a vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "PPFile"
    
    On Error GoTo eh
    Dim fl  As File
    
    If VarType(pp_file) = vbObject Then
        If pp_file Is Nothing Then
            If Not mFile.Picked(p_title:="Select a PrivateProfile file not provided for the service '" & pp_service & "'" _
                              , p_init_path:=ThisWorkbook.Path _
                              , p_filters:="Config files, *.cfg; Init files, *.ini; Application data files, *.dat; All files, *.*" _
                              , p_file:=fl _
                               ) _
            Then GoTo xt
            pp_file_name = fl.Path
            PPFile = fl.Path
        ElseIf TypeName(pp_file) = "File" Then
            pp_file_name = pp_file.Path
            PPFile = pp_file.Path
            GoTo xt
        End If
    ElseIf VarType(pp_file) = vbString Then
        If pp_file = vbNullString Then
            If Not mFile.Picked(p_title:="Select a PrivateProfile file" _
                              , p_init_path:=ThisWorkbook.Path _
                              , p_filters:="Config files, *.cfg; Init files, *.ini; Application data files, *.dat; All files, *.*" _
                              , p_file:=fl _
                               ) _
            Then GoTo xt
            pp_file_name = fl.Path
            PPFile = fl.Path
        Else
            pp_file_name = pp_file
            PPFile = pp_file
        End If
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ReArrange(Optional ByVal pp_file As Variant = Nothing)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Rearranges all sections and all value-names therein in ascending order.
' When no file (pp_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notic.
' Restriction: The file must not contain any comment lines othe than at the
'              top of the very first sections.
' ----------------------------------------------------------------------------
    Const PROC = "ReArrange"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim Section     As String
    Dim dctSections As Dictionary
    Dim dctValues   As Dictionary
    Dim vValue      As Variant
    Dim fl          As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    
    Set dctSections = Sections(fl)
    For Each vSection In dctSections
        Section = vSection
        Set dctValues = dctSections(vSection)
        
        mFile.RemoveSections pp_file:=fl, pp_sections:=Section
        For Each vValue In dctValues
            mFile.Value(pp_file:=fl, pp_section:=Section, pp_value_name:=vValue) = dctValues(vValue)
        Next vValue
    Next vSection

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RemoveNames(ByVal pp_file As Variant, _
                       ByVal pp_section As String, _
                       ByVal pp_value_names As Variant)
' -----------------------------------------------------------------------------
' PrivateProfile file service.
' Removes from the section (pp_section) in the PrivateProfile file (pp_file)
' a names (pp_value_names) provided either as: comma delimited string,
' Dictionary with name items, or Collection of names.
' When no file (pp_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notice.
' -----------------------------------------------------------------------------
    Const PROC = "RemoveNames"
    
    On Error GoTo eh
    Dim fl  As String
    Dim v   As Variant
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    For Each v In NamesInArg(pp_value_names)
        DeletePrivateProfileKey Section:=pp_section _
                               , Key:=v _
                               , Setting:=0 _
                               , Name:=fl
    Next v

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RemoveSections(ByVal pp_file As Variant, _
                 Optional ByVal pp_sections As String = vbNullString)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Removes the sections (pp_sections) - provided as: comma delimited string,
' Dictionary with name items, or Collection of names, in file (pp_file) -
' provided as full name string or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "RemoveSections"
    
    On Error GoTo eh
    Dim fl  As String
    Dim v   As Variant
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    For Each v In NamesInArg(pp_sections)
        DeletePrivateProfileSection Section:=v _
                                  , NoKey:=0 _
                                  , NoSetting:=0 _
                                  , Name:=fl
    Next v
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Function SelectFile( _
'            Optional ByVal sel_init_path As String = vbNullString, _
'            Optional ByVal sel_filters As String = "*.*", _
'            Optional ByVal sel_filter_name As String = "File", _
'            Optional ByVal sel_title As String = vbNullString, _
'            Optional ByRef sel_result As File) As Boolean
'' --------------------------------------------------------------
'' When a file had been selected TRUE is returned and the
'' selected file is returned as File object (sel_result).
'' --------------------------------------------------------------
'    Const PROC = "SelectFile"
'
'    On Error GoTo eh
'    Dim fDialog As FileDialog
'    Dim v       As Variant
'
'    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
'    With fDialog
'        .AllowMultiSelect = False
'        If sel_title = vbNullString _
'        Then .Title = "Select a(n) " & sel_filter_name _
'        Else .Title = sel_title
'        .InitialFileName = sel_init_path
'        .Filters.Clear
'        For Each v In Split(sel_filters, ",")
'            .Filters.Add sel_filter_name, v
'         Next v
'
'        If .Show = -1 Then
'            '~~ A fie had been selected
'           With New FileSystemObject
'            Set sel_result = .GetFile(fDialog.SelectedItems(1))
'            SelectFile = True
'           End With
'        End If
'        '~~ When no file had been selected the sel_result will be Nothing
'    End With
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function

Public Function Search(ByVal fs_root As String, _
              Optional ByVal fs_mask As String = "*", _
              Optional ByVal fs_in_subfolders As Boolean = True, _
              Optional ByVal fs_stop_after As Long = 100) As Collection
' ---------------------------------------------------------------------
' Returns a collection of all file names which meet the criteria:
' - in any subfolder of the root (fs_root)
' - meeting the wildcard comparison (fs_file_mask)
' ---------------------------------------------------------------------
    Const PROC = "Search"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim fo      As Folder
    Dim sfo     As Folder
    Dim fl      As File
    Dim queue   As New Collection

    Set Search = New Collection
    If Right(fs_root, 1) = "\" Then fs_root = Left(fs_root, Len(fs_root) - 1)
    If Not fso.FolderExists(fs_root) Then GoTo xt
    queue.Add fso.GetFolder(fs_root)

    Do While queue.Count > 0
        Set fo = queue(queue.Count)
        queue.Remove queue.Count ' dequeue the processed subfolder
        For Each sfo In fo.SubFolders
            queue.Add sfo ' enqueue (collect) all subfolders
        Next sfo
        For Each fl In fo.Files
            If VBA.Left$(fl.Name, 1) <> "~" _
            And fl.Name Like fs_mask _
            Then Search.Add fl
        Next fl
    Loop

xt: Set fso = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

                 
Public Function SectionExists(ByVal pp_file As Variant, _
                              ByVal pp_section As String) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Returns TRUE when the section (pp_section) exists in file (pp_file) -
' provided as full name string or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "SectionExists"
    
    Dim fl As String
    If PPFile(pp_file, ErrSrc(PROC), fl) <> vbNullString _
    Then SectionExists = mFile.SectionNames(fl).Exists(pp_section)
End Function

Public Function SectionNames( _
              Optional ByVal pp_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Returns a Dictionary of all section names in file (pp_file) - provided as
' full name string or as file object - in ascending order.
' ----------------------------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim asSections()    As String
    Dim dct             As New Dictionary
    Dim i               As Long
    Dim iLen            As Long
    Dim strBuffer       As String
    Dim fl              As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    If Len(mFile.Txt(fl)) = 0 Then GoTo xt
    Set SectionNames = New Dictionary
    
    Do While (iLen = Len(strBuffer) - 2) Or (iLen = 0)
        If strBuffer = vbNullString _
        Then strBuffer = Space$(256) _
        Else strBuffer = String(Len(strBuffer) * 2, 0)
        iLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), fl)
    Loop
    strBuffer = Left$(strBuffer, iLen)
    
    If Len(strBuffer) <> 0 Then
        i = 0
        asSections = Split(strBuffer, vbNullChar)
        For i = LBound(asSections) To UBound(asSections)
            If Len(asSections(i)) <> 0 _
            Then AddAscByKey add_dct:=SectionNames _
                           , add_key:=asSections(i) _
                           , add_item:=asSections(i)
        Next i
    End If
    
xt: Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SectionsCopy(ByVal pp_source As String, _
                        ByVal pp_target As String, _
               Optional ByVal pp_sections As Variant = Nothing, _
               Optional ByVal pp_merge As Boolean = False)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Copies sections from file (pp_source) to file (pp_target), when no section
' names (pp_sections) are provided all. By default (pp_merge) all sections
' are replaced.
' ----------------------------------------------------------------------------
    Const PROC = "SectionCopy"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim dct         As Dictionary
    Dim vName       As Variant
    Dim vSection    As Variant
    
    For Each vSection In NamesInArg(pp_sections)
        If Not pp_merge Then
            '~~ Section will be replaced
            mFile.RemoveSections pp_file:=pp_target _
                               , pp_sections:=vSection
        End If
        Set dct = mFile.Values(pp_file:=pp_source _
                            , pp_section:=vSection)
        For Each vName In dct
            mFile.Value(pp_file:=pp_target _
                      , pp_section:=vSection _
                      , pp_value_name:=vName) = dct(vName)
        Next vName
     Next vSection

xt: Set dct = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ShellRun(sCmd As String) As String
' ----------------------------------------------------------------------------
' Run a shell command, returning the output as a string.
' ----------------------------------------------------------------------------
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function

Public Function ValueExists(ByVal pp_file As Variant, _
                            ByVal pp_value_name As String, _
                            ByVal pp_section As String) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service. Returns TRUE when the named value
' (pp_value_name) exists in the provided section (pp_section) in the file
' (pp_file) - provided as full name string or file object. When no file is
' provided a file selection dialog is displayed. When finally there's still no
' file provided the function returns FALSE without notice.
' ----------------------------------------------------------------------------
    Const PROC = "ValueExists"
    
    Dim fl As String
    If PPFile(pp_file, ErrSrc(PROC), fl) <> vbNullString Then
        If mFile.SectionExists(pp_file:=fl, pp_section:=pp_section) Then
            ValueExists = mFile.ValueNames(fl, pp_section).Exists(pp_value_name)
        End If
    End If
End Function

Public Function ValueNames(ByVal pp_file As Variant, _
                           ByVal pp_section As String) As Dictionary
' ------------------------------------------------------------------------
' PrivateProfile file service. Returns a Dictionary of all value names,
' with the value name as key and the value as item, in ascending order by
' value name, in the provided section (pp_section) in file (pp_file) -
' provided either a full name string or as file object.
' ------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim dctNames    As New Dictionary
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim v           As Variant
    Dim sSection    As String
    Dim sName       As String
    Dim vNames      As Variant
    Dim fl          As String
    
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    
    Set vNames = NamesInArg(pp_section)
    If vNames.Count = 0 Then Set vNames = mFile.SectionNames(fl)
    
    For Each v In vNames
        sSection = v
        '~~> Retrieve the names for the provided section
        strBuffer = Space$(32767)
        lResult = GetPrivateProfileString(lpg_ApplicationName:=sSection _
                                        , lpg_KeyName:=vbNullString _
                                        , lpg_Default:=vbNullString _
                                        , lpg_ReturnedString:=strBuffer _
                                        , nSize:=Len(strBuffer) _
                                        , lpg_FileName:=fl _
                                         )
        sNames = Left$(strBuffer, lResult)
    
        If sNames <> vbNullString Then                                         ' If there were any names
            asNames = Split(sNames, vbNullChar)                      ' have them split into an array
            For i = LBound(asNames) To UBound(asNames)
                sName = asNames(i)
                If Len(sName) <> 0 Then
                    If Not dctNames.Exists(sName) _
                    Then AddAscByKey add_dct:=dctNames _
                                   , add_key:=sName _
                                   , add_item:=mFile.Value(pp_file:=fl _
                                                         , pp_section:=sSection _
                                                         , pp_value_name:=sName)
                End If
            Next i
        End If
    Next v
        
    Set ValueNames = dctNames

xt: Set dctNames = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Values(ByVal pp_file As Variant, _
                       ByVal pp_section As String) As Dictionary
' ------------------------------------------------------------------------
' PrivateProfile file service. Returns all values in section (pp_section)
' in file (pp_file) as Dictionary, with the value name as key and the
' value as item, with the value names in asscending order. When the
' provided section does not exist the returned Dictionary is empty.
' ------------------------------------------------------------------------
    Const PROC = "Values"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim dctNames    As New Dictionary
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim sName       As String
    Dim fl          As String
    
    Set Values = New Dictionary ' may remain empty though
    If PPFile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    If Not SectionExists(pp_file, pp_section) Then GoTo xt
    
    '~~> Retrieve the names in the provided section
    strBuffer = Space$(32767)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=pp_section _
                                    , lpg_KeyName:=vbNullString _
                                    , lpg_Default:=vbNullString _
                                    , lpg_ReturnedString:=strBuffer _
                                    , nSize:=Len(strBuffer) _
                                    , lpg_FileName:=fl _
                                     )
    sNames = Left$(strBuffer, lResult)

    If sNames <> vbNullString Then                                         ' If there were any names
        asNames = Split(sNames, vbNullChar)                      ' have them split into an array
        For i = LBound(asNames) To UBound(asNames)
            sName = asNames(i)
            If Len(sName) <> 0 Then
                If Not dctNames.Exists(sName) _
                Then AddAscByKey add_dct:=Values _
                               , add_key:=sName _
                               , add_item:=mFile.Value(pp_file:=fl _
                                                     , pp_section:=pp_section _
                                                     , pp_value_name:=sName)
            End If
        Next i
    End If

xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

