Attribute VB_Name = "mFso"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard  Module mFso: Common services regarding files system objects
' ====================== (files and folders). The module makes extensive use
' of the FileSystemObject by extending and supplementing its services. The
' componet works autonomous and does not require any other components.
'
' Public properties and services:
' -------------------------------
' Exists                Universal existence function. Returns TRUE when the
'                       folder, file, section ([xxxx]), name exists
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
' FileTxt           r/w Returns the content of a text file as text string
'                       or writes a string to a text file - optionally
'                       appended
' RenameSubFolders      Renames all folders and sub-folders in a provided
'                       path from a given old to a given new name.
' SortCutTargetPath r   Returns a provided shortcut's target path
'                       Let Creates a shortcut at a provided location with a
'                       provided path
'
' Public PrivateProfile file (PP) properties and services:
' ---------------------------------------------------
' Exists            see above
' PPsectionExists   Returns TRUE when a given section exists in a given PP
' PPsectionNames    Returns a Dictionary of all section names [...] in a PP.
' PPsections    r/w Returns named (or if no names are provideds all) sections
'                   as Dictionary with the section name as the key and the
'                   Values Dictionary as item or writes from the returned
'                   Dictionary all sections provided with all names to a PP.
' PPremoveSections  Removes the sections provided via by their given name.
'                   Sections not existing are ignored.
' PPreorg           Reorganizes all sections and their value-names in a PP
'                   by ordering sections and names in ascending sequence.
' PPvalue       r/w Reads from or writes to a name from/to a PP.
' PPvalueNameExists Returns TRUE when a value-name exists in a PP under a
'                   given section.
' PPvalueNameRename Function replaces an old value name with a new one
'                   either in a specific section or in all sections when no
'                   specific section is provided. Optionally not reorgs the
'                   file, returns True when at least one name has been
'                   replaced.
' PPvalueNames      Returns a Dictionary of all value-names within given
'                   sections in a PP with the value-name and the section name
'                   as key (<name>[section]) and the value as item, the names
'                   in ascending order in a Dictionary. Section names may be
'                   provided as a comma delimited string, a Dictionary or
'                   Collection. Non existing sections are ignored.
' PPvalues          Returns the value-names and values of a given section
'                   in a PP as Dictionary with the value-name as the key
'                   (in ascending order) and the value as item.
'
' Requires References to:
' -----------------------
' Microsoft Scripting Runtine
' Windows Script Host Object Model
'
' Uses no other components. Will use optionally mErH, fMsg/mMsg when installed and
' activated (Cond. Comp. Args. `ErHComp = 1 : MsgComp = 1`).
'
' W. Rauschenberger, Berlin June 2022
' See also https://github.com/warbe-maker/VBA-File-System-Objects.
' ----------------------------------------------------------------------------
Private Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-File-System-Objects"
Private fso                     As New FileSystemObject

#If MsgComp = 0 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' approach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Conditional Compile Argument MsgComp = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

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
Private Const WIN_MAX = 3            'Open Maximized
Private Const WIN_MIN = 2            'Open Minimized

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16


Public Property Let FileArry(Optional ByVal f_file As String, _
                             Optional ByVal f_excl_empty_lines As Boolean = False, _
                             Optional ByRef f_split As String = vbCrLf, _
                             Optional ByVal f_append As Boolean = False, _
                                      ByVal f_ar As Variant)
' ----------------------------------------------------------------------------
' Writes array (f_ar) to file (f_file) whereby the array is joined to a text
' string using the line break string (f_split) which defaults to vbCrLf and
' is optionally returned by Arry-Get.
' ----------------------------------------------------------------------------
                    
    mFso.FileTxt(f_file:=f_file _
               , f_append:=f_append _
               , f_split:=f_split _
                ) = Join(f_ar, f_split)
             
End Property

Public Property Get FileArry(Optional ByVal f_file As String, _
                             Optional ByVal f_excl_empty_lines As Boolean = False, _
                             Optional ByRef f_split As String, _
                             Optional ByVal f_append As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Returns the content of the file (f_file) - a files full name - as array,
' with the used line break string returned in (f_split).
' ----------------------------------------------------------------------------
    Const PROC  As String = "FileArry"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim a       As Variant
    Dim a1()    As String
    Dim sSplit  As String
    Dim sFile   As String
    Dim i       As Long
    Dim j       As Long
    Dim v       As Variant
    
    If Not fso.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named '" & f_file & "' does not exist!"
    
    '~~ Unload file to a string
    sFile = mFso.FileTxt(f_file:=f_file _
                    , f_split:=sSplit _
                     )
    If sFile = vbNullString Then GoTo xt
    a = Split(sFile, sSplit)
    
    If Not f_excl_empty_lines Then
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
    
xt: FileArry = a1
    f_split = sSplit
    Set cll = Nothing
    Set fso = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get FileDict(ByVal f_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Returns the content of the file (f_file) - which may be provided as file
' object or full file name - as Dictionary by considering any kind of line
' break characters.
' ----------------------------------------------------------------------------
    Const PROC  As String = "FileDict-Get"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim a       As Variant
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim flo     As File
    Dim sFile   As String
    Dim i       As Long
    
    Select Case TypeName(f_file)
        Case "File"
            If Not fso.FileExists(f_file.Path) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (f_file) does not exist!"
            Set flo = f_file
        Case "String"
            If Not fso.FileExists(f_file) _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided file string (f_file) is not an existing file!"
            Set flo = fso.GetFile(f_file)
    End Select
    
    '~~ Unload file into a test stream
    With fso
        Set ts = .OpenTextFile(flo.Path, 1)
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
        
xt: Set FileDict = dct
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get FileTemp(Optional ByVal f_path As String = vbNullString, _
                             Optional ByVal f_extension As String = ".tmp") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(fso.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    FileTemp = sTemp
    
    Set fso = Nothing
End Property

Public Property Get FileTxt(Optional ByVal f_file As Variant, _
                            Optional ByVal f_append As Boolean = True, _
                            Optional ByRef f_split As String) As String
' ----------------------------------------------------------------------------
' Returns the text file's (f_file) content as string with VBA.Split() string
' in (f_split). When the file doesn't exist a vbNullString is returned.
' Note: f_append is not used but specified to comply with Property Let.
' ----------------------------------------------------------------------------
    Const PROC = "FileTxt-Get"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim s       As String
    Dim sFl As String
   
    f_split = f_split  ' not used! for declaration compliance and dead code check only
    f_append = f_append ' not used! for declaration compliance and dead code check only
    
    With fso
        If TypeName(f_file) = "File" Then
            sFl = f_file.Path
        Else
            '~~ f_file is regarded a file's full name, created if not existing
            sFl = f_file
            If Not .FileExists(sFl) Then GoTo xt
        End If
        Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForReading)
    End With
    
    If Not ts.AtEndOfStream Then
        s = ts.ReadAll
        f_split = SplitStr(s)
        If VBA.Right$(s, 2) = vbCrLf Then
            s = VBA.Left$(s, Len(s) - 2)
        End If
    Else
        FileTxt = vbNullString
    End If
    If FileTxt = vbCrLf Then FileTxt = vbNullString Else FileTxt = s

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let FileTxt(Optional ByVal f_file As Variant, _
                            Optional ByVal f_append As Boolean = True, _
                            Optional ByRef f_split As String, _
                                     ByVal f_string As String)
' ----------------------------------------------------------------------------
' Writes the string (f_string) into the file (f_file) which might be a file
' object or a file's full name.
' Note: f_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "FileTxt-Let"
    
    On Error GoTo eh
    Dim ts  As TextStream
    Dim sFl As String
   
    f_split = f_split ' not used! just for coincidence with Get
    With fso
        If TypeName(f_file) = "File" Then
            sFl = f_file.Path
        Else
            '~~ f_file is regarded a file's full name, created if not existing
            sFl = f_file
            If Not .FileExists(sFl) Then .CreateTextFile sFl
        End If
        
        If f_append _
        Then Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForWriting)
    End With
    
    ts.WriteLine f_string

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property
             
Public Property Get PPsections(Optional ByVal pp_file As Variant, _
                               Optional ByVal pp_sections As Variant = vbNullString, _
                               Optional ByRef pp_file_result As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with named sections (pp_sections) in file (pp_file) -
' provided as full name string or as file object - whereby each section is a
' Dictionary with the section name as the key - in ascending
' order - and a Dictionary of the section's values as item with the value name
' as key and the value as item.
' ----------------------------------------------------------------------------
    Const PROC = "PPsections-Get"
    
    On Error GoTo eh
    Dim vName       As Variant
    Dim fl          As String
    Dim dct         As New Dictionary
    Dim dctValues   As Dictionary
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
        
    If pp_sections = vbNullString Then
        '~~ Return all sections
        For Each vName In mFso.PPsectionNames(fl)
            If Not dct.Exists(vName) Then
                Set dctValues = mFso.PPvalues(pp_file:=fl, pp_section:=vName)
                dct.Add vName, dctValues
            Else
                Debug.Print "Duplicate section name '" & vName & "' ignored!"
            End If
        Next vName
    Else
        '~~ Return named sections
        For Each vName In NamesInArg(pp_sections)
            If mFso.PPsectionExists(pp_file, vName) Then
                If Not dct.Exists(vName) Then
                    Set dctValues = mFso.PPvalues(pp_file, vName)
                    dct.Add vName, dctValues
                Else
                    Debug.Print "Duplicate section name '" & vName & "' ignored!"
                End If
            End If
        Next vName
    End If

xt: Set PPsections = KeySort(dct)
    Set dct = Nothing
    Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let PPsections(Optional ByVal pp_file As Variant, _
                               Optional ByVal pp_sections As Variant, _
                               Optional ByRef pp_file_result As String, _
                                        ByVal pp_dct As Dictionary)
' ------------------------------------------------------------------------
' Writes the sections provided as Dictionary (pp_dct) to file (pp_file) -
' provided as full name string or as file object. Existing sections/values
' are overwritten new sections/values are added.
' ------------------------------------------------------------------------
    Const PROC = "PPsections-Get"
    
    On Error GoTo eh
    Dim vN          As Variant
    Dim vS          As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim sName       As String
    Dim fl          As String
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    For Each vS In pp_dct
        sSection = vS
        Set dctValues = pp_dct(vS)
        For Each vN In dctValues
            sName = vN
            mFso.PPvalue(pp_file:=fl _
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

Public Property Get PPvalue(Optional ByVal pp_file As Variant, _
                            Optional ByVal pp_section As String, _
                            Optional ByVal pp_value_name As String, _
                            Optional ByRef pp_file_result As String) As Variant
' ----------------------------------------------------------------------------
' Reads a value with a specific name from a section from a file organized:
'[section]
'<value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC  As String = "PPvalueGet"
    
    On Error GoTo eh
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant
    Dim fl      As String
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=pp_section _
                                    , lpg_KeyName:=pp_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=fl _
                                     )
    vValue = Left$(sRetVal, lResult)
    PPvalue = vValue
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let PPvalue(Optional ByVal pp_file As Variant, _
                            Optional ByVal pp_section As String, _
                            Optional ByVal pp_value_name As String, _
                            Optional ByRef pp_file_result As String, _
                                     ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value under a given name (pp_value_name) into a section (pp_section)
' in a file (pp_file) organized:
' [section]
' <value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC = "PPvalueLet"
        
    On Error GoTo eh
    Dim lChars  As Long
    Dim sValue  As String
    Dim fl      As String
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
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

Public Function KeySort(ByRef s_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (s_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim Txt     As String
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
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
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
    ErrSrc = "mFso." & sProc
End Function

Public Function Exists(Optional ByVal x_folder As String = vbNullString, _
                       Optional ByVal x_file As String = vbNullString, _
                       Optional ByVal x_section As String = vbNullString, _
                       Optional ByVal x_value_name As String = vbNullString, _
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
    Dim sTest           As String
    Dim sFile           As String
    Dim fo              As Folder   ' Folder
    Dim sfo             As Folder   ' Sub-Folder
    Dim fl              As File
    Dim queue           As New Collection
    Dim FolderExists    As Boolean
    Dim FileExists      As Boolean
    Dim SectionExists   As Boolean
    Dim ValueNameExists As Boolean
    Dim sFolder         As String
    
    Set x_result_files = New Collection

    With fso
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
            sFile = x_file
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
            If x_section = vbNullString Then
                '~~ When no section is provided, that's it
                Exists = True
                GoTo xt
            End If
        End If
        
        '~~ At this point either a provided folder together with a provided file matched exactly one existing file
        '~~ or a specified file's existence had been proved
        If x_section <> vbNullString Then
            If Not mFso.PPsectionExists(pp_file:=x_file, pp_section:=x_section) Then GoTo xt
            If x_value_name = vbNullString Then
                '~~ When no x_value_name is provided, that's it
                Exists = True
            Else
                Exists = mFso.PPvalueExists(pp_file:=x_file, pp_section:=x_section, pp_value_name:=x_value_name)
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
    
    If Not fso.FileExists(f_file1) _
    Then Err.Raise Number:=AppErr(2) _
                 , source:=ErrSrc(PROC) _
                 , Description:="The file """ & f_file1 & """ does not exist!"
    
    If Not fso.FileExists(f_file2) _
    Then Err.Raise Number:=AppErr(3) _
                 , source:=ErrSrc(PROC) _
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
    Then Err.Raise Number:=AppErr(1), source:=ErrSrc(PROC), Description:= _
                   "WinMerge is obligatory for the Compare service of this module " & _
                   "but not installed!" & vbLf & vbLf & _
                   "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(f_file_left) _
    Then Err.Raise Number:=AppErr(2), source:=ErrSrc(PROC), Description:= _
                   "The file """ & f_file_left & """ does not exist!"
    
    If Not fso.FileExists(f_file_right) _
    Then Err.Raise Number:=AppErr(3), source:=ErrSrc(PROC), Description:= _
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

    With fso
        If TypeName(v) = "File" Then
            .DeleteFile v.Path
        ElseIf TypeName(v) = "String" Then
            If .FileExists(v) Then .DeleteFile v
        End If
    End With
    
End Sub

Public Function FileDiffers(ByVal f_file1 As File, _
                            ByVal f_file2 As File, _
                   Optional ByVal f_stop_after As Long = 0, _
                   Optional ByVal f_ignore_empty_records As Boolean = False, _
                   Optional ByVal f_compare As VbCompareMethod = vbTextCompare) As Dictionary
' -----------------------------------------------------------------------------
' Returns TRUE when the content of file (f_file1) differs from the content in
' file (f_file2). The comparison stops after (f_stop_after) detected
' differences. The detected different lines are optionally returned (vResult).
' ------------------------------------------------------------------------------
    Const PROC = "FileDiffers"
    
    On Error GoTo eh
    Dim a1          As Variant
    Dim a2          As Variant
    Dim dct1        As New Dictionary
    Dim dct2        As New Dictionary
    Dim dctDif      As New Dictionary
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
    
    sFile1 = f_file1.Path
    sFile2 = f_file2.Path
    
    s1 = mFso.FileTxt(f_file:=sFile1, f_split:=sSplit)
    If f_ignore_empty_records Then
        '~~ Eliminate empty records
        sTest1 = VBA.Replace$(s1, sSplit & sSplit, sSplit)
    End If
    
    s2 = mFso.FileTxt(f_file:=sFile2, f_split:=sSplit)
    If f_ignore_empty_records Then
        '~~ Eliminate empty records
        sTest2 = VBA.Replace$(s2, sSplit & sSplit, sSplit)
    End If
    
    If VBA.StrComp(s1, s2, f_compare) = 0 Then GoTo xt

     
    a1 = Split(s1, sSplit)
    For i = LBound(a1) To UBound(a1)
        dctF1.Add i + 1, a1(i)
        If f_ignore_empty_records Then
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
        If f_ignore_empty_records Then
            If VBA.Trim$(a2(i)) <> vbNullString Then
                dct2.Add i + 1, a2(i)
            End If
        Else
            dct2.Add i + 1, a2(i)
        End If
    Next i
    If VBA.StrComp(Join(dct1.Items(), sSplit), Join(dct2.Items(), sSplit), f_compare) = 0 Then GoTo xt
    
    '~~ Get and detect the difference by comparing the items one by one
    '~~ and optaining the line number from the Dictionary when different
    If dct1.Count <= dct2.Count Then
        For Each v In dct1 ' v - 1 = array index
            If VBA.StrComp(a1(v - 1), a2(v - 1), f_compare) <> 0 Then
                lDiffLine = v
                dctDif.Add lDiffLine, DiffItem(d_line:=lDiffLine _
                                             , d_file_left:=sFile1 _
                                             , d_file_right:=sFile2 _
                                             , d_line_left:=a1(v - 1) _
                                             , d_line_right:=a2(v - 1) _
                                              )
                If f_stop_after > 0 And dctDif.Count >= f_stop_after Then GoTo xt
            End If
        Next v
        
        For i = dct1.Count + 1 To dct2.Count
            lDiffLine = dct2.Keys()(i - 1)
            dctDif.Add lDiffLine, DiffItem(d_line:=lDiffLine _
                                         , d_file_left:=sFile1 _
                                         , d_file_right:=sFile2 _
                                         , d_line_right:=a2(i - 1) _
                                          )
        Next i

    ElseIf dct2.Count < dct1.Count Then
        For Each v In dct2 ' v - 1 = array index
            If VBA.StrComp(a1(v - 1), a2(v - 1), f_compare) <> 0 Then
                lDiffLine = v
                dctDif.Add lDiffLine, DiffItem(d_line:=lDiffLine _
                                             , d_file_left:=sFile1 _
                                             , d_file_right:=sFile2 _
                                             , d_line_left:=a1(v - 1) _
                                             , d_line_right:=a2(v - 1) _
                                              )
                If f_stop_after > 0 And dctDif.Count >= f_stop_after Then GoTo xt
            End If
        Next v
        For i = dct2.Count + 1 To dct1.Count
            lDiffLine = dct1.Keys()(i - 1)
            dctDif.Add lDiffLine, DiffItem(d_line:=lDiffLine _
                                         , d_file_left:=sFile1 _
                                         , d_file_right:=sFile2 _
                                         , d_line_left:=a1(i - 1) _
                                          )
        Next i
    End If
        
xt: Set FileDiffers = dctDif
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

Public Function FileExtension(ByVal fe_file As Variant)

    With fso
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
    Dim sFile   As String
    Dim a()     As String
    Dim i       As Long
    Dim v       As Variant
    Dim FolderIsValidName   As Boolean
    
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
                FileIsValidName = Err.Number = 0
                If Not FileIsValidName Then GoTo xt
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

Public Sub FilePathSplit(ByRef f_path As String, _
                Optional ByRef f_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Returns the provided file path string (f_path) split in path and file name
' provided the argument is either an existing folder's path or an existing
' file's path. When sp_path is a folder only string for sp_file a vbNullString
' is returned, when sp_path only contains a file name, sp_folder = vbNullString
' is returned else sp_file = the file's name
' When the provided path is neither an existing folder's path nor an existing
' file's full name both arguments are returned a vbNullString.
' Note: For a non existing folder or file it is impossible to decide whether
'       the string ends with the name of a folder or the name of a file
'       both have exactly the the same restrictions regardin the name. I.e.
'       a file name xxxx.dat may as well be the name of a folder.
' ----------------------------------------------------------------------------
    Const PROC = "FilePathSplit"
    
    On Error GoTo eh
    Dim aPath() As String
    Dim i       As Long
    Dim sFolder As String
    Dim sFile   As String
    
    With fso
        Select Case True
            Case Not .FileExists(f_path) And Not .FolderExists(f_path)
                Debug.Print ErrSrc(PROC) & ": Neither a folder nor a file exists with the provided string '" & f_path & "'!"
                f_path = vbNullString
                f_file = vbNullString
            Case .FileExists(f_path)
                f_file = .GetFileName(f_path)
                f_path = Replace(f_path, f_file, vbNullString)
            Case .FolderExists(f_path)
                f_file = vbNullString
        End Select
    End With
            
xt: Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
    Dim fd  As FileDialog
    Dim v   As Variant
    
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
            FilePicked = True
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

Public Function FilesSearch(ByVal fs_root As String, _
                   Optional ByVal fs_mask As String = "*", _
                   Optional ByVal fs_in_subfolders As Boolean = True, _
                   Optional ByVal fs_stop_after As Long = 100) As Collection
' ---------------------------------------------------------------------
' Returns a collection of all file names which meet the criteria:
' - in any subfolder of the root (fs_root)
' - meeting the wildcard comparison (fs_file_mask)
' ---------------------------------------------------------------------
    Const PROC = "FilesSearch"
    
    On Error GoTo eh
    Dim fo      As Folder
    Dim sfo     As Folder
    Dim fl      As File
    Dim queue   As New Collection

    Set FilesSearch = New Collection
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
            Then FilesSearch.Add fl
        Next fl
    Loop

xt: Set fso = Nothing
    Exit Function

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
    Dim i   As Long
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
        With fso
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
    
xt: Set fso = Nothing
    Exit Function
 
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
    Static cll      As Collection
    Static queue    As Collection   ' FiFo queue for folders with sub-folders
    Static Stack    As Collection   ' LiFo stack for recursive calls
    Static foStart  As Folder
    Dim aFolders()  As Variant
    Dim fl          As File
    Dim flStart     As Folder
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
            If Not fso.FolderExists(fo_spec) Then
                fo_result = fo_spec
                GoTo xt
            End If
            Set fo1 = fso.GetFolder(fo_spec)
        Else
            Application.DisplayAlerts = False
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "Please select the desired folder!"
                .InitialFileName = CurDir
                .AllowMultiSelect = False
                If .Show <> -1 Then GoTo xt
                Set fo1 = fso.GetFolder(.SelectedItems(1))
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
        Debug.Print "Remove " & queue(1).Path & " from stack (" & queue.Count - 1 & " still in queue)"
        queue.Remove 1
    End If
    If queue.Count > 0 Then
        Folders queue(1).Path ' recursive call for each folder with subfolders
    End If
    
xt: Set fso = Nothing
    If Stack.Count > 0 Then
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
                Set fo1 = fso.GetFolder(aFolders(i))
                cll.Add fo1
            Next i
        End If
        Set Folders = cll
        If Not foStart Is Nothing Then fo_result = foStart.Path
    End If
    Set cll = Nothing
    
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

Private Function NamesInArg(Optional ByVal v As Variant = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns the provided argument (v) as Collection of string items whereby (v)
' may be: not provided, a comma delimited string, a Dictionary, or a
' Collection of string items.
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

Private Function PPfile(ByVal pp_file As Variant, _
                        ByVal pp_service As String, _
                        ByRef pp_file_name As String) As String
' ----------------------------------------------------------------------------
' PrivateProfile file function. Returns the PrivateProfile file (pp_file) as
' full file name string (pp_file_name). When no file (pp_file) is provided,
' i.e. is a vbNullString, a file selection dialog is evoked to select one.
' When finally there's still no file provided a vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "PPfile"
    
    On Error GoTo eh
    Dim fl  As File
    
    If VarType(pp_file) = vbObject Then
        If pp_file Is Nothing Then
            If Not mFso.FilePicked(p_title:="Select a PrivateProfile file not provided for the service '" & pp_service & "'" _
                              , p_init_path:=ThisWorkbook.Path _
                              , p_filters:="Config files, *.cfg; Init files, *.ini; Application data files, *.dat; All files, *.*" _
                              , p_file:=fl _
                               ) _
            Then GoTo xt
            pp_file_name = fl.Path
            PPfile = fl.Path
        ElseIf TypeName(pp_file) = "File" Then
            pp_file_name = pp_file.Path
            PPfile = pp_file.Path
            GoTo xt
        End If
    ElseIf VarType(pp_file) = vbString Then
        If pp_file = vbNullString Then
            If Not mFso.FilePicked(p_title:="Select a PrivateProfile file" _
                              , p_init_path:=ThisWorkbook.Path _
                              , p_filters:="Config files, *.cfg; Init files, *.ini; Application data files, *.dat; All files, *.*" _
                              , p_file:=fl _
                               ) _
            Then GoTo xt
            pp_file_name = fl.Path
            PPfile = fl.Path
        Else
            pp_file_name = pp_file
            PPfile = pp_file
        End If
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function PPremoveNames(ByVal pp_file As Variant, _
                              ByVal pp_section As String, _
                              ByVal pp_value_names As Variant, _
                     Optional ByRef pp_file_result As String) As Boolean
' -----------------------------------------------------------------------------
' PrivateProfile file service. Removes from a PrivateProfile file's (pp_file)
' section (pp_section) a name (pp_value_names) provided either as a
' comma delimited string, or as one string only.
' When no file (pp_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notice.
' When the name existed and has been removed, the function returns TRUE.
' -----------------------------------------------------------------------------
    Const PROC = "PPremoveNames"
    
    On Error GoTo eh
    Dim fl  As String
    Dim v   As Variant
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    For Each v In NamesInArg(pp_value_names)
        If mFso.PPvalueExists(pp_file, v, pp_section) Then
            DeletePrivateProfileKey Section:=pp_section _
                                  , Key:=v _
                                  , Setting:=0 _
                                  , Name:=fl
            PPremoveNames = True
        End If
    Next v

xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub PPremoveSections(ByVal pp_file As Variant, _
                   Optional ByVal pp_sections As String = vbNullString, _
                   Optional ByRef pp_file_result As String)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Removes the sections (pp_sections) - provided as: comma delimited string,
' Dictionary with name items, or Collection of names, in file (pp_file) -
' provided as full name string or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "PPremoveSections"
    
    On Error GoTo eh
    Dim fl  As String
    Dim v   As Variant
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
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

Public Sub PPreorg(Optional ByVal pp_file As Variant = Nothing, _
                   Optional ByRef pp_file_result As String)
' ----------------------------------------------------------------------------
' Private-Profile file service. Re-organizes all sections and all value-names
' therein in ascending order. When no file (pp_file) is provided - either as
' full name or as file object a file selection dialog is displayed. When
' finally there's still no file provided the service ends without notice. The
' processed file is returned (pp_file_result).
' Restriction: The file must not contain any comment lines other than at the
'              top of the very first section.
' ----------------------------------------------------------------------------
    Const PROC = "PPreorg"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim Section     As String
    Dim dctSections As Dictionary
    Dim dctValues   As Dictionary
    Dim vValue      As Variant
    Dim fl          As String
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    Set dctSections = PPsections(fl)
    For Each vSection In dctSections
        Section = vSection
        Set dctValues = dctSections(vSection)
        
        mFso.PPremoveSections pp_file:=fl, pp_sections:=Section
        For Each vValue In dctValues
            mFso.PPvalue(pp_file:=fl, pp_section:=Section, pp_value_name:=vValue) = dctValues(vValue)
        Next vValue
    Next vSection

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

                 
Public Function PPsectionExists(ByVal pp_file As Variant, _
                                ByVal pp_section As String, _
                       Optional ByRef pp_file_result As String) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Returns TRUE when the section (pp_section) exists in file (pp_file) -
' provided as full name string or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "PPsectionExists"
    
    Dim fl As String
    If PPfile(pp_file, ErrSrc(PROC), fl) <> vbNullString Then
        pp_file_result = fl
        PPsectionExists = mFso.PPsectionNames(fl).Exists(pp_section)
    End If
    
End Function

Public Function PPsectionNames(Optional ByVal pp_file As Variant, _
                              Optional ByRef pp_file_result As String) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Returns a Dictionary of all section names in file (pp_file) - provided as
' full name string or as file object - in ascending order.
' ----------------------------------------------------------------------------
    Const PROC = "PPsectionNames"
    
    On Error GoTo eh
    Dim asSections()    As String
    Dim dct             As New Dictionary
    Dim fl              As String
    Dim i               As Long
    Dim iLen            As Long
    Dim strBuffer       As String
    
    Set PPsectionNames = New Dictionary
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    If Len(mFso.FileTxt(fl)) = 0 Then GoTo xt
    
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
            If asSections(i) <> vbNullString Then
               dct.Add asSections(i), asSections(i)
            End If
        Next i
    End If
    
xt: Set PPsectionNames = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub PPsectionsCopy(ByVal pp_source As String, _
                          ByVal pp_target As String, _
                 Optional ByVal pp_sections As Variant = Nothing, _
                 Optional ByVal pp_merge As Boolean = False)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Copies sections from file (pp_source) to file (pp_target), when no section
' names (pp_sections) are provided all. By default (pp_merge) all sections
' are replaced.
' ----------------------------------------------------------------------------
    Const PROC = "PPsectionCopy"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim vName       As Variant
    Dim vSection    As Variant
    
    For Each vSection In NamesInArg(pp_sections)
        If Not pp_merge Then
            '~~ Section will be replaced
            mFso.PPremoveSections pp_file:=pp_target _
                               , pp_sections:=vSection
        End If
        Set dct = mFso.PPvalues(pp_file:=pp_source _
                            , pp_section:=vSection)
        For Each vName In dct
            mFso.PPvalue(pp_file:=pp_target _
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

Public Function PPvalueExists(ByVal pp_file As Variant, _
                              ByVal pp_value_name As String, _
                              ByVal pp_section As String, _
                     Optional ByRef pp_file_result As String) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service. Returns TRUE when the named value
' (pp_value_name) exists in the provided section (pp_section) in the file
' (pp_file) - provided as full name string or file object. When no file is
' provided a file selection dialog is displayed. When finally there's still no
' file provided the function returns FALSE without notice.
' ----------------------------------------------------------------------------
    Const PROC = "PPvalueExists"
    
    Dim fl As String
    If PPfile(pp_file, ErrSrc(PROC), fl) <> vbNullString Then
        pp_file_result = fl
        If mFso.PPsectionExists(pp_file:=fl, pp_section:=pp_section) Then
            PPvalueExists = mFso.PPvalueNames(fl, pp_section).Exists(pp_value_name)
        End If
    End If
End Function

Public Function PPvalueNameRename(ByVal pp_name_old As String, _
                                  ByVal pp_name_new As String, _
                                  ByVal pp_file As Variant, _
                         Optional ByVal pp_section As String = vbNullString, _
                         Optional ByRef pp_reorg As Boolean = True) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service. Rename an old value name (pp_name_old) to a new
' name (pp_nmae_new) either in a specific section (pp_section) or in all
' sections when no specific section is provided. Optionally not reorgs the
' file when (pp_reorg) = False (defaults to True).
' ----------------------------------------------------------------------------
    Const PROC = "PPvalueNameRename"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim sSect       As String
    Dim vValue      As Variant
    Dim dctSects    As Dictionary
    Dim vSect       As Variant
    
    If pp_name_old = vbNullString Or pp_name_new = vbNullString Then GoTo xt
    
    Set dct = PPsections(pp_file)
    
    For Each v In dct
        If pp_section <> vbNullString Then
            If v = pp_section Then
                '~~> Rename Name just within one specific section
                If mFso.PPvalueExists(pp_file, pp_name_old, pp_section) Then
                    vValue = PPvalue(pp_file, pp_section, pp_name_old)
                    PPremoveNames pp_file, pp_section, pp_name_old
                    If vValue <> vbNullString _
                    Then PPvalue(pp_file, pp_section, pp_name_new) = vValue
                    PPvalueNameRename = True
                End If
                GoTo xt
            End If
        Else
            Set dctSects = PPsections(pp_file)
            For Each vSect In dctSects
            '~~> Rename Name in all sections
                If mFso.PPvalueExists(pp_file, pp_name_old, pp_section) Then
                    vValue = PPvalue(pp_file, vSect, pp_name_old)
                    PPremoveNames pp_file, vSect, pp_name_old
                    If vValue <> vbNullString Then PPvalue(pp_file, vSect, pp_name_new) = vValue
                    PPvalueNameRename = True
                End If
            Next vSect
        End If
    Next v
    
xt: Set dct = Nothing
    If pp_reorg Then PPreorg pp_file
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function PPvalueNames(ByVal pp_file As Variant, _
                           ByVal pp_section As String, _
                  Optional ByRef pp_file_result As String) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile file service. Returns a Dictionary of all value names, with
' the value name as key and the value as item, in ascending order by value
' name, in the provided section (pp_section) in file (pp_file) - provided
' either a full name string or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "PPvalueNames"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim dct         As New Dictionary
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim v           As Variant
    Dim sSection    As String
    Dim sName       As String
    Dim vNames      As Variant
    Dim fl          As String
    
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    Set vNames = NamesInArg(pp_section)
    If vNames.Count = 0 Then Set vNames = mFso.PPsectionNames(fl)
    
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
                    If Not dct.Exists(sName) _
                    Then dct.Add sName, mFso.PPvalue(pp_file:=fl _
                                                        , pp_section:=sSection _
                                                        , pp_value_name:=sName)
                End If
            Next i
        End If
    Next v
        
    Set PPvalueNames = KeySort(dct)

xt: Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function PPvalues(ByVal pp_file As Variant, _
                         ByVal pp_section As String, _
                Optional ByRef pp_file_result As String) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile file service. Returns all values in section (pp_section) in
' file (pp_file) as Dictionary, with the value name as key and the value as
' item, with the value names in asscending order. When the provided section
' does not exist the returned Dictionary is empty.
' ----------------------------------------------------------------------------
    Const PROC = "PPvalues"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim sName       As String
    Dim fl          As String
    Dim dct         As New Dictionary
    
    Set PPvalues = New Dictionary ' may remain empty though
    If PPfile(pp_file, ErrSrc(PROC), fl) = vbNullString Then GoTo xt
    pp_file_result = fl
    
    If Not PPsectionExists(pp_file, pp_section) Then GoTo xt
    
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
                If Not dct.Exists(sName) Then
                    dct.Add sName, mFso.PPvalue(pp_file:=fl _
                                              , pp_section:=pp_section _
                                              , pp_value_name:=sName)
                Else
                    Debug.Print "Duplicate Name '" & sName & "' in section '" & pp_section & "' ignored!"
                End If
            End If
        Next i
    End If

xt: Set PPvalues = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Public Sub RenameSubFolders(ByVal rsf_path As String, _
                            ByVal rsf_old_name As String, _
                            ByVal rsf_new_name As String, _
                   Optional ByRef rsf_renamed As Collection)
' ----------------------------------------------------------------------------
' Rename any sub-folder in the provided path (rsf_path) named (rsf_old_name)
' to (rsf_new_name). Any subfolders in renamed folders are ignored. Returns
' all renamed folders as Collection (rsf_renamed).
' ----------------------------------------------------------------------------
    Dim fld         As Folder
    Dim fldSub      As Folder
    Dim cllQueue    As New Collection
    Dim cll         As New Collection
    
    cllQueue.Add fso.GetFolder(rsf_path)

    Do While cllQueue.Count > 0
        Set fld = cllQueue(1)
        cllQueue.Remove 1 'dequeue
        If fld.Name = rsf_old_name Then
            fld.Name = rsf_new_name
            cll.Add fld
            Debug.Print "Renamed: " & fld.Path
        Else
            For Each fldSub In fld.SubFolders
                cllQueue.Add fldSub 'enqueue
            Next fldSub
        End If
    Loop

xt: Set rsf_renamed = cll
    Set cll = Nothing
    Set fso = Nothing
    Exit Sub

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

