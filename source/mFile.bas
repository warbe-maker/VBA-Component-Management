Attribute VB_Name = "mFile"
Option Explicit
Option Compare Text
Option Private Module
' ------------------------------------------------------------------------
' Standard  Module mFile
'           Common methods and functions regarding file objects.
'
' Public services:
' - Exists          Returns TRUE when the provided file exists
' - Compare         Displays the differences of two files by using WinMerge
' - Differs         Returns a Dictionary with those records/lines which
'                   differ between two provided files, with the options
'                   'ignore case' and 'ignore empty records'
' - Delete          Deletes a file provided either as object or as full name
' - Extension       Returns the extension of a file's name
' - GetFile         Returns a file object for a given name
' - Arry            From/to file. Get the content of a text file as an arry
'                   of records(lines, Write an array of text to a file.
' - SectionNames    Returns a Dictionary of all section names
'                   [.....] in a file.
' - Sections        Returns named - or if no names are provideds all - sections
'                   as Dictionary. The Dictionary contains for each section a
'                   Dictionary with the name as key and the value as item
'                   (see ValueNames)
' - SectionsRemove  Removes the sections provided via their name. When no section
'                   names are provided (pp_sections) none are removed.
' - Txt             Get th content of a text file as string or write
'                   a string to a file - optionally appended
' - Value           Reads a named value from or writes a named value to a file
' - ValueNames      Returns a Dictionary of all named sections in a file with the value
'                   name as key and the value as item, in ascending order by the names.
'                   Section names may be provided as a comma delimited string of names,
'                   as a Dictionary or Collection of names. When no section names
'                   are provided all unique! value names of all sections are returned.
'                   Of duplicate names the value will be of the first one found.
' - Values          Returns a Dictionary with value as key and value name(s) in a
'                   Collection as item. When no section names are provided the values
'                   of all sections are returned. Section names may be provided as a
'                   comma delimited string of names, or a Collection or Dictionary of
'                   section name (items).
' Uses:
' - mDct            Service DctAdd is used to provide Dictionaries in ascending
'                   order by item or by key.
'                   Note! The components mTrc, fMsg, mMsg, and mErH are only used for testing
'                   by the module mTest. They are not part of the to-be-installed
'                   components of mFile.
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' This 'Common VBA Component is developed, maintained and tested (regression test
' available in mTest module and obligatory with each code modification) in the
' public Github repo: https://github.com/warbe-maker/Common-VBA-File-Services.
' Contribution in whichever form is welcome.
'
' W. Rauschenberger, Berlin Nov 2020
' ------------------------------------------------------------------------
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

Private Property Get SplitStr(ByRef s As String)
' ----------------------------------------------
' Returns the split string in string (s) used by
' VBA.Split() to turn the string into an array.
' ----------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
End Property

Public Property Let arry( _
           Optional ByVal fa_file As String, _
           Optional ByVal fa_excl_empty_lines As Boolean = False, _
           Optional ByRef fa_split As String = vbCrLf, _
           Optional ByVal fa_append As Boolean = False, _
                    ByVal fa_ar As Variant)
' -----------------------------------------------------------------
' Writes array (fa_ar) to file (fa_file) whereby the array is
' joined to a text string using the line break string (fa_split)
' which defaults to vbCrLf and is optionally returned by Arry-Get.
' -----------------------------------------------------------------
                    
    mFile.Txt(ft_file:=fa_file _
            , ft_append:=fa_append _
            , ft_split:=fa_split _
             ) = Join(fa_ar, fa_split)
             
End Property

Public Property Get arry( _
           Optional ByVal fa_file As String, _
           Optional ByVal fa_excl_empty_lines As Boolean = False, _
           Optional ByRef fa_split As String, _
           Optional ByVal fa_append As Boolean = False) As Variant
' -----------------------------------------------------------------------
' Returns the content of the file (fa_file) - a files full name - as
' array, with the used line break string returned in (fa_split).
' -----------------------------------------------------------------------
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
    
xt: arry = a1
    fa_split = sSplit
    Set cll = Nothing
    Set fso = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Public Function ValueExists( _
                          ByVal pp_file As String, _
                          ByVal pp_value As Variant, _
                 Optional ByVal pp_sections As Variant = Nothing) As Boolean
' --------------------------------------------------------------------------
' Returns True when the value (pp_value) exists in file (pp_file) - when no
' section name is provided in any section, else in the given sections.
' Section names (pp_sections) may be provided as comma delimited string or
' as Dictionary or Collection with name items.
' --------------------------------------------------------------------------
    ValueExists = mFile.Values(pp_file, pp_sections).Exists(pp_value)
End Function

Public Function ValueNameExists( _
                          ByVal pp_file As String, _
                          ByVal pp_valuename As String, _
                 Optional ByVal pp_sections As Variant = Nothing) As Boolean
' --------------------------------------------------------------------------
' Returns True when the value name (pp_valuename) exists in file (pp_file)
' - when no section name is provided in any section, else in the given
' sections. Section names (pp_sections) may be provided as comma delimited
' string or as Dictionary or Collection with name items.
' -------------------------------------------------------------------------
    ValueNameExists = mFile.ValueNames(pp_file, pp_sections).Exists(pp_valuename)
End Function
                 
Public Function SectionExists( _
                        ByVal pp_file As String, _
                        ByVal pp_section As String) As Boolean
' --------------------------------------------------------------------
' Returns True when the section (pp_section) exists in file (pp_file).
' --------------------------------------------------------------------
    SectionExists = mFile.SectionNames(pp_file).Exists(pp_section)
End Function

Public Function SectionNames( _
              Optional ByVal pp_file As String) As Dictionary
' -----------------------------------------------------------
' Returns a Dictionary of all section names [.....] in file
' (pp_file) in ascending sequence.
'
' Uses: mDct.DctAdd to order the sections in ascending
' sequence.
' -----------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim asSections()    As String
    Dim dct             As New Dictionary
    Dim i               As Long
    Dim iLen            As Long
    Dim strBuffer       As String
    
    If Len(mFile.Txt(pp_file)) = 0 Then GoTo xt
    If Not fso.FileExists(pp_file) Then GoTo xt
    
    Do While (iLen = Len(strBuffer) - 2) Or (iLen = 0)
        If strBuffer = vbNullString _
        Then strBuffer = Space$(256) _
        Else strBuffer = String(Len(strBuffer) * 2, 0)
        iLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), pp_file)
    Loop
    strBuffer = Left$(strBuffer, iLen)
    
    If Len(strBuffer) <> 0 Then
        i = 0
        asSections = Split(strBuffer, vbNullChar)
        For i = LBound(asSections) To UBound(asSections)
            If Len(asSections(i)) <> 0 _
            Then mDct.DctAdd add_dct:=dct _
                           , add_key:=asSections(i) _
                           , add_item:=asSections(i) _
                           , add_seq:=seq_ascending
        Next i
    End If
    
xt: Set SectionNames = dct
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

Public Property Get Txt( _
         Optional ByVal ft_file As Variant, _
         Optional ByVal ft_append As Boolean = True, _
         Optional ByRef ft_split As String) As String
' ----------------------------------------------------
' Returns the text file's (ft_file) content as string
' with VBA.Split() string in (ft_split). When the file
' doesn't exist a vbNullString is returned.
' Note: ft_append is not used but specified to comply
'       with the Get Property declaration.
' ----------------------------------------------------
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
        Set ts = .OpenTextFile(Filename:=sFl, IOMode:=ForReading)
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
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Public Property Let Txt( _
         Optional ByVal ft_file As Variant, _
         Optional ByVal ft_append As Boolean = True, _
         Optional ByRef ft_split As String, _
                  ByVal ft_string As String)
' -----------------------------------------------------
' Writes the string (ft_string) into the file (ft_file)
' which might be a file object of a file's full name.
' Note: ft_split is not used but specified to comply
'       with the Get Property declaration.
' -----------------------------------------------------
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
        Then Set ts = .OpenTextFile(Filename:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(Filename:=sFl, IOMode:=ForWriting)
    End With
    
    ts.WriteLine ft_string

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Public Property Get Value( _
           Optional ByVal pp_file As String, _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
' -----------------------------------------------------------
' Read a value with a specific name from a section
' [section]
' <value-name>=<value>
' -----------------------------------------------------------
    Const PROC  As String = "ValueGet"
    
    On Error GoTo eh
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant

    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString( _
                                      lpg_ApplicationName:=pp_section _
                                    , lpg_KeyName:=pp_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=pp_file _
                                     )
    vValue = Left$(sRetVal, lResult)
    Value = vValue
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Public Property Let Value( _
           Optional ByVal pp_file As String, _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' --------------------------------------------------
' Write a value under a name into a section in a
' file in the form: [section]
'                   <value-name>=<value>
' --------------------------------------------------
    Const PROC = "ValueLet"
        
    On Error GoTo eh
    Dim lChars      As Long
    Dim sValue      As String
    
    Select Case VarType(pp_value)
        Case vbBoolean: sValue = VBA.CStr(VBA.CLng(pp_value))
        Case Else:      sValue = pp_value
    End Select
    
    lChars = WritePrivateProfileString(lpw_ApplicationName:=pp_section _
                                     , lpw_KeyName:=pp_value_name _
                                     , lpw_String:=sValue _
                                     , lpw_FileName:=pp_file)
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & pp_section & "'" & vbLf & _
               "Value name = '" & pp_value_name & "'" & vbLf & _
               "Value      = '" & CStr(pp_value) & "'" & vbLf & _
               "Value file = '" & pp_file & "'"
    End If

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function AppIsInstalled(ByVal sApp As String) As Boolean
    
    Dim i As Long: i = 1
    
    Do Until VBA.Left$(Environ$(i), 5) = "Path="
        i = i + 1
    Loop
    AppIsInstalled = InStr(Environ$(i), sApp) <> 0

End Function

Public Function Compare(ByVal fc_file_left As String, _
                        ByVal fc_left_title As String, _
                        ByVal fc_file_right As String, _
                        ByVal fc_right_title As String) As Long
' ---------------------------------------------------------------------
' Compares two text files by means of WinMerge. An error is raised when
' WinMerge is not installed of one of the two files doesn't exist.
' ----------------------------------------------------------------------
    Const PROC = "Compare"
    
    On Error GoTo eh
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String
    Dim fso             As New FileSystemObject
    Dim wshShell        As Object
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="WinMerge is obligatory for the Compare service of this module but not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(fc_file_left) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & fc_file_left & """ does not exist!"
    
    If Not fso.FileExists(fc_file_right) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
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
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' This is a kind of universal error message which includes a debugging option.
' It may be copied into any module - turned into a Private function. When the/my
' Common VBA Error Handling Component (ErH) is installed and the Conditional
' Compile Argument 'CommErHComp = 1' the error message will be displayed by
' means of the Common VBA Message Component (fMsg, mMsg).
'
' Usage: When this procedure is copied as a Private Function into any desired
'        module an error handling which consideres the possible Conditional
'        Compile Argument 'Debugging = 1' will look as follows
'
'            Const PROC = "procedure-name"
'            On Error Goto eh
'        ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC)
'               Case vbYes: Stop: Resume
'               Case vbNo:  Resume Next
'               Case Else:  Goto xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Used:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
' ------------------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
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
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              err_dscrptn & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    
#If Debugging Then
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
#If CommErHComp Then
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    ErrMsg = ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    '~~ Translate back the elaborated reply buttons mErrH.ErrMsg displays and returns to the simple yes/No/Cancel
    '~~ replies with the VBA MsgBox.
    Select Case ErrMsg
        Case mErH.DebugOptResumeErrorLine:  ErrMsg = vbYes
        Case mErH.DebugOptResumeNext:       ErrMsg = vbNo
        Case Else:                          ErrMsg = vbCancel
    End Select
#Else
    '~~ When the Common VBA Error Handling Component (ErH) is not used/installed there might still be the
    '~~ Common VBA Message Component (Msg) be installed/used
#If CommMsgComp Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
#Else
    '~~ None of the Common Components is installed/used
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
#End If
#End If
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mFile." & sProc
End Function

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
    Dim cll     As New Collection
    Dim cllRet  As New Collection

    cll.Add fso.GetFolder(fs_root)

    Do While cll.Count > 0
        Set fo = cll(1)
        cll.Remove 1 'dequeue
        If fs_in_subfolders Then
            For Each sfo In fo.SubFolders
                cll.Add sfo 'enqueue
            Next sfo
        End If
        For Each fl In fo.Files
            If fl.Path Like fs_mask Then
                DoEvents
                cllRet.Add fl
                If cllRet.Count >= fs_stop_after Then GoTo xt
            End If
        Next fl
    Loop

xt: Set Search = cllRet
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

Public Function Exists(ByVal fe_file As Variant, _
              Optional ByRef fe_fso As File = Nothing, _
              Optional ByRef fe_cll As Collection = Nothing) As Boolean
' ------------------------------------------------------------------
' Returns TRUE when the file (fe_file) - which may be a file object
' or a file's full name - exists and furthermore:
' - when the file's full name ends with a wildcard * all
'   subfolders are scanned and any file which meets the criteria
'   is returned as File object in a collection (fe_cll),
' - when the files's full name does not end with a wildcard * the
'   existing file is returned as a File object (fe_fso).
' ----------------------------------------------------------------
    Const PROC  As String = "Exists"    ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim sTest   As String
    Dim sFile   As String
    Dim fldr    As Folder
    Dim sfldr   As Folder   ' Sub-Folder
    Dim fl      As File
    Dim sPath   As String
    Dim queue   As Collection

    Exists = False
    Set fe_cll = New Collection

    If TypeName(fe_file) <> "File" And TypeName(fe_file) <> "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The File (parameter fe_file) for the File's existence check is neither a full path/file name nor a file object!"
    If Not TypeName(fe_fso) = "Nothing" And Not TypeName(fe_fso) = "File" _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided return parameter (fe_fso) is not a File type!"
    If Not TypeName(fe_cll) = "Nothing" And Not TypeName(fe_cll) = "Collection" _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided return parameter (fe_cll) is not a Collection type!"

    If TypeOf fe_file Is File Then
        With New FileSystemObject
            On Error Resume Next
            sTest = fe_file.Name
            Exists = Err.Number = 0
            If Exists Then
                '~~ Return the existing file as File object
                Set fe_fso = .GetFile(fe_file.Path)
                GoTo xt
            End If
        End With
    ElseIf VarType(fe_file) = vbString Then
        With New FileSystemObject
            sFile = Split(fe_file, "\")(UBound(Split(fe_file, "\")))
            If Not Right(sFile, 1) = "*" Then
                Exists = .FileExists(fe_file)
                If Exists Then
                    '~~ Return the existing file as File object
                    Set fe_fso = .GetFile(fe_file)
                    GoTo xt
                End If
            Else
                sPath = Replace(fe_file, "\" & sFile, vbNullString)
                sFile = Replace(sFile, "*", vbNullString)
                '~~ Wildcard file existence check is due
                Set fldr = .GetFolder(sPath)
                Set queue = New Collection
                queue.Add .GetFolder(sPath)

                Do While queue.Count > 0
                    Set fldr = queue(queue.Count)
                    queue.Remove queue.Count ' dequeue the processed subfolder
                    For Each sfldr In fldr.SubFolders
                        queue.Add sfldr ' enqueue (collect) all subfolders
                    Next sfldr
                    For Each fl In fldr.Files
                        If InStr(fl.Name, sFile) <> 0 And VBA.Left$(fl.Name, 1) <> "~" Then
                            '~~ Return the existing file which meets the search criteria
                            '~~ as File object in a collection
                            fe_cll.Add fl
                         End If
                    Next fl
                Loop
                If fe_cll.Count > 0 Then Exists = True
            End If
        End With
    End If

xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & fc_file1 & """ does not exist!"
    
    If Not fso.FileExists(fc_file2) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
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
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

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

Public Function GetFile(ByVal fg_path As String) As File
    With New FileSystemObject
        Set GetFile = .GetFile(fg_path)
    End With
End Function

Public Sub NameRemove(ByVal pp_file As String, _
                      ByVal pp_section As String, _
                      ByVal pp_value_name As String)
' --------------------------------------------------
'
' --------------------------------------------------
    DeletePrivateProfileKey Section:=pp_section _
                          , Key:=pp_value_name _
                          , Setting:=0 _
                          , Name:=pp_file
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
    Dim s1          As String
    Dim s2          As String
    Dim a1          As Variant
    Dim a2          As Variant
    Dim dctF1       As New Dictionary
    Dim dctF2       As New Dictionary
    
    Dim dct1        As New Dictionary
    Dim dct2        As New Dictionary
    Dim sTest1      As String
    Dim sTest2      As String
    Dim dctDif      As Dictionary
    Dim i           As Long
    Dim lDiffLine   As Long
    Dim sFile1      As String
    Dim sFile2      As String
    Dim v           As Variant
    Dim sSplit      As String
    
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
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    
    For i = 1 To mBasic.Min(Len(di_file_left), Len(di_file_right))
        If VBA.Mid$(di_file_left, i, 1) <> VBA.Mid$(di_file_right, i, 1) _
        Then Exit For
    Next i
    i = i - 2
    sFileLeft = "..." & VBA.Right$(di_file_left, Len(di_file_left) - i) & "Line " & Format(di_line, "0000") & ": "
    sFileRight = "..." & VBA.Right$(di_file_right, Len(di_file_right) - i) & "Line " & Format(di_line, "0000") & ": "
    
    DiffItem = sFileLeft & "'" & di_line_left & "'" & vbLf & sFileRight & "'" & di_line_right & "'"

End Function

'Public Sub SectionMove()
'
'End Sub
'
'Public Sub SectionReplace()
'
'End Sub

Public Sub SectionsCopy(ByVal pp_source As String, _
                        ByVal pp_target As String, _
               Optional ByVal pp_sections As Variant = Nothing, _
               Optional ByVal pp_replace As Boolean = False)
' ---------------------------------------------------------------
' Copies sections from file (pp_source) to file (pp_target), when
' no section names (pp_section_names) are provided all, by
' default (pp_replace) the sections are merged.
' ---------------------------------------------------------------
    Const PROC = "SectionCopy"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim dct         As Dictionary
    Dim vNames      As Variant
    
    '~~ Provide all section names when no section named are provided via pp_sections
    Set vNames = NamesInArg(pp_sections)
    If vNames.Count = 0 Then Set vNames = mFile.SectionNames(pp_source)
    
    '~~ Get the named or all sections as Dictionary
    Set dct = mFile.Sections(pp_file:=pp_source _
                           , pp_sections:=vNames _
                            )
     
     If fso.FileExists(pp_target) And pp_replace _
     Then mFile.SectionsRemove pp_file:=pp_target _
                             , pp_sections:=vNames
     
     '~~ Write all sections from the source file to the target file
     mFile.Sections(pp_target) = dct

xt: Set vNames = Nothing
    Set dct = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Property Get Sections( _
                   Optional ByVal pp_file As String, _
                   Optional ByVal pp_sections As Variant = Nothing, _
                   Optional ByVal pp_replace As Boolean = False) As Dictionary
' ----------------------------------------------------------------------------
' Returns the named sections (pp_section_names) - if not provided all sections - in
' file (pp_file) as Dictionary with the section name as the key - in ascending order! -
' and a Dictionary of the sections values as item with the value name as key and the
' value as item.
' The section names (pp_section_names) may be a comma delimmited string of names a
' Dictionary or a Collection, both with the item as name.
'
' Requires: Service mDct.DctAdd to order the sections in ascending sequence.
' -------------------------------------------------------------------------------------
    Const PROC = "Sections-Get"
    
    On Error GoTo eh
    Dim dctS    As New Dictionary   ' Result Sections
    Dim dctV    As Dictionary       ' Section values
    Dim v       As Variant
    Dim sName   As String           ' A section's name
    Dim vNames  As Variant
    
    pp_replace = pp_replace ' not used! declared for property Get/Let conformity
    '~~ Provide all section names when no section named are provided via pp_sections
    Set vNames = NamesInArg(pp_sections)
    If vNames.Count = 0 Then Set vNames = mFile.SectionNames(pp_file)
    
    For Each v In vNames
        sName = v
        Set dctV = mFile.ValueNames(pp_file:=pp_file _
                                  , pp_sections:=sName _
                                   )
        mDct.DctAdd add_dct:=dctS _
                  , add_key:=sName _
                  , add_item:=dctV _
                  , add_seq:=seq_ascending _
                  , add_order:=order_bykey
    Next v

xt: Set Sections = dctS
    Set dctS = Nothing
    Set dctV = Nothing
    Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

#If Test Then
Public Property Let Sections( _
              Optional ByVal pp_file As String, _
              Optional ByVal pp_sections As Variant = Nothing, _
              Optional ByVal pp_replace As Boolean = False, _
                       ByVal pp_dct As Dictionary)
#Else
Private Property Let Sections( _
              Optional ByVal pp_file As String, _
              Optional ByVal pp_sections As Variant = Nothing, _
              Optional ByVal pp_replace As Boolean = False, _
                       ByVal pp_dct As Dictionary)
#End If
' ------------------------------------------------------------------------
' Writes the sections in a Dictionary (pp_dct) to the file (pp_file) by
' default merging or by replacing existing sections.
' ------------------------------------------------------------------------
    Const PROC = "Sections-Get"
    
    On Error GoTo eh
    Dim vN          As Variant
    Dim vS          As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim sName       As String
    
    pp_replace = pp_replace     ' not used! declared for Property Get/Let conformity only
    Set pp_sections = Nothing   ' not used! declared for Property Get/Let conformity only
    
    For Each vS In pp_dct
        sSection = vS
        Set dctValues = pp_dct(vS)
        For Each vN In dctValues
            sName = vN
            mFile.Value(pp_file:=pp_file _
                      , pp_section:=sSection _
                      , pp_value_name:=sName _
                       ) = dctValues.Item(vN)
        Next vN
    Next vS
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Private Function NamesInArg( _
            Optional ByVal v As Variant = Nothing) As Collection
' --------------------------------------------------------------
' Returns (v) as Collection of string items whereby (v) may not
' be provided, a comma delimited string, or a Dictionary or
' Collection of string items.
' --------------------------------------------------------------
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
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

Public Sub SectionsRemove( _
                    ByVal pp_file As String, _
           Optional ByVal pp_sections As Variant = Nothing)
' ----------------------------------------------------------
' Removes the sections provided via their name. When no
' section names are provided (pp_sections) none are removed.
' ----------------------------------------------------------
    Const PROC = "SectionsRemove"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim vNames  As Variant
    
    '~~ Provide all section names when no section named are provided via pp_sections
    Set vNames = NamesInArg(pp_sections)
    If vNames.Count = 0 Then GoTo xt
    
    For Each v In vNames
        DeletePrivateProfileSection Section:=v _
                                  , NoKey:=0 _
                                  , NoSetting:=0 _
                                  , Name:=pp_file
    Next v
    
xt: Set vNames = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Function SelectFile( _
            Optional ByVal sel_init_path As String = vbNullString, _
            Optional ByVal sel_filters As String = "*.*", _
            Optional ByVal sel_filter_name As String = "File", _
            Optional ByVal sel_title As String = vbNullString, _
            Optional ByRef sel_result As File) As Boolean
' --------------------------------------------------------------
' When a file had been selected TRUE is returned and the
' selected file is returned as File object (sel_result).
' --------------------------------------------------------------
    Const PROC = "SelectFile"
    
    On Error GoTo eh
    Dim fDialog As FileDialog
    Dim v       As Variant

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        If sel_title = vbNullString _
        Then .Title = "Select a(n) " & sel_filter_name _
        Else .Title = sel_title
        .InitialFileName = sel_init_path
        .Filters.Clear
        For Each v In Split(sel_filters, ",")
            .Filters.Add sel_filter_name, v
         Next v
         
        If .Show = -1 Then
            '~~ A fie had been selected
           With New FileSystemObject
            Set sel_result = .GetFile(fDialog.SelectedItems(1))
            SelectFile = True
           End With
        End If
        '~~ When no file had been selected the sel_result will be Nothing
    End With

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

Private Function ShellRun(sCmd As String) As String
' ------------------------------------------------------
' Run a shell command, returning the output as a string.
' ------------------------------------------------------
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
    mBasic.ArrayTrimm a
    
    For i = LBound(a) To UBound(a)
        dct.Add i + 1, a(i)
    Next i
        
xt: Set Dict = dct
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Property

Public Function ValueNames( _
                     ByVal pp_file As String, _
            Optional ByVal pp_sections As Variant = Nothing) As Dictionary
' ------------------------------------------------------------------------
' Returns a Dictionary of all value names (with the value name as key and
' the value as item) in file (pp_file) of the sections (pp_sections) in
' asscending order. Sections names (pp_sections) may be provided as a
' comma delimited string of names, or a Dictionary or Collection of name
' items. When no section names (pp_sections) are provided all unique!
' value names of all sections in file (pp_file) are returned. Of duplicate
' names the value will be of the first one found.
'
' Uses: mDct.DctAdd to order the sections in ascending sequence.
' ------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    Dim asNames()       As String
    Dim dctNames        As New Dictionary
    Dim i               As Long
    Dim lResult         As Long
    Dim sNames          As String
    Dim strBuffer       As String
    Dim v               As Variant
    Dim sSection        As String
    Dim sName           As String
    Dim vNames          As Variant
    
    '~~ When no section names are provided the name of all values in all
    '~~ sections are collected in ascending order ignoring duplicates
    Set vNames = NamesInArg(pp_sections)
    If vNames.Count = 0 Then Set vNames = mFile.SectionNames(pp_file)
    
    For Each v In vNames
        sSection = v
        '~~> Retrieve the names for the provided section
        strBuffer = Space$(32767)
        lResult = GetPrivateProfileString(lpg_ApplicationName:=sSection _
                                        , lpg_KeyName:=vbNullString _
                                        , lpg_Default:=vbNullString _
                                        , lpg_ReturnedString:=strBuffer _
                                        , nSize:=Len(strBuffer) _
                                        , lpg_FileName:=pp_file _
                                         )
        sNames = Left$(strBuffer, lResult)
    
        If sNames <> vbNullString Then                                         ' If there were any names
            asNames = Split(sNames, vbNullChar)                      ' have them split into an array
            For i = LBound(asNames) To UBound(asNames)
                sName = asNames(i)
                If Len(sName) <> 0 Then
                    If Not dctNames.Exists(sName) _
                    Then mDct.DctAdd add_dct:=dctNames _
                                   , add_key:=sName _
                                   , add_item:=mFile.Value(pp_file:=pp_file, pp_section:=sSection, pp_value_name:=sName) _
                                   , add_seq:=seq_ascending
                End If
            Next i
        End If
    Next v
        
    Set ValueNames = dctNames

xt: Set dctNames = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

Public Function Values( _
                 ByVal pp_file As String, _
        Optional ByVal pp_sections As Variant = Nothing) As Dictionary
' --------------------------------------------------------------------
' Returns a Dictionary with the values in file (pp_file) as key.
' Attention! Because the same value may appear with several names when
' all names of all sections are returned, the value names are returned
' as Collection item. When (pp_sections) is ommited of all sections,
' otherwise of the named sections which may be provided as a comma
' delimited string of names, or a Collection or Dictionary of section
' name items.
'
' Uses: mDct.DctAdd to order the sections in ascending sequence.
' --------------------------------------------------------------------
    Const PROC = "Values"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim vName       As Variant
    Dim vNames      As Variant ' a Collection of provided section names or a Dictionary of all section names
    Dim vSection    As Variant
    Dim sSection    As String
    Dim sValName    As String
    Dim dctNames    As Dictionary
    Dim sValue      As String
    Dim cllNames    As Collection
    
    Set vNames = NamesInArg(pp_sections)
    If vNames.Count = 0 Then Set vNames = mFile.SectionNames(pp_file)
    
    For Each vSection In vNames
        sSection = vSection
        Set dctNames = mFile.ValueNames(pp_file:=pp_file, pp_sections:=sSection)
        For Each vName In dctNames
            sValue = dctNames(vName)
            sValName = vName
            If Not dct.Exists(sValue) Then
                Set cllNames = New Collection
                cllNames.Add sValName
                mDct.DctAdd add_dct:=dct _
                          , add_key:=sValue _
                          , add_item:=cllNames _
                          , add_seq:=seq_ascending
            Else
                Set cllNames = dct.Item(sValue)
                cllNames.Add sValName
                dct.Remove sValue
                mDct.DctAdd add_dct:=dct _
                          , add_key:=sValue _
                          , add_item:=cllNames _
                          , add_seq:=seq_ascending
            End If
        Next vName
    Next vSection
    
xt: Set Values = dct
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function


