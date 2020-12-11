Attribute VB_Name = "mFl"
Option Explicit
Option Compare Text
Option Private Module
' --------------------------------------------------------------
' Standard  Module mFile
'           Common methods and functions regarding file objects.
'
' Methods:  Exists      Returns TRUE when the file exists
'           Compare     Displays differences of two files by
'                       means of WinMerge
'           Differ      Returns TRUE when two files have a
'                       different content
'           Delete      Deletes a file
'           Extension   Returns the extension of a file's name
'           GetFile     Returns a file object for a given name
'           ToArray     Returns a file's content in an array
'
' Uses:     No other components
'           (mTrc, fMsg, mMsg and mErH are used by module mTest only).
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' W. Rauschenberger, Berlin Nov 2020
' -----------------------------------------------------------------------------------
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
                ByVal name As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileKey _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal Key As String, _
                ByVal Setting As Long, _
                ByVal name As String) As Long
                 
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
'                ByVal name As String) As Long

Private dctSection As Dictionary

Public Property Get Value( _
           Optional ByVal vl_file As String, _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
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
                                      lpg_ApplicationName:=vl_section _
                                    , lpg_KeyName:=vl_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=vl_file _
                                     )
    vValue = Left$(sRetVal, lResult)
    Value = vValue
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Let Value( _
           Optional ByVal vl_file As String, _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
' --------------------------------------------------
' Write a value under a name into a section in a
' file in the form: [section]
'                   <value-name>=<value>
' --------------------------------------------------
    Const PROC = "ValueLet"
        
    On Error GoTo eh
    Dim lChars As Long
    
    Select Case VarType(vl_value)
        Case vbBoolean: lChars = WritePrivateProfileString(lpw_ApplicationName:=vl_section _
                                                         , lpw_KeyName:=vl_value_name _
                                                         , lpw_String:=VBA.CStr(VBA.CLng(vl_value)) _
                                                         , lpw_FileName:=vl_file _
                                                         )
        Case Else:      lChars = WritePrivateProfileString(vl_section, vl_value_name, CStr(vl_value), vl_file)
    End Select
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & vl_section & "'" & vbLf & _
               "Value name = '" & vl_value_name & "'" & vbLf & _
               "Value      = '" & CStr(vl_value) & "'" & vbLf & _
               "Value file = '" & vl_file & "'"
    End If

xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Function AppErr(ByVal err_no As Long) As Long
' -----------------------------------------------------------------
' Used with Err.Raise AppErr(<l>).
' When the error number <l> is > 0 it is considered an "Application
' Error Number and vbObjectErrror is added to it into a negative
' number in order not to confuse with a VB runtime error.
' When the error number <l> is negative it is considered an
' Application Error and vbObjectError is added to convert it back
' into its origin positive number.
' ------------------------------------------------------------------
    If err_no < 0 Then
        AppErr = err_no - vbObjectError
    Else
        AppErr = vbObjectError + err_no
    End If
End Function

Private Function AppIsInstalled(ByVal sApp As String) As Boolean
    
    Dim i As Long: i = 1
    
    Do Until Left(Environ$(i), 5) = "Path="
        i = i + 1
    Loop
    AppIsInstalled = InStr(Environ$(i), sApp) <> 0

End Function

Public Function Compare(ByVal file_left_full_name As String, _
                        ByVal file_left_title As String, _
                        ByVal file_right_full_name As String, _
                        ByVal file_right_title As String) As Long
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
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="WinMerge is obligatory for the Compare service of this module but not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(file_left_full_name) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & file_left_full_name & """ does not exist!"
    
    If Not fso.FileExists(file_right_full_name) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & file_right_full_name & """ does not exist!"
    
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & file_left_title & DQUOTE & _
               " /dr " & DQUOTE & file_right_title & DQUOTE & " " & _
               """" & file_left_full_name & """" & " " & _
               """" & file_right_full_name & """"
    
    With New WshShell
        Compare = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub Delete(ByVal v As Variant)

    Dim fl  As FILE

    With New FileSystemObject
        If TypeName(v) = "File" Then
            Set fl = v
            .DeleteFile fl.PATH
        ElseIf TypeName(v) = "String" Then
            If .FileExists(v) Then
                .DeleteFile v
            End If
        End If
    End With
    
End Sub

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString)
' ------------------------------------------------------
' This Common Component does not have its own error
' handling. Instead it passes on any error to the
' caller's error handling.
' ------------------------------------------------------
    
    If err_no = 0 Then err_no = Err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description

    Err.Raise Number:=err_no, Source:=err_source, Description:=err_dscrptn

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.name & ": mFl." & sProc
End Function

Public Function Exists(ByVal xst_file As Variant, _
              Optional ByRef xst_fso As FILE = Nothing, _
              Optional ByRef xst_cll As Collection = Nothing) As Boolean
' ------------------------------------------------------------------
' Returns TRUE when the file (xst_file) - which may be a file object
' or a file's full name - exists and furthermore:
' - when the file's full name ends with a wildcard * all
'   subfolders are scanned and any file which meets the criteria
'   is returned as File object in a collection (xst_cll),
' - when the files's full name does not end with a wildcard * the
'   existing file is returned as a File object (xst_fso).
' ----------------------------------------------------------------
    Const PROC  As String = "Exists"    ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim sTest   As String
    Dim sFile   As String
    Dim fldr    As Folder
    Dim sfldr   As Folder   ' Sub-Folder
    Dim fl      As FILE
    Dim sPath   As String
    Dim queue   As Collection

    Exists = False
    Set xst_cll = New Collection

    If TypeName(xst_file) <> "File" And TypeName(xst_file) <> "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The File (parameter xst_file) for the File's existence check is neither a full path/file name nor a file object!"
    If Not TypeName(xst_fso) = "Nothing" And Not TypeName(xst_fso) = "File" _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided return parameter (xst_fso) is not a File type!"
    If Not TypeName(xst_cll) = "Nothing" And Not TypeName(xst_cll) = "Collection" _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided return parameter (xst_cll) is not a Collection type!"

    If TypeOf xst_file Is FILE Then
        With New FileSystemObject
            On Error Resume Next
            sTest = xst_file.name
            Exists = Err.Number = 0
            If Exists Then
                '~~ Return the existing file as File object
                Set xst_fso = .GetFile(xst_file.PATH)
                GoTo xt
            End If
        End With
    ElseIf VarType(xst_file) = vbString Then
        With New FileSystemObject
            sFile = Split(xst_file, "\")(UBound(Split(xst_file, "\")))
            If Not Right(sFile, 1) = "*" Then
                Exists = .FileExists(xst_file)
                If Exists Then
                    '~~ Return the existing file as File object
                    Set xst_fso = .GetFile(xst_file)
                    GoTo xt
                End If
            Else
                sPath = Replace(xst_file, "\" & sFile, vbNullString)
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
                        If InStr(fl.name, sFile) <> 0 And Left(fl.name, 1) <> "~" Then
                            '~~ Return the existing file which meets the search criteria
                            '~~ as File object in a collection
                            xst_cll.Add fl
                         End If
                    Next fl
                Loop
                If xst_cll.Count > 0 Then Exists = True
            End If
        End With
    End If

xt: Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function Extension(ByVal vFile As Variant)

    With New FileSystemObject
        If TypeName(vFile) = "File" Then
            Extension = .GetExtensionName(vFile.PATH)
        Else
            Extension = .GetExtensionName(vFile)
        End If
    End With

End Function

Public Function GetFile(ByVal gf_path As String) As FILE
    With New FileSystemObject
        Set GetFile = .GetFile(gf_path)
    End With
End Function

Public Sub NameRemove(ByVal nr_file As String, _
                      ByVal nr_section As String, _
                      ByVal nr_name As String)
' --------------------------------------------------
'
' --------------------------------------------------
    DeletePrivateProfileKey Section:=nr_section, Key:=nr_name, Setting:=0, name:=nr_file
End Sub

Private Function Fc(ByVal fc_file1 As String, fc_file2 As String)
    Const PROC = "Compare"
    
    On Error GoTo eh
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String
    Dim fso             As New FileSystemObject
            
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
    
    With New WshShell
        Fc = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

Public Function sDiffer( _
                  ByVal dif_file1 As FILE, _
                  ByVal dif_file2 As FILE, _
         Optional ByVal dif_stop_after As Long = 1, _
         Optional ByVal dif_ignore_empty_records As Boolean = False, _
         Optional ByRef dif_lines As Variant) As Boolean
' -----------------------------------------------------------------------------
' Returns TRUE when the content of file (dif_file1) differs from the content in
' file (dif_file2). The comparison stops after (dif_stop_after) detected
' differences. The detected different lines are optionally returned (vResult).
' ------------------------------------------------------------------------------
    Const PROC = "sDiffer"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    Dim vLines  As Variant

    a1 = mFl.ToArray(ta_file:=dif_file1, ta_exclude_empty_records:=dif_ignore_empty_records)
    a2 = mFl.ToArray(ta_file:=dif_file2, ta_exclude_empty_records:=dif_ignore_empty_records)
    vLines = mBasic.ArrayCompare(ac_a1:=a1, ac_a2:=a2, ac_stop_after:=dif_stop_after)
    If mBasic.ArrayIsAllocated(arr:=vLines) Then
        sDiffer = True
    End If
    dif_lines = vLines
    
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

Private Function SectionsGet( _
                       ByVal sg_file As String, _
                       ByVal sg_section_names As Variant) As Dictionary
' -----------------------------------------------------------------------
' Returns a Dictionary with complete sections, one for each provided
' section name (sg_section_names). Each section is identified by the key
' and the item is a Dictionary of all values - with the value name as
' the key and the value as the item.
' Recursively called until the sg_section_names argument is a Collection.
' -----------------------------------------------------------------------
    Const PROC = "SectionGet"
    
    On Error GoTo eh
    Dim dctSections As Dictionary
    Dim dctValues   As Dictionary
    Dim i           As Long
    Dim sSection    As String
    Dim vSections   As Variant
    Dim cll         As New Collection
    
    Select Case VarType(sg_section_names)
        Case Is >= vbArray
            For i = LBound(sg_section_names) To UBound(sg_section_names)
                cll.Add sg_section_names(i)
            Next i
            Set dctSections = SectionsGet(sg_file:=sg_file, sg_section_names:=cll)
            GoTo xt
        Case vbString
            For i = LBound(Split(sg_section_names, ",")) To UBound(Split(sg_section_names, ","))
                cll.Add Split(sg_section_names, ",")(i)
            Next i
            Set dctSections = SectionsGet(sg_file:=sg_file, sg_section_names:=cll)
            GoTo xt
        Case vbObject
            Select Case TypeName(sg_section_names)
                Case "Dictionary"
                    For i = 0 To sg_section_names.Count
                        cll.Add sg_section_names.Items()(i)
                    Next i
                    Set dctSections = SectionsGet(sg_file:=sg_file, sg_section_names:=cll)
                    GoTo xt
                Case "Collection"
                    Set dctSections = New Dictionary
                    Set cll = sg_section_names
                    For i = 1 To cll.Count
                        sSection = cll(i)
                        Set dctValues = mFl.Values(vl_file:=sg_file _
                                                 , vl_section:=sSection _
                                                  )
                        dctSections.Add Key:=sSection _
                                     , Item:=dctValues
                    Next i
                    GoTo xt
                Case Else: GoTo xt
            End Select
    End Select
            
xt: Set cll = Nothing
    If dctSections.Count = 0 _
    Then Err.Raise Number:=AppErr(1) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The name of the section(s) is provided neither as a comma delimited " & _
                                "string, nor an array of strings, nor a Dictionary, nor a Collection!"
    Set SectionsGet = dctSections
    Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Private Sub SectionsLet( _
                  ByVal sl_sections As Dictionary, _
                  ByVal sl_file As String)
' -----------------------------------------------------
' Writes for each item in the Dictionary (sl_sections
' provided by the SectionsGet service) a complete
' section into the file (sl_file).
' In other words: Together with the SectionsGet service
' this allows to transfer sections between files
' -----------------------------------------------------
    Const PROC = "SectionLet"
    
    On Error GoTo eh
    Dim vn          As Variant
    Dim vs          As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim vValue      As Variant
    Dim sName       As String
    
    For Each vs In sl_sections
        sSection = vs
        Set dctValues = sl_sections(vs)
        For Each vn In dctValues
            sName = vn
            vValue = dctValues(vn)
            mFl.Value(vl_file:=sl_file _
                    , vl_section:=sSection _
                    , vl_value_name:=sName _
                    ) = vValue
        Next vn
    Next vs
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Public Function SectionNames(ByVal sn_file As String) As Dictionary
' -----------------------------------------------------------------
' Returns a Dictionary of all section names [.....] in a file.
' -----------------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim asSections()     As String
    Dim dct             As Dictionary
    Dim i               As Long
    Dim iLen            As Long
    Dim strBuffer       As String
    Dim sSectionName    As String
    
    Set dct = New Dictionary
    Set SectionNames = New Dictionary
    
    Do While (iLen = Len(strBuffer) - 2) Or (iLen = 0)
        If strBuffer = vbNullString _
        Then strBuffer = mBasic.Space$(256) _
        Else strBuffer = String(Len(strBuffer) * 2, 0)
        iLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), sn_file)
    Loop
    strBuffer = Left$(strBuffer, iLen)
    
    If Len(strBuffer) <> 0 Then
        i = 0
        asSections = Split(strBuffer, vbNullChar)
        For i = LBound(asSections) To UBound(asSections)
            sSectionName = asSections(i)
            If Len(sSectionName) <> 0 Then
                If Not dct.Exists(sSectionName) _
                Then mDct.DctAdd add_dct:=dct, add_key:=sSectionName, add_item:=sSectionName, add_seq:=seq_ascending
            End If
        Next i
    End If
    
    Set SectionNames = dct

xt: Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Sub SectionsRemove( _
                    ByVal sr_file As String, _
                    ByVal sr_section_names As Variant)
' ----------------------------------------------------
' Removes all sections in sr_section_names from file
' sr_file.
' ----------------------------------------------------
    Const PROC = "SectionsRemove"
    
    On Error GoTo eh
    Dim i               As Long
    Dim cll             As New Collection
    Dim v               As Variant
    Dim sSectionName    As String
    
    Select Case VarType(sr_section_names)
        Case Is >= vbArray
            For i = LBound(sr_section_names) To UBound(sr_section_names)
                cll.Add sr_section_names(i)
            Next i
            SectionsRemove sr_file:=sr_file, sr_section_names:=cll
            GoTo xt
        Case vbString
            For i = LBound(Split(sr_section_names, ",")) To UBound(Split(sr_section_names, ","))
                cll.Add Split(sr_section_names, ",")(i)
            Next i
            SectionsRemove sr_file:=sr_file, sr_section_names:=cll
            GoTo xt
        Case vbObject
            Select Case TypeName(sr_section_names)
                Case "Dictionary"
                    For i = 0 To sr_section_names.Count
                        cll.Add sr_section_names.Items()(i)
                    Next i
                    SectionsRemove sr_file:=sr_file, sr_section_names:=cll
                    GoTo xt
                Case "Collection"
                    For Each v In sr_section_names
                        sSectionName = v
                        DeletePrivateProfileSection Section:=sSectionName, NoKey:=0, NoSetting:=0, name:=sr_file
                    Next v
                    GoTo xt
                Case Else: GoTo xt
            End Select
    End Select
    
xt: Set cll = Nothing
    Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Public Function SelectFile( _
            Optional ByVal sel_init_path As String = vbNullString, _
            Optional ByVal sel_filters As String = "*.*", _
            Optional ByVal sel_filter_name As String = "File", _
            Optional ByVal sel_title As String = vbNullString, _
            Optional ByRef sel_result As FILE) As Boolean
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
         
        If .show = -1 Then
            '~~ A fie had been selected
           With New FileSystemObject
            Set sel_result = .GetFile(fDialog.SelectedItems(1))
            SelectFile = True
           End With
        End If
        '~~ When no file had been selected the sel_result will be Nothing
    End With

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

Public Function ToArray(ByVal ta_file As Variant, _
               Optional ByVal ta_exclude_empty_records As Boolean = False) As String()
' ------------------------------------------------------------------------------------
' Returns the content of the file (vFile) - which may be provided as file object or
' full file name - as array by considering any kind of line break characters.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "ToArray"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim a       As Variant
    Dim a1()    As String
    Dim sSplit  As String
    Dim fso     As FILE
    Dim sFile   As String
    Dim i       As Long
    Dim j       As Long
    
    If Not Exists(ta_file, fso) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (vFile) does not exist!"
    
    '~~ Unload file into a test stream
    With New FileSystemObject
        Set ts = .OpenTextFile(fso.PATH, 1)
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
    
    If Not ta_exclude_empty_records Then
        ToArray = a
    Else
        '~~ Count empty records
        j = 0
        For i = LBound(a) To UBound(a)
            If Len(Trim$(a(i))) = 0 Then j = j + 1
        Next i
        j = UBound(a) - j
        ReDim a1(j - 1)
        j = 0
        For i = LBound(a) To UBound(a)
            If Len(Trim$(a(i))) > 0 Then
                a1(j) = a(i)
                j = j + 1
            End If
        Next i
        ToArray = a1
    End If
    
xt: Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function ToDict(ByVal td_file As Variant) As Dictionary
' ----------------------------------------------------------
' Returns the content of the file (td_file) - which may be
' provided as file object or full file name - as Dictionary
' by considering any kind of line break characters.
' ---------------------------------------------------------
    Const PROC  As String = "ToDict"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim a       As Variant
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim fso     As FILE
    Dim sFile   As String
    Dim i       As Long
    
    If Not Exists(td_file, fso) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (td_file) does not exist!"
    
    '~~ Unload file into a test stream
    With New FileSystemObject
        Set ts = .OpenTextFile(fso.PATH, 1)
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
        
xt: Set ToDict = dct
    Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function ValueNames( _
                     ByVal vn_file As String, _
            Optional ByVal vn_section As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all value names in a given file, when a section is
' provides only of this section.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    Dim asNames()       As String
    Dim dctSectionNames As Dictionary
    Dim dctNames        As Dictionary
    Dim i               As Long
    Dim lResult         As Long
    Dim sNames          As String
    Dim strBuffer       As String
    Dim v               As Variant
    Dim sSection        As String
    Dim sName           As String
    
    Set dctNames = New Dictionary
    Set dctSectionNames = New Dictionary
    Set ValueNames = New Dictionary    ' Empty in case no names are returned
    
    If vn_section <> vbNullString Then
        '~~> Retrieve the names for the provided Aspect
        strBuffer = mBasic.Space$(32767)
        lResult = GetPrivateProfileString(lpg_ApplicationName:=vn_section _
                                        , lpg_KeyName:=vbNullString _
                                        , lpg_Default:=vbNullString _
                                        , lpg_ReturnedString:=strBuffer _
                                        , nSize:=Len(strBuffer) _
                                        , lpg_FileName:=vn_file _
                                         )
        sNames = Left$(strBuffer, lResult)
    
        If sNames <> vbNullString Then                                         ' If there were any names
            asNames = Split(sNames, vbNullChar)                      ' have them split into an array
            For i = LBound(asNames) To UBound(asNames)
                sName = asNames(i)
                If Len(sName) <> 0 Then
                    mDct.DctAdd add_dct:=dctNames, add_key:=sName, add_item:=sName, add_seq:=seq_ascending
                End If
            Next i
        End If
    Else
        '~~> Retrieve the names of all sections
        Set dctSectionNames = SectionNames(sn_file:=vn_file)
        For Each v In dctSectionNames
            sSection = v
            strBuffer = mBasic.Space$(32767)
            lResult = GetPrivateProfileString(lpg_ApplicationName:=sSection _
                                            , lpg_KeyName:=vbNullString _
                                            , lpg_Default:=vbNullString _
                                            , lpg_ReturnedString:=strBuffer _
                                            , nSize:=Len(strBuffer) _
                                            , lpg_FileName:=vn_file _
                                             )
            sNames = Left$(strBuffer, lResult)
        
            If sNames <> vbNullString Then                                         ' If there were any names
                asNames = Split(sNames, vbNullChar)                      ' have them split into an array
                For i = LBound(asNames) To UBound(asNames)
                    sName = asNames(i)
                    If Len(sName) <> 0 Then
                        If Not dctNames.Exists(sName) Then
                            mDct.DctAdd add_dct:=dctNames, add_key:=sName, add_item:=sName, add_seq:=seq_ascending
                        End If
                    End If
                Next i
            End If
            
        Next v
    End If
        
    Set ValueNames = dctNames

xt: Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function Values( _
                 ByVal vl_file As String, _
        Optional ByVal vl_section As String = vbNullString) As Dictionary
' -----------------------------------------------------------------------
' Returns a Dictionary with value name as key and value as item - of all
' sections in a file or of a specific one when a section is provided.
' -----------------------------------------------------------------------
    Dim dctValueNames   As Dictionary
    Dim dctValues       As New Dictionary
    Dim vn              As Variant
    
    Set dctValueNames = mFl.ValueNames(vn_file:=vl_file, vn_section:=vl_section)
    For Each vn In dctValueNames
        If Not dctValues.Exists(vn) _
        Then mDct.DctAdd add_dct:=dctValues, add_key:=vn, add_item:=mFl.Value(vl_file:=vl_file, vl_section:=vl_section, vl_value_name:=vn)
    Next vn
    Set Values = dctValues
    
End Function

Public Sub SectionsCopy(ByVal sc_section_names As Variant, _
                        ByVal sc_file_from As String, _
                        ByVal sc_file_to As String, _
               Optional ByVal sc_replace As Boolean = False)
' ----------------------------------------------------------
'
' ----------------------------------------------------------
    Const PROC = "SectionCopy"
    
    On Error GoTo eh
    Dim dctSections As Dictionary
    Dim i           As Long
    Dim cll         As New Collection
    
    Select Case VarType(sc_section_names)
        Case Is >= vbArray
            For i = LBound(sc_section_names) To UBound(sc_section_names)
                cll.Add sc_section_names(i)
            Next i
            SectionsCopy sc_file_from:=sc_file_from, sc_file_to:=sc_file_to, sc_section_names:=cll, sc_replace:=sc_replace
            GoTo xt
        Case vbString
            For i = LBound(Split(sc_section_names, ",")) To UBound(Split(sc_section_names, ","))
                cll.Add Split(sc_section_names, ",")(i)
            Next i
            SectionsCopy sc_file_from:=sc_file_from, sc_file_to:=sc_file_to, sc_section_names:=cll, sc_replace:=sc_replace
            GoTo xt
        Case vbObject
            Select Case TypeName(sc_section_names)
                Case "Dictionary"
                    For i = 0 To sc_section_names.Count
                        cll.Add sc_section_names.Items()(i)
                    Next i
                    SectionsCopy sc_file_from:=sc_file_from, sc_file_to:=sc_file_to, sc_section_names:=cll, sc_replace:=sc_replace
                    GoTo xt
                Case "Collection"
                    Set dctSections = SectionsGet(sg_file:=sc_file_from, sg_section_names:=sc_section_names)
                    If sc_replace Then mFl.SectionsRemove sr_file:=sc_file_to, sr_section_names:=sc_section_names
                    SectionsLet sl_sections:=dctSections, sl_file:=sc_file_to
                    GoTo xt
                Case Else: GoTo xt
            End Select
    End Select

xt: Set cll = Nothing
    Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Public Sub SectionReplace()

End Sub

Public Sub SectionMove()

End Sub
