Attribute VB_Name = "mFile"
Option Explicit
Option Compare Text
Option Private Module
' --------------------------------------------------------------
' Standard  Module mFile
'           Common methods and functions regarding file objects.
'
' Methods:  Exists      Returns TRUE when the file exists
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
Private Const CONCAT = "||"

Public Function Exists(ByVal xst_file As Variant, _
              Optional ByRef xst_fso As File = Nothing, _
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
    Dim fl      As File
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

    If TypeOf xst_file Is File Then
        With New FileSystemObject
            On Error Resume Next
            sTest = xst_file.Name
            Exists = Err.Number = 0
            If Exists Then
                '~~ Return the existing file as File object
                Set xst_fso = .GetFile(xst_file.Path)
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
                        If InStr(fl.Name, sFile) <> 0 And left(fl.Name, 1) <> "~" Then
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

Public Function GetFile(ByVal gf_path As String) As File
    With New FileSystemObject
        Set GetFile = .GetFile(gf_path)
    End With
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
    Dim fso     As File
    Dim sFile   As String
    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    
    If Not Exists(ta_file, fso) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (vFile) does not exist!"
    
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
    Dim fso     As File
    Dim sFile   As String
    Dim i       As Long
    
    If Not Exists(td_file, fso) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The file object (td_file) does not exist!"
    
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
        
xt: Set ToDict = dct
    Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

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

End Function

Public Function sDiffer( _
                  ByVal dif_file1 As File, _
                  ByVal dif_file2 As File, _
         Optional ByVal dif_stop_after As Long = 1, _
         Optional ByVal dif_ignore_empty_records As Boolean = False, _
         Optional ByVal dif_lines As Variant) As Boolean
' -----------------------------------------------------------------------------
' Returns TRUE when the content of file (dif_file1) differs from the content in
' file (dif_file2). The comparison stops after (dif_stop_after) detected
' differences. The detected different lines are optionally returned (vResult).
' ------------------------------------------------------------------------------

    Dim a1      As Variant
    Dim a2      As Variant
    Dim vLines  As Variant

    a1 = mFile.ToArray(ta_file:=dif_file1, ta_exclude_empty_records:=dif_ignore_empty_records)
    a2 = mFile.ToArray(ta_file:=dif_file2, ta_exclude_empty_records:=dif_ignore_empty_records)
    vLines = mBasic.ArrayCompare(a1, a2, dif_stop_after)
    If mBasic.ArrayIsAllocated(arr:=vLines) Then
        sDiffer = True
    End If
    dif_lines = vLines
    
End Function

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

Public Sub Delete(ByVal v As Variant)

    Dim fl  As File

    With New FileSystemObject
        If TypeName(v) = "File" Then
            Set fl = v
            .DeleteFile fl.Path
        ElseIf TypeName(v) = "String" Then
            If .FileExists(v) Then
                .DeleteFile v
            End If
        End If
    End With
    
End Sub

Public Function Extension(ByVal vFile As Variant)

    With New FileSystemObject
        If TypeName(vFile) = "File" Then
            Extension = .GetExtensionName(vFile.Path)
        Else
            Extension = .GetExtensionName(vFile)
        End If
    End With

End Function

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
    ErrSrc = ThisWorkbook.Name & ": mFile." & sProc
End Function
