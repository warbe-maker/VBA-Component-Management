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
' Uses:     Common Components mBasic and mErrHndlr. Both are regarded available
'           through the CompMan Addin. when the Addin is not used, both have to
'           imported.
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' W. Rauschenberger, Berlin August 2019
' -----------------------------------------------------------------------------------

Public Function Exists(ByVal vFile As Variant, _
              Optional ByRef fso As File = Nothing, _
              Optional ByRef clt As Collection = Nothing) As Boolean
' ------------------------------------------------------------------
' Returns TRUE when the file (vFile) - which may be a file object
' or a file's full name - exists and furthermore:
' - when the file's full name ends with a wildcard * all
'   subfolders are scanned and any file which meets the criteria
'   is returned as File object in a collection (clt),
' - when the files's full name does not end with a wildcard * the
'   existing file is returned as a File object (fso).
' ----------------------------------------------------------------
Const PROC  As String = "Exists"    ' This procedure's name for the error handling and execution tracking
Dim sTest   As String
Dim sFile   As String
Dim fldr    As Folder
Dim sfldr   As Folder   ' Sub-Folder
Dim fl      As File
Dim sPath   As String
Dim queue   As Collection

    On Error GoTo on_error
    Exists = False
    Set clt = New Collection

    If TypeName(vFile) <> "File" And TypeName(vFile) <> "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The File (parameter vFile) for the File's existence check is neither a full path/file name nor a file object!"
    If Not TypeName(fso) = "Nothing" And Not TypeName(fso) = "File" _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided return parameter (fso) is not a File type!"
    If Not TypeName(clt) = "Nothing" And Not TypeName(clt) = "Collection" _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided return parameter (clt) is not a Collection type!"

    If TypeOf vFile Is File Then
        With New FileSystemObject
            On Error Resume Next
            sTest = vFile.Name
            Exists = Err.Number = 0
            If Exists Then
                '~~ Return the existing file as File object
                Set fso = .GetFile(vFile.Path)
                GoTo exit_proc
            End If
        End With
    ElseIf VarType(vFile) = vbString Then
        With New FileSystemObject
            sFile = Split(vFile, "\")(UBound(Split(vFile, "\")))
            If Not Right(sFile, 1) = "*" Then
                Exists = .FileExists(vFile)
                If Exists Then
                    '~~ Return the existing file as File object
                    Set fso = .GetFile(vFile)
                    GoTo exit_proc
                End If
            Else
                sPath = Replace(vFile, "\" & sFile, vbNullString)
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
                        If InStr(fl.Name, sFile) <> 0 And Left(fl.Name, 1) <> "~" Then
                            '~~ Return the existing file which meets the search criteria
                            '~~ as File object in a collection
                            clt.Add fl
                         End If
                    Next fl
                Loop
                If clt.Count > 0 Then Exists = True
            End If
        End With
    End If

exit_proc:
    Exit Function
    
on_error:
#If Debugging = 1 Then
    Debug.Print Err.Description: Stop: ' Resume
#End If
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Public Function GetFile(ByVal sPath As String) As File
    With New FileSystemObject
        Set GetFile = .GetFile(sPath)
    End With
End Function

Public Function ToArray(ByVal vFile As Variant) As String()
' ---------------------------------------------------------
' Returns the content of the file (vFile) - which may be
' provided as file object or full file name - as array
' by considering any kind of line break characters.
' ---------------------------------------------------------
Const PROC  As String = "ToArray"
Dim ts      As TextStream
Dim a       As Variant
Dim sSplit  As String
Dim fso     As File
Dim sFile   As String

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    If Not Exists(vFile, fso) _
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
    
    If sFile = vbNullString Then GoTo exit_proc
    
    '~~ Get the kind of line break used
    If InStr(sFile, vbCr) <> 0 Then sSplit = vbCr
    If InStr(sFile, vbLf) <> 0 Then sSplit = sSplit & vbLf
    
    '~~ Test stream to array
    a = Split(sFile, sSplit)
    
    '~~ Remove any leading or trailing empty items
    mBasic.ArrayTrimm a
    ToArray = a
    
exit_proc:
    EoP ErrSrc(PROC)
    Exit Function
    
on_error:
#If Debugging = 1 Then
    Debug.Print Err.Description: Stop: ' Resume
#End If
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Public Function ToDict(ByVal vFile As Variant) As Dictionary
' ----------------------------------------------------------
' Returns the content of the file (vFile) - which may be
' provided as file object or full file name - as Dictionary
' by considering any kind of line break characters.
' ---------------------------------------------------------
Const PROC  As String = "ToDict"
Dim ts      As TextStream
Dim a       As Variant
Dim dct     As New Dictionary
Dim sSplit  As String
Dim fso     As File
Dim sFile   As String
Dim i       As Long

    On Error GoTo on_error
    
    If Not Exists(vFile, fso) _
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
    
    If sFile = vbNullString Then GoTo exit_proc
    
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
        
exit_proc:
    Set ToDict = dct
    Exit Function
    
on_error:
#If Debugging = 1 Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Public Function SelectFile(Optional ByVal sInitPath As String = vbNullString, _
                           Optional ByVal sFilters As String = "*.*", _
                           Optional ByVal sFilterName As String = "File", _
                           Optional flResult As File) As Boolean
' -----------------------------------------------------------------------------
' When a file had been selected TRUE is returned and the selected file is
' returned as File object (flResult).
' -----------------------------------------------------------------------------
Dim fDialog As FileDialog
Dim v       As Variant

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Select a(n) " & sFilterName
        .InitialFileName = sInitPath
        .Filters.Clear
        For Each v In Split(sFilters, ",")
            .Filters.Add sFilterName, v
         Next v
         
        If .Show = -1 Then
            '~~ A fie had been selected
           With New FileSystemObject
            Set flResult = .GetFile(fDialog.SelectedItems(1))
            SelectFile = True
           End With
        End If
        '~~ When no file had been selected the flResult will be Nothing
    End With

End Function

Public Function Differ(ByVal f1 As File, _
                       ByVal f2 As File, _
              Optional ByVal lStopAfter As Long = 1) As Boolean
' ---------------------------------------------------------
' Returns TRUE when the content of file (f1) differs from
' the content in file (f2). The comparison stops after
' (lStopAfter) detected differences. The detected different
' lines are optionally returned (vResult).
' ----------------------------------------------------------
Dim a1      As Variant
Dim a2      As Variant
Dim vLines  As Variant

    a1 = mFile.ToArray(f1)
    a2 = mFile.ToArray(f2)
    vLines = mBasic.ArrayCompare(a1, a2, lStopAfter)
    If mBasic.ArrayIsAllocated(vLines) Then
        Differ = True
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mFile" & ">" & sProc
End Function
