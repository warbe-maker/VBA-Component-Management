Attribute VB_Name = "mExport"
Option Explicit

Public Sub All()
' ----------------------------------------------------------------------------
' Standard-Module mExport
'
' Public serviced:
' - All                 Exports all VBComponentnts whether the code has
'                       changed or not
' - ChangedComponents   Exports all VBComponents of which the code has
'                       changed, i.e. a temporary Export-File differs from the
'                       regular Export-File (of the previous code change).
'
' ----------------------------------------------------------------------------
    Const PROC = "All"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim sStatus     As String
    Dim Comp        As clsComp
    Dim Comps       As New clsComps
    Dim dctAll      As Dictionary
    Dim v           As Variant
    Dim wb          As Workbook
    Dim lAll        As Long
    Dim lExported   As Long
    Dim lRemaining  As String
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied Then GoTo xt

    '~~ Remove any obsolete Export-Files within the Workbook folder
    '~~ I.e. of no longer existing VBComponents or at an outdated location
    CleanUpObsoleteExpFiles
    
    If mMe.IsAddinInstnc _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
    
    Set wb = mService.Serviced
    With wb.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        
        For Each vbc In .VBComponents
            If Not mService.IsRenamedByCompMan(vbc.Name) Then
                Set Comp = New clsComp
                With Comp
                    .Wrkbk = mService.Serviced
                    .CompName = vbc.Name
                    '~~ Only export if it is not a component renamed by CompMan which is a left over
                    .Export
                    lExported = lExported + 1
                    lRemaining = lRemaining - 1
                End With
                Set Comp = Nothing
                mService.Progress p_service:=sStatus _
                                , p_result:=lExported _
                                , p_of:=lAll _
                                , p_op:="exported" _
                                , p_dots:=String(lRemaining, ".")
            End If
        Next vbc
    End With

    mService.RemoveTempRenamed
    
xt: Set Comps = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function AppErr(ByVal app_err_no As Long) As Long
' ----------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub ChangedComponents()
' ----------------------------------------------------------------------------
' Exports all components the code had been modified, detected by the
' comparison of a temporary export file with last modification's export file.
' Removes all Export Files of components which do not longer exist or exist in
' another but the current configured export folder. When a Used Common
' Component had been modified a due warning message is registered, displayed
' when the modification is reverted along with the next Workbook open.
' This se4rvice is exclusively performed/triggered by the Workbook_BeforeSave
' event.
' ----------------------------------------------------------------------------
    Const PROC = "ChangedComponents"
    
    On Error GoTo eh
    Dim Comp        As clsComp
    Dim Comps       As New clsComps
    Dim dctAll      As Dictionary
    Dim fso         As New FileSystemObject
    Dim lAll        As Long
    Dim lExported   As Long
    Dim lRemaining  As String
    Dim sExported   As String
    Dim sStatus     As String
    Dim v           As Variant
    Dim vbc         As VBComponent
    Dim wb          As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied Then GoTo xt
    sStatus = Log.Service

    '~~ Remove any obsolete Export-Files within the Workbook folder
    '~~ I.e. of no longer existing VBComponents or at an outdated location
    CleanUpObsoleteExpFiles
        
    Set wb = mService.Serviced
    Set dctAll = mService.AllComps(wb, sStatus)
    
    With wb.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        
        For Each v In dctAll
            Set vbc = dctAll(v)
            If Not mService.IsRenamedByCompMan(vbc.Name) Then
                Set Comp = New clsComp
                With Comp
                    Set .Wrkbk = mService.Serviced
                    .CompName = vbc.Name
                    Log.ServicedItem = vbc
                    Set .VBComp = vbc
                    Select Case .KindOfComp
                        Case enCommCompHosted
                            If .Changed Then
                                Log.Entry = "Hosted Raw Common Component code modified"
                                .Export
                                .RevisionNumberIncrease
                                .CopyExportFileToCommonComponentsFolder
                                sExported = sExported & vbc.Name & ", "
                                lExported = lExported + 1
                                Log.Entry = "Hosted Raw Common Component code modified"
                                Log.Entry = "Exported, Revision Number increased, Export File copied to 'Common Components Folder'"
                            ElseIf Not mComCompsRawsSaved.SavedExpFileExists(.CompName) Then
                                .CopyExportFileToCommonComponentsFolder ' ensure completenes
                                Log.Entry = "Hosted Raw Common Component unchanged"
                            End If
                        Case enCommCompUsed
                            If .Changed Then
                                '~~ A warning will be displayed when the modification is about to be reverted
                                '~~ when the component is updated at Workbook open
                                .DueModificationWarning = True
                                Log.Entry = "Used Common Component code modified (due revert allert registered)!"
                                .Export
                                sExported = sExported & vbc.Name & ", "
                                lExported = lExported + 1
                                Log.Entry = "Modified component exported (due modification warning set for update)"
                            Else
                                Log.Entry = "Used Common Component code unchanged"
                            End If
                        Case Else
                            If .Changed Then
                                Log.Entry = "Other component code modification exported"
                                .Export
                                sExported = sExported & vbc.Name & ", "
                                lExported = lExported + 1
                                Log.Entry = "Modification exported"
                            Else
                               Log.Entry = "Other component unchanged"
                            End If
                    End Select
                End With
                Set Comp = Nothing
            End If
            lRemaining = lRemaining - 1
            Application.StatusBar = _
            mService.Progress(p_service:=sStatus _
                            , p_result:=lExported _
                            , p_of:=lAll _
                            , p_op:="exported" _
                            , p_comps:=sExported _
                            , p_dots:=lRemaining _
                             )
            Set Comps = Nothing
        Next v
    End With

    Application.StatusBar = vbNullString
    Application.StatusBar = _
    mService.Progress(p_service:=sStatus _
                    , p_result:=lExported _
                    , p_of:=lAll _
                    , p_op:="exported" _
                    , p_comps:=sExported _
                     )
            
    mService.RemoveTempRenamed
    
xt: mBasic.EoP ErrSrc(PROC)
    Set fso = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CleanUpObsoleteExpFiles()
' ---------------------------------------------------------------------------
' - Deletes all Export-Files for which the corresponding component not or no
'   longer exists.
' - Delete all Export-Files in another but the current Export-Folder
' ---------------------------------------------------------------------------
    Const PROC = "CleanUpObsoleteExpFiles"
    
    On Error GoTo eh
    Dim cll     As Collection
    Dim fso     As New FileSystemObject
    Dim fl      As File
    Dim v       As Variant
    Dim Comp    As New clsComp
    Dim sExp    As String
    Dim fo      As Folder
    Dim fosub   As Folder
    
    sExp = mExport.ExpFileFolderPath(mService.Serviced) ' the current specified Export-Folder

    '~~ Cleanup of any Export-Files residing outside the specified 'Export-Folder'
    Set cll = New Collection
    cll.Add fso.GetFolder(mService.Serviced.Path)
    Do While cll.Count > 0
        Set fo = cll(1): cll.Remove 1 'get folder and dequeue it
        If fo.Path <> sExp Then
            For Each fosub In fo.SubFolders
                cll.Add fosub ' enqueue it
            Next fosub
            If fo.ParentFolder = mService.Serviced.Path Or fo.Path = mService.Serviced.Path Then
                '~~ Cleanup is done only in the Workbook-folder and any direct sub-folder
                '~~ Folders in sub-folders are exempted.
                For Each fl In fo.Files
                    Select Case fso.GetExtensionName(fl.Path)
                        Case "bas", "cls", "frm", "frx"
                            fso.DeleteFile (fl)
                    End Select
                Next fl
            End If
        End If
    Loop
    Set cll = Nothing
    
    '~~ Collect all outdated Export-Files in the specified Export-Folder
    Set cll = New Collection
    For Each fl In fso.GetFolder(sExp).Files
        Select Case fso.GetExtensionName(fl.Path)
            Case "bas", "cls", "frm", "frx"
                If Not mComp.Exists(mService.Serviced, fso.GetBaseName(fl)) Then cll.Add fl.Path
        End Select
    Next fl
        
    '~~ Remove all obsolete Export-Files
    With fso
        For Each v In cll
            .DeleteFile v
            Log.Entry = "Export-File obsolete (deleted because component no longer exists)"
        Next v
    End With
    
    RemoveEmptyFolders ref_folder:=mService.Serviced.Path, ref_sub_folders:=False
    
xt: Set cll = Nothing
    Set fso = Nothing
    Set fo = Nothing
    Set fosub = Nothing
    Set fl = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mExport." & sProc
End Function

Public Function ExpFileFolderPath(ByVal v As Variant) As String
' --------------------------------------------------------------------------
' Returns a Workbook's path for all Export Files whereby the name of the
' folder is the the configured one which defaults to 'source'.
' ----------------------------------------------------------------------------
    Const PROC = "ExpFileFolderPath"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim wb      As Workbook
    Dim s       As String
    Dim sPath   As String
    
    With fso
        Select Case TypeName(v)
            Case "Workbook"
                Set wb = v
                sPath = wb.Path & "\" & mConfig.FolderExport
            Case "String"
                s = v
                If Not .FileExists(s) _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "'" & s & "' is not the FullName of an existing Workbook!"
                sPath = .GetParentFolderName(s) & "\" & mConfig.FolderExport
            Case Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The required information about the concerned Workbook is neither provided as a Workbook object nor as a string identifying an existing Workbooks FullName"
        End Select
        If Not .FolderExists(sPath) Then .CreateFolder sPath
    End With
    
xt: ExpFileFolderPath = sPath
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub RemoveEmptyFolders(ByVal ref_folder As String, _
                      Optional ByVal ref_sub_folders As Boolean = False)
' --------------------------------------------------------------------------
' Removes any empty folder in the folder (ref_folder)
' --------------------------------------------------------------------------
    Const PROC = "RemoveEmptyFolders"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim iFolders    As Long
    Dim oFile       As File
    Dim oFolder     As Folder
    Dim oSubFolder  As Folder
    
    ref_folder = Replace(ref_folder & "\", "\\", "\") ' add a possibly missing trailing \
    Set oFolder = fso.GetFolder(ref_folder)
    If Not fso.FolderExists(ref_folder) Then GoTo xt
    
    For Each oSubFolder In oFolder.SubFolders
        '~~ Loop through all subfolders
        iFolders = oSubFolder.SubFolders.Count
        If oSubFolder.Files.Count = 0 And iFolders = 0 Then
            '~~ Remove the folder when it has no files and no sub-folders
            RmDir oSubFolder.Path
        End If
        '~~ Recursively repeat if there are any subfolders within the subfolder
        If ref_sub_folders Then
            If iFolders <> 0 Then RemoveEmptyFolders (ref_folder & oSubFolder.Name)
        End If
    Next

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

