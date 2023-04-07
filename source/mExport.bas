Attribute VB_Name = "mExport"
Option Explicit
' ----------------------------------------------------------------------------
' Standard-Module mExport: Services specifically for the export of changed or
' ------------------------ all components.
' Public services:
' - All                 Exports all VBComponentnts whether the code has
'                       changed or not
' - ChangedComponents   Exports all VBComponents of which the code has
'                       changed, i.e. a temporary Export-File differs from the
'                       regular Export-File (of the previous code change).
' - ExpFileFolderPath   Returns a serviced Workbook's path for all Export-
'                       Files whereby the name of the folder is the current
'                       configured one (fefaulting  to 'source'). When no
'                       Export-Folder exists, one is created. In case an
'                       outdated export folder exists, i.e. one with an
'                       outdated name, this one is renamed instead.
'
' ----------------------------------------------------------------------------

Public Sub All()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "All"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim Comp        As clsComp
    Dim Comps       As New clsComps
    Dim wbk         As Workbook
    Dim lAll        As Long
    Dim lExported   As Long
    Dim lRemaining  As String
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied(mCompManClient.SRVC_EXPORT_ALL) Then GoTo xt

    '~~ Remove any obsolete Export-Files within the Workbook folder
    '~~ I.e. of no longer existing VBComponents or at an outdated location
    Hskpng
    
    If mMe.IsAddinInstnc _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Add-in instance which is impossible for this operation!"
    
    Set wbk = mService.WbkServiced
    With wbk.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        
        For Each vbc In .VBComponents
            If Not mService.IsRenamedByCompMan(vbc.Name) Then
                Set Comp = New clsComp
                With Comp
                    .Wrkbk = mService.WbkServiced
                    .CompName = vbc.Name
                    '~~ Only export if it is not a component renamed by CompMan which is a left over
                    .Export
                    lExported = lExported + 1
                    lRemaining = lRemaining - 1
                End With
                Set Comp = Nothing
                mService.Progress p_result:=lExported _
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

Private Function AppErr(ByVal app_err_no As Long) As Long
' ----------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub ChangedComponents(ByVal c_hosted As String)
' ----------------------------------------------------------------------------
' - Exports all components the code had been modified, detected by the com-
'   parison of a temporary export file with last modification's export file.
' - Removes all Export Files of which the corresponding component no longer
'   exist.
' - Registers a due warning message when a Used Common Component had been
'   modified in the Workbook which uses but not hosts it.
' - Forwards (renames) and outdated Export-Folder name to the name currently
'   configured (Hskpng).
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
    Dim wbk         As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied(mCompManClient.SRVC_EXPORT_CHANGED) Then GoTo xt
    sStatus = mService.CurrentServiceStatusBar
    
    Hskpng                      ' forward outdated export folder and remove obsolete Export files
    mCompManDat.Hskpng c_hosted ' remove obsolete sections in CompMan.dat
    
    Set wbk = mService.WbkServiced
    Set dctAll = mService.AllComps(wbk)
    
    With wbk.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        
        For Each v In dctAll
            Set vbc = dctAll(v)
            If Not mService.IsRenamedByCompMan(vbc.Name) Then
                Set Comp = New clsComp
                With Comp
                    Set .Wrkbk = mService.WbkServiced
                    .CompName = vbc.Name
                    mService.Log.ServicedItem = vbc
                    Set .VBComp = vbc
                    Select Case .KindOfComp
                        Case mCompMan.enCommCompHosted
                            If .Changed Then
                                mService.Log.Entry = "Hosted Raw Common Component code modified"
                                .Export
                                mCompManDat.RawRevisionNumberIncrease v
                                mCommComps.SaveToCommonComponentsFolder .CompName, .ExpFile, .ExpFileFullName
                                mCompManDat.RegistrationState(.CompName) = enRegStateHosted
                                sExported = sExported & vbc.Name & ", "
                                lExported = lExported + 1
                                mService.Log.Entry = "Code modified of Hosted Raw Common Component"
                                mService.Log.Entry = "Exported, Revision Number increased, Export File copied to 'Common Components Folder'"
                            ElseIf Not mCommComps.SavedExpFileExists(.CompName) Then
                                mCommComps.SaveToCommonComponentsFolder .CompName, .ExpFile, .ExpFileFullName ' ensure completenes
                                mService.Log.Entry = "Unchanged Hosted Raw Common Component"
                                mCompManDat.RegistrationState(.CompName) = enRegStateHosted
                            End If
                        Case mCompMan.enCommCompUsed
                            If .Changed Then
                                '~~ A warning will be displayed when the modification is about to be reverted
                                '~~ when the component is updated at Workbook open
                                .DueModificationWarning = True
                                .Export
                                sExported = sExported & vbc.Name & ", "
                                lExported = lExported + 1
                                mService.Log.Entry = "Modified Used Common Component exported (due revert allert registered!)"
                            Else
                                mService.Log.Entry = "Unchanged Used Common Component"
                            End If
                            
                            If mCompManDat.CommCompUsedIsKnown(vbc.Name) Then
                            Else
                            
                            End If
                        
                        Case Else
                            If .Changed Then
                                .Export
                                sExported = sExported & vbc.Name & ", "
                                lExported = lExported + 1
                                mService.Log.Entry = "Modified code exported"
                            Else
                               mService.Log.Entry = "Unchanged component"
                            End If
                    End Select
                End With
                Set Comp = Nothing
            
            
            End If
            lRemaining = lRemaining - 1
            mService.DsplyStatus _
            mService.Progress(p_result:=lExported _
                            , p_of:=lAll _
                            , p_op:="exported" _
                            , p_comps:=sExported _
                            , p_dots:=lRemaining _
                             )
            Set Comps = Nothing
        Next v
    End With

    mService.DsplyStatus _
    mService.Progress(p_result:=lExported _
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

Private Sub Hskpng()
' ---------------------------------------------------------------------------
' - Forwards an outdated (last used) export folder to the one currently
'   configured
' - Deletes all Export-Files for which the corresponding component not or no
'   longer exists.
' ---------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim fl                  As File
    Dim sExpFldrCurrentName As String
    Dim sExpFldrRecentName  As String
    Dim sExpFldrCurrentPath As String
    Dim sExpFileName        As String
    Dim sExpFldrRecentPath  As String
    
    '~~ Rename the export folder when the one last used is no longe the one currently configured
    sExpFldrCurrentPath = mExport.ExpFileFolderPath(mService.WbkServiced)
    sExpFldrCurrentName = wsConfig.FolderExport
    sExpFldrRecentName = mCompManDat.RecentlyUsedExportFolder
    If sExpFldrRecentName = vbNullString Then
        mCompManDat.RecentlyUsedExportFolder = sExpFldrCurrentName
        sExpFldrRecentName = sExpFldrCurrentName
    End If
    If sExpFldrRecentName <> sExpFldrCurrentName Then
        sExpFldrRecentPath = Replace(sExpFldrCurrentPath, "\" & sExpFldrCurrentName, "\" & sExpFldrRecentName)
        fso.GetFolder(sExpFldrRecentPath).Name = sExpFldrCurrentName
        mCompManDat.RecentlyUsedExportFolder = sExpFldrCurrentName
    End If
    
    '~~ Remove all Export-Files not corresponding to an existing VBComponet
    With fso
        For Each fl In .GetFolder(sExpFldrCurrentPath).Files
            Select Case .GetExtensionName(fl.Path)
                Case "bas", "cls", "frm", "frx"
                    If Not mComp.Exists(.GetBaseName(fl), mService.WbkServiced) Then
                        sExpFileName = .GetFileName(fl.Path)
                        .DeleteFile fl
                        mService.Log.Entry = "Obsolete Export-File '" & sExpFileName & "' deleted (VBComponent no longer exists)"
                    End If
            End Select
        Next fl
    End With
        
xt: Set fso = Nothing
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
' ----------------------------------------------------------------------------
' Returns a Workbook's path for all Export Files whereby the name of the
' folder is the configured one (which defaults to 'source'). When no export
' folder exists, one is created. In case an outdated export folder exists,
' i.e. one with an outdated name, this one is renamed instead.
' ----------------------------------------------------------------------------
    Const PROC = "ExpFileFolderPath"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim wbk         As Workbook
    Dim s           As String
    Dim sPath       As String
    Dim fldExisting As Folder
    Dim sPathParent As String
    
    With fso
        Select Case TypeName(v)
            Case "Workbook"
                Set wbk = v
                sPathParent = wbk.Path
                sPath = sPathParent & "\" & wsConfig.FolderExport
            Case "String"
                s = v
                If Not .FileExists(s) _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "'" & s & "' is not the FullName of an existing Workbook!"
                sPathParent = .GetParentFolderName(s)
                sPath = sPathParent & "\" & wsConfig.FolderExport
            Case Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The required information about the concerned Workbook is neither provided as a Workbook object nor as a string identifying an existing Workbooks FullName"
        End Select
        If Not .FolderExists(sPath) Then
            '~~ When no 'Export' folder exists there may still be an outdated one of which the nmae had not already been changed.
            '~~ When an export folder with a different name already exists this one should be renamed.
            If AnExportFolderExists(sPathParent, fldExisting) Then
                fldExisting.Name = wsConfig.FolderExport
            Else
                .CreateFolder sPath
            End If
        End If
    End With
    
xt: ExpFileFolderPath = sPath
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AnExportFolderExists(ByVal oef_path As String, _
                                      ByRef oef_fld As Folder) As Boolean
' --------------------------------------------------------------------------
' When a folder exists with files *.bas, *.cls, or *.frm the function
' returns TRUE and the identified folder (oef_flf).
' --------------------------------------------------------------------------
    Const PROC = "AnExportFolderExists"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Dim fle As File
    
    For Each fld In fso.GetFolder(oef_path).SubFolders
        For Each fle In fld.Files
            Select Case fso.GetExtensionName(fle.Path)
                Case "bas", "cls", "frm"
                    Set oef_fld = fld
                    AnExportFolderExists = True
                    GoTo xt
            End Select
        Next fle
    Next fld
    
xt: Set fso = Nothing
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

