Attribute VB_Name = "mExport"
Option Explicit
' ----------------------------------------------------------------------------
' Standard-Module mExport: Services specifically for the export of changed or
' ======================== all components.
'
' Public services:
' ----------------
' All               Exports all VBComponentnts whether the code has changed
'                   or not
' ChangedComponents Exports all VBComponents of which the code has changed,
'                   i.e. a temporary Export-File differs from the
'                       regular Export-File (of the previous code change).
' ExpFileFolderPath Returns a serviced Workbook's path for all Export-Files
'                   whereby the name of the folder is the current
'                   configured one (fefaulting  to 'source'). When no
'                   Export-Folder exists, one is created. In case an
'                   outdated export folder exists, i.e. one with an
'                   outdated name, this one is renamed instead.
' ----------------------------------------------------------------------------

Public Sub ChangedComponents(ByVal c_hosted As String)
' ----------------------------------------------------------------------------
' - Exports all components the code had been modified
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
    Dim fso         As New FileSystemObject
    Dim v           As Variant
    Dim vbc         As VBComponent
    Dim wbk         As Workbook
    Dim dct         As Dictionary
    Dim sRevNo      As String
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Prevent any action when the required preconditins are not met
    If Services.Denied(mCompManClient.SRVC_EXPORT_CHANGED) Then GoTo xt
    
    Set wbk = Services.Serviced
    Set dct = Comps.All
    For Each v In dct
        Set Comp = dct(v)
        Set vbc = Comp.VBComp
        If Not Services.IsRenamedByCompMan(vbc.Name) Then
            With Comp
                Services.ServicedItem = vbc
                .Wrkbk = Services.Serviced
                .CompName = vbc.Name
                .VBComp = vbc
                    Select Case .KindOfComp
                        Case mCompMan.enCommCompHosted
                            If .Changed Then
                                '~~ The current Export-File is not identicall with a temporary export's Export-File
                                sRevNo = .RevisionNumber
                                .Export
                                With Services
                                    .NoOfItemsServiced = .NoOfItemsServiced + 1
                                    .ServicedItemLogEntry "Modified Common Component hosted: e x p o r t e d !"
                                    .ServicedItemLogEntry "Modified Common Component hosted: Revision Number increased from " & sRevNo & " to " & Comp.RevisionNumber
                                    .ServicedItemLogEntry "Modified Common Component hosted: Export-File copied to " & wsConfig.FolderCommonComponentsPath
                                End With
                            Else
                                With Services
                                    .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                                    .ServicedItemLogEntry "Unchanged Common Component hosted ignored"
                                End With
                            End If
                    Case mCompMan.enCommCompUsed
                            If .Changed Then
                                '~~ The current Export-File is not identicall with a temporary export's Export-File
                                sRevNo = .RevisionNumber
                                .Export
                                With Services
                                    .NoOfItemsServiced = .NoOfItemsServiced + 1
                                    .ServicedItemLogEntry "Modified Common Component used: e x p o r t e d !"
                                    .ServicedItemLogEntry "Modified Common Component used: Revision Number increased from " & sRevNo & " to " & Comp.RevisionNumber
                                    .ServicedItemLogEntry "Modified Common Component used: Export-File copied to " & wsConfig.FolderCommonComponentsPath
                                End With
                            Else
                                With Services
                                    .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                                    .ServicedItemLogEntry "Unchanged Common Component used"
                                End With
                            End If
                                            
                    Case Else
                        '~~ Neither a hosted nor a used Common Component
                        If .Changed Then
                            .Export
                            With Services
                                .NoOfItemsServiced = .NoOfItemsServiced + 1
                                .NoOfItemsServicedNames = vbc.Name
                                .ServicedItemLogEntry "Modified VBComponent e x p o r t e d !"
                            End With
                        Else
                            With Services
                                .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                                .ServicedItemLogEntry "Unchanged VBComponent ignored"
                            End With
                        End If
                End Select
            End With
            Set Comp = Nothing
        Else
            With Services
                .NoOfItemsIgnored = .NoOfItemsIgnored + 1
            End With
        End If
        Services.DsplyProgress "exported"
        Set Comps = Nothing
    Next v

    With Services
        .DsplyProgress "exported"
        .RemoveTempRenamed
        .TempExportFolderRemove
    End With
    
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
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Rename the export folder when the one last used is no longe the one currently configured
    sExpFldrCurrentPath = mExport.ExpFileFolderPath(Services.Serviced)
    sExpFldrCurrentName = wsConfig.FolderExport
    sExpFldrRecentName = CompManDat.RecentlyUsedExportFolder
    If sExpFldrRecentName = vbNullString Then
        CompManDat.RecentlyUsedExportFolder = sExpFldrCurrentName
        sExpFldrRecentName = sExpFldrCurrentName
    End If
    If sExpFldrRecentName <> sExpFldrCurrentName Then
        sExpFldrRecentPath = Replace(sExpFldrCurrentPath, "\" & sExpFldrCurrentName, "\" & sExpFldrRecentName)
        fso.GetFolder(sExpFldrRecentPath).Name = sExpFldrCurrentName
        CompManDat.RecentlyUsedExportFolder = sExpFldrCurrentName
    End If
    
    '~~ Remove all Export-Files not corresponding to an existing VBComponet
    With fso
        For Each fl In .GetFolder(sExpFldrCurrentPath).Files
            Select Case .GetExtensionName(fl.Path)
                Case "bas", "cls", "frm", "frx"
                    If Not mComp.Exists(.GetBaseName(fl), Services.Serviced) Then
                        sExpFileName = .GetFileName(fl.Path)
                        .DeleteFile fl
                        LogServiced.Entry "Obsolete Export-File '" & sExpFileName & "' deleted (VBComponent no longer exists)"
                    End If
            End Select
        Next fl
    End With
        
xt: Set fso = Nothing
    Set fl = Nothing
    mBasic.EoP ErrSrc(PROC)
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
                Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "'" & s & "' is not the FullName of an existing Workbook!"
                sPathParent = .GetParentFolderName(s)
                sPath = sPathParent & "\" & wsConfig.FolderExport
            Case Else
                Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The required information about the concerned Workbook is neither provided as a Workbook object nor as a string identifying an existing Workbooks FullName"
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

