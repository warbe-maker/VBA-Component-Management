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
    If mService.Denied(PROC) Then GoTo xt
    sStatus = Log.Service

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
            If Not mService.IsRenamedByCompMan(vbc.name) Then
                Set Comp = New clsComp
                With Comp
                    .Wrkbk = mService.Serviced
                    .CompName = vbc.name
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
' Exclusively performed/trigered by the Workbook_BeforeSave event. Exports
' all changed components. I.e. any code change (detected by the comparison of
' a temporary export file with the current export file) is exported. Outdated
' Export Files (of components which do not longer exist or exist in another
' but the currently used export folder) are removed. Because any code
' modification in a raw clone (a used common conponent managed by CompMan
' services) is reverted along with the next open
' - Clone code modifications update the raw code when confirmed by the
'   user
' ----------------------------------------------------------------------------
    Const PROC = "ChangedComponents"
    
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
    Dim sExported   As String
    Dim fso         As New FileSystemObject
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied(PROC) Then GoTo xt
    sStatus = Log.Service

    '~~ Remove any obsolete Export-Files within the Workbook folder
    '~~ I.e. of no longer existing VBComponents or at an outdated location
    CleanUpObsoleteExpFiles
        
    Set wb = mService.Serviced
    With wb.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        
        For Each vbc In .VBComponents
            If Not mService.IsRenamedByCompMan(vbc.name) Then
                Set Comp = New clsComp
                With Comp
                    Set .Wrkbk = mService.Serviced
                    .CompName = vbc.name
                    Set .VBComp = vbc
                    Select Case .KindOfComp
                        Case enCommCompHosted
                            If .Changed Then
                                .Export
                                .RevisionNumberIncrease
                                .CopyExportFileToCommonComponentsFolder
                                sExported = sExported & vbc.name & ", "
                                lExported = lExported + 1
                            ElseIf Not mComCompsRawsSaved.ExpFileExists(.CompName) Then
                                .CopyExportFileToCommonComponentsFolder ' ensure completenes
                            End If
                        Case enCommCompUsed
                            If UsedCommonComponent(Comp) Then
                                sExported = sExported & vbc.name & ", "
                                lExported = lExported + 1
                            End If
                        Case Else
                            If .Changed Then
                                .Export
                                sExported = sExported & vbc.name & ", "
                                lExported = lExported + 1
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
        Next vbc
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

Private Function UsedCommonComponent(ByRef cucc_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Used Common Component (cucc_comp) had been exported.
' Processes a changed Used Common Component. I.e. a Used Common Component of
' which the current code differs from its last Export File.
' ----------------------------------------------------------------------------
    
    With cucc_comp
        If .Changed Then
            If .RevisionNumber = .Raw.RevisionNumber Then
                '~~ The Used Common Component compared with ist corresponding Raw
                '~~ is up-to-date but the code has been modified in this Workbook.
                '~~ This modification will be reverted with the next Workbook open.
                If .Raw.HostIsOpen Then
                    '~~ A code modification of the Raw, yet nbot saved/exported may have been
                    '~~ dragged to the Workbook using it. In this case a change is indicated but
                    '~~ the code may be 'already' identical
                    If .ExpFileTempsDiffer Then
                        CommCompModificationWarning _
                            cmod_comp_name:=.CompName _
                          , cmod_exp_file_full_name:=.ExpFileTempFullName _
                          , cmod_raw_exp_file_full_name:=.Raw.ExpFileTempFullName _
                          , cmod_diff_message:= _
                          "The code of the Used Common Component '" & mBasic.Spaced(.CompName) & "' had " & _
                          "been modified in this Workbook! This modification will be reverted with the next open!"
                    Else
                        '~~ Manually 'antidated' updated. When the open Raw is closed the Expost Files will become identical.
                    End If
                ElseIf .ExpFileTempAndRawExpFileDiffers Then
                    CommCompModificationWarning _
                        cmod_comp_name:=.CompName _
                      , cmod_exp_file_full_name:=.ExpFileTempFullName _
                      , cmod_raw_exp_file_full_name:=.Raw.ExpFileFullName _
                      , cmod_diff_message:= _
                      "The code of the Used Common Component '" & mBasic.Spaced(.CompName) & "' had " & _
                      "been modified in this Workbook! This modification will be reverted with the next open!"
                End If
                
                .Export
                UsedCommonComponent = True
            
            ElseIf .RevisionNumber < .Raw.RevisionNumber Then
                CommCompModificationWarning _
                    cmod_comp_name:=.CompName _
                  , cmod_exp_file_full_name:=.ExpFileTempFullName _
                  , cmod_raw_exp_file_full_name:=.Raw.ExpFileFullName _
                  , cmod_diff_message:= _
                  "Attention! The code of the Raw Common Component '" & mBasic.Spaced(.CompName) & "' " & _
                  "has been modified since the last open of this Workbook. Just in case the code " & _
                  "of the Used Common Component had also bee3n modified in this Workbook this " & _
                  "modification will be reverted with the next Workbook open."
            
            ElseIf .RevisionNumber > .Raw.RevisionNumber Then
                '~~ The content of the saved 'Common Components' folder is outdated. Something which likely never happens.
                '~~ However, when the Common Componernts folder represents a Github repo the clone apparently requres a fetch!
                mMsg.Box box_title:="Content of Common Component' folder is outdated" _
                       , box_msg:="The Used Common Component in this Workbook is more up-to-date than the available " & _
                                  "Raw Common Component in the saved folder 'Common Components'. Precise: The RevisionNumber " & _
                                  "of the Used Common Component is greater the the RevisionNumber of the Raw."
            End If
        
        ElseIf .ExpFilesDiffer Then
            '~~ Even when the Used Common Component not had been changed it may have been exported because of a previous detected change.
            '~~ When the RevisionNumbers are equal and the Export Files differ the Used Common Component must have been changed and
            '~~ this code modification will be reverted with the next Workbook open.
            CommCompModificationWarning _
                cmod_comp_name:="" _
              , cmod_exp_file_full_name:=.ExpFileTempFullName _
              , cmod_raw_exp_file_full_name:=.Raw.ExpFileFullName _
              , cmod_diff_message:="Attention! The code of the Used Common Component '" & mBasic.Spaced(.CompName) & "' " & _
                                   "has been modified!! This modification will be reverted with the next Workbook open."
        End If
    End With

End Function


Private Sub CleanUpObsoleteExpFiles()
' ------------------------------------------------------
' - Deletes all Export-Files for which the corresponding
'   component not or no longer exists.
' - Delete all Export-Files in another but the current
'   Export-Folder
' -------------------------------------------------------
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

Private Sub CommCompModificationWarning(ByVal cmod_comp_name As String, _
                                        ByVal cmod_exp_file_full_name, _
                                        ByVal cmod_raw_exp_file_full_name, _
                                        ByVal cmod_diff_message)
' ----------------------------------------------------------------------------
' Displays an information about a modification of a Used Common Component.
' The disaplay offers the option to display the code difference.
' ----------------------------------------------------------------------------
    Const PROC = "CommCompModificationWarning"
    
    On Error GoTo eh
    Dim msg         As mMsg.TypeMsg
    Dim cllBttns    As Collection
    Dim BttnDsply   As String
    
    BttnDsply = "Display" & vbLf & "code difference"
    mMsg.Buttons cllBttns, BttnDsply, vbLf, vbOKOnly
    With msg.Section(1)
        .Text.Text = cmod_diff_message
        .Text.FontColor = rgbRed
    End With
        
    Do
        Select Case mMsg.Dsply(dsply_title:="Used Common Component code modified" _
                             , dsply_msg:=msg _
                             , dsply_buttons:=cllBttns _
                              )
            Case BttnDsply
                mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=cmod_exp_file_full_name _
                                               , fd_exp_file_left_title:="Used Common Component: '" & cmod_exp_file_full_name & "'" _
                                               , fd_exp_file_right_full_name:=cmod_raw_exp_file_full_name _
                                               , fd_exp_file_right_title:="Raw Common Component: '" & cmod_raw_exp_file_full_name & "'"
    
            Case vbOK: Exit Do
        End Select
    Loop

xt: Exit Sub

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

