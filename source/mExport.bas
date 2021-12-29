Attribute VB_Name = "mExport"
Option Explicit

Public Sub All()
' --------------------------------------------------------------
' Standard-Module mExport
'
' Public serviced:
' - All                 Exports all VBComponentnts whether the code
'                       has changed or not
' - ChangedComponents   Exports all VBComponents of which the code
'                       has changed, i.e. a temporary Export-File
'                       differs from the regular Export-File (of
'                       the previous code change).
'
' --------------------------------------------------------------
    Const PROC = "All"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim sStatus As String
    Dim Comp    As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prevent any action when the required preconditins are not met
    If mService.Denied(PROC) Then GoTo xt
    
    sStatus = Log.Service

    '~~ Remove any obsolete Export-Files within the Workbook folder
    '~~ I.e. of no longer existing VBComponents or at an outdated location
    CleanUpObsoleteExpFiles
    
    If mMe.IsAddinInstnc _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
    
    For Each vbc In mService.Serviced.VBProject.VBComponents
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = mService.Serviced
            .CompName = vbc.Name ' this assignment provides the name for the export file
            '~~ Only export if it is not a component renamed by CompMan which is a left over
            If InStr(.CompName, RENAMED_BY_COMPMAN) <> 0 Then
                .Wrkbk.VBProject.VBComponents.Remove .VBComp
            Else
                .Export
            End If
        End With
        Set Comp = Nothing
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function FilesDiffer(ByVal fd_exp_file_1 As File, _
                            ByVal fd_exp_file_2 As File) As Boolean
' --------------------------------------------------------------
' Returns TRUE when file 1 fdiffers from file 2 whereby case
' differences and empty lines are ignored.
' This function guarantees a uniform comparison of export files
' throughout CompMan.
' --------------------------------------------------------------
    FilesDiffer = mFile.Differs(fd_file1:=fd_exp_file_1 _
                              , fd_file2:=fd_exp_file_2 _
                              , fd_stop_after:=1 _
                              , fd_ignore_empty_records:=True _
                              , fd_compare:=vbTextCompare).Count <> 0
                            
End Function

Public Function FilesDifference(ByVal fd_exp_file_1 As File, _
                                ByVal fd_exp_file_2 As File) As Dictionary
' --------------------------------------------------------------
' Returns a Dictionary with either 0 items when file 1 equals
' file 2 or with one item when the two files differ whereby
' case differences and empty lines are ignored.
' This function guarantees a uniform comparison of export files
' throughout CompMan.
' --------------------------------------------------------------
    Set FilesDifference = mFile.Differs(fd_file1:=fd_exp_file_1 _
                                      , fd_file2:=fd_exp_file_2 _
                                      , fd_stop_after:=1 _
                                      , fd_ignore_empty_records:=True _
                                      , fd_compare:=vbTextCompare)
                            
End Function

Public Function FilesDifferencesDisplay( _
          ByVal fd_exp_file_left_full_name As String, _
          ByVal fd_exp_file_right_full_name As String, _
          ByVal fd_exp_file_left_title As String, _
          ByVal fd_exp_file_right_title As String) As Boolean
' ----------------------------------------------------------------------------
' Displays the differences between export file 1 and 2 by means of WinMerge!
' Note: CompMan ignores any differences caused by empty code lines or case.
'       When a difference is displayed it is thus not because of this kind of
'       differneces but of others. Unfortunately it depends on the installed
'       WinMerge's set option wether or not these kind of differences are
'       displayed.
' ----------------------------------------------------------------------------
    Const PROC = "FilesDifferencesDisplay"
    
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
        
    If Not fso.FileExists(fd_exp_file_left_full_name) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & fd_exp_file_left_full_name & """ does not exist!"
    
    If Not fso.FileExists(fd_exp_file_right_full_name) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & fd_exp_file_right_full_name & """ does not exist!"
    
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & fd_exp_file_left_title & DQUOTE & _
               " /dr " & DQUOTE & fd_exp_file_right_title & DQUOTE & " " & _
               """" & fd_exp_file_left_full_name & """" & " " & _
               """" & fd_exp_file_right_full_name & """"
    
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        FilesDifferencesDisplay = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AllComps() As Dictionary
    Dim wb      As Workbook
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    Dim dct     As Dictionary
    
    Set wb = mService.Serviced
    With wb.VBProject
        For Each vbc In .VBComponents
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wb
                .CompName = vbc.Name
                Set .VBComp = vbc
            End With
            mDct.DctAdd add_dct:=dct, add_key:=vbc.Name, add_item:=Comp, add_order:=order_bykey
        Next vbc
    End With
    Set AllComps = dct
    
End Function

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
    Dim vbc1        As VBComponent
    Dim lComponents As Long
    Dim lRemaining  As Long
    Dim lExported   As Long
    Dim sExported   As String   ' comma separated exported component names
    Dim bUpdated    As Boolean
    Dim lUpdated    As Long
    Dim sUpdated    As String
    Dim sMsg        As String
    Dim fso         As New FileSystemObject
    Dim v           As Variant
    Dim Comps       As clsComps
    Dim dctChanged  As Dictionary
    Dim Comp        As clsComp
    Dim RawComp     As clsRaw
    Dim dctAll      As Dictionary
    Dim wb          As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    If mService.Denied(PROC) Then GoTo xt
    
    Set Stats = New clsStats
    Set Comps = New clsComps
    
    '~~ Remove any obsolete Export-Files within the Workbook folder
    '~~ I.e. of no longer existing VBComponents or at an outdated location
    CleanUpObsoleteExpFiles
    
    Set dctAll = AllComps
    lComponents = dctAll.Count
    Stats.Count sic_comps_total, lComponents
    lRemaining = lComponents
    If Log Is Nothing Then Set Log = New clsLog
    Log.Service = ErrSrc(PROC)
    
    For Each v In dctAll
        Set Comp = dctAll(v)
        With Comp
            '~~ Any 'left-over' component renamed by CompMan is removed
            If InStr(.CompName, RENAMED_BY_COMPMAN) <> 0 Then
                .Wrkbk.VBProject.VBComponents.Remove .VBComp
                GoTo next_v
            End If
            Log.ServicedItem = .VBComp
            
            Select Case .KindOfComp
                Case enCommCompUsed
                    '~~ --------------------------------------------------------------------------
                    '~~ The code of a used Common Component is only regarded changed when it
                    '~~ differs from the most recent ExportFile of the hoste raw's Export File.
                    '~~ --------------------------------------------------------------------------
                    If .Changed Then
                        '~~ The current Export File differs from a temporary export file. This may have
                        '~~ reasons: either the code of this used Common Component had been modified
                        '~~ or the new version of the component had already been imported manually.
                        If FilesDiffer(fd_exp_file_1:=fso.GetFile(.ExpFileTempFullName) _
                                     , fd_exp_file_2:=.RawMostRecentExpFile) Then
                            '~~ The used Common Component's code has been updated antidated
                            '~~ (directly transferred from the hosting Workbook/VB-Project within the VB-Editior)
                            Log.Entry = "The used Common Component's code had been updated antidated manually!"
                            .Export
                            mComCompsUsed.RevisionNumber(.CompName) = .RawMostRecentRevisionNumber
                            lExported = lExported + 1
                            If lExported = 1 _
                            Then sExported = .CompName _
                            Else sExported = sExported & ", " & .CompName
                        Else
                            Log.Entry = "The used Common Component's code had been modified in this VB-Project!"
                            DisplayRawCloneModificationWarning Comp.CompName
                            .Export
                            lExported = lExported + 1
                            If lExported = 1 _
                            Then sExported = .CompName _
                            Else sExported = sExported & ", " & .CompName
                        End If
                    End If
                Case enCommCompHosted
                    If .Changed Then
                        Log.Entry = "The hosted raw's code has been modified! " & _
                                    "(temporary Export-File differs from last changes Export-File)"
                        .Export
                        Log.Entry = "Exported to '" & .ExpFileFullName & "'"
                        mComCompsHosted.RevisionNumberIncrease .CompName
                        mComCompsHosted.ExpFileFullName(.CompName) = .ExpFileFullName
                        mComCompsSaved.RevisionNumber(.CompName) = mComCompsHosted.RevisionNumber(.CompName)
                        mComCompsSaved.ExpFile(.CompName) = .ExpFile ' maintain a copy as source for all using VB-Projects
                        Log.Entry = "Export file updated in folder '" & mComCompsSaved.ComCompsFile & "'"

                        lExported = lExported + 1
                        If lExported = 1 _
                        Then sExported = .CompName _
                        Else sExported = sExported & ", " & .CompName
                        mComCompsSaved.ExpFileFullName(.CompName) = .ExpFileFullName
                    Else
                        If mComCompsHosted.RevisionNumber(.CompName) = vbNullString Then
                            mComCompsHosted.RevisionNumberIncrease .CompName                               ' initial setting
                            mComCompsSaved.Update comp_name:=.CompName, exp_file:=.ExpFile                                           ' update in case appropriate
                            mComCompsSaved.RevisionNumber(.CompName) = mComCompsHosted.RevisionNumber(.CompName)
                            Log.Entry = "Export file in folder '" & mComCompsSaved.ComCompsFile & "' updated"
                        End If
                    End If
                Case enInternal
                    If .Changed Then
                        .Export
                        lExported = lExported + 1
                        If lExported = 1 _
                        Then sExported = .CompName _
                        Else sExported = sExported & ", " & .CompName
                    End If
                Case Else: Stop
            End Select

next_v:
            lRemaining = lRemaining - 1
            sMsg = Log.Service
            sMsg = sMsg & lExported & " of " & lComponents & " changed"
            If lExported > 0 _
            Then sMsg = sMsg & " (" & sExported & ")"
            sMsg = sMsg & " " & String(lRemaining, ".")
            If Len(sMsg) > 255 Then
                sMsg = Left(sMsg, 251) & " ..."
            End If
            Application.StatusBar = sMsg
        End With
        Set Comp = Nothing
    Next v
    Application.StatusBar = vbNullString: Application.StatusBar = sMsg
    
    
xt: Set dctHostedRaws = Nothing
    Set Comp = Nothing
    Set RawComp = Nothing
    Set Log = Nothing
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
    
    sExp = mCompMan.ExpFileFolderPath(mService.Serviced) ' the current specified Export-Folder

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

Private Sub DisplayRawCloneModificationWarning(ByVal comp_name As String)
    Dim Msg As mMsg.TypeMsg
    Msg.Section(1).Text.Text = "The Used Common Component '" & comp_name & "' may heve been modified in this Workbook"
    With Msg.Section(2)
        .Label.Text = "1. possible reason:"
        .Text.Text = "It is true! In order to provide a quick fix the code had been modified - and it is clear " & _
                     "that this modification will be reverted with the Workbook's next open."
    End With
    With Msg.Section(3)
        .Label.Text = "2. possible reason:"
        .Text.Text = "The code of the Common Component '" & comp_name & "' had been modified within its hosting Workbook " & _
                     "but this modification has yet not been saved i.e. exported. Instead the modified component had directly " & _
                     "been transferred from the raw Workbook's VB-Project in this one via the VB-Project-Editor." & vbLf & _
                     "The update of the used Common Component in this VP-Project had thus been antidated."
    End With
    
    mMsg.Dsply dsply_title:="Used Common Component indicated changed!" _
             , dsply_msg:=Msg
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mExport." & sProc
End Function

