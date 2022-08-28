Attribute VB_Name = "mSync"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSync
'
' Provides all services and means for the synchronization of a Target- with
' its corresponding Source-Workbook.
'
' Note: All Run.... services are invoked via Application.Run
'
'
' W. Rauschenberger, Berlin June 2022
' ----------------------------------------------------------------------------
Private Const SHEET_CONTROL_CONCAT          As String = ": "                                ' Sheet-Shape concatenator
Private Const SYNC_COPY_SUFFIX              As String = "_Synched"                          ' name suffix for the sync target working copy

Public Const APP_RUN_ARG_BUTTON_CAPTION     As Long = 1
Public Const APP_RUN_ARG_SERVICING_WORKBOOK As Long = 2
Public Const APP_RUN_ARG_SERVICE            As Long = 3
Public Const APP_RUN_ARG_SERVICE_ARG1       As Long = 4
Public Const APP_RUN_ARG_SERVICE_ARG2       As Long = 5
Public Const APP_RUN_ARG_SERVICE_ARG3       As Long = 6
Public Const TITLE_SYNC_NAMES               As String = "Synchronize Range Names"           ' Identifies the synchronization dialog
Public Const TITLE_SYNC_RANGES              As String = "Synchronize Named Ranges"          ' Identifies the synchronization dialog
Public Const TITLE_SYNC_REFS                As String = "Synchronize VB-Project References" ' Identifies the synchronization dialog
Public Const TITLE_SYNC_SHEETS              As String = "Synchronize Worksheets"            ' Identifies the synchronization dialog
Public Const TITLE_SYNC_SHEET_SHAPES        As String = "Synchronize Sheet Shapes"          ' Identifies the synchronization dialog
Public Const TITLE_SYNC_VBCOMPS             As String = "Synchronize VBComponents"          ' Identifies the synchronization dialog
Public Const SYNC_ALL_BTTN                  As String = "Synchronize"                       ' Identifies the synchronization dialog

Private dctNew                              As Dictionary
Private dctObsolete                         As Dictionary
Private dctChanged                          As Dictionary

Public Enum SyncMode
    Count = 1
    Confirm = 2
    Synchronize = 3
End Enum

Public Sub AddApplRunArgs(ByVal ra_button As String, _
                          ByVal ra_service_name As String, _
                          ByVal ra_object As Variant, _
                          ByRef ra_cll As Collection)
' ------------------------------------------------------------------------------
' Returns the Collection (ra_cll) with the Application.Run arguments added
' ------------------------------------------------------------------------------
    Dim cll As New Collection
    
    Log.ServicedItem = ra_object
    cll.Add ra_button       ' 1. The button caption
    cll.Add ThisWorkbook    ' 2. The servicing Workbook
    cll.Add ra_service_name ' 3. The service to run
    cll.Add ra_object       ' 4. The VBComponent to add
    ra_cll.Add cll
    Set cll = Nothing
    
End Sub

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function SyncTargetsSource(ByVal sts_target_wbk As Workbook) As String
' ----------------------------------------------------------------------------
' Returns the full name of the Workbook which is considered the
' 'Sync-Target-Workbook's' corresponding source. A vbNullString is returned
' when it cannot be identified in CompMan's 'Serviced-Folder.
' ----------------------------------------------------------------------------
    Const PROC = "SyncTargetsSource"
    
    On Error GoTo eh
    Dim cll As Collection
    
    If mFile.Exists(ex_folder:=mConfig.FolderServiced, ex_file:=mSync.SyncTargetOriginName(sts_target_wbk), ex_result_files:=cll) Then
        If cll.Count <> 1 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "Unable to identify the 'Sync-Target-Workbook's corresponding 'Sync-Source-Workbook' because the " & _
                                                "'Serviced-Folder' contains either none or more than one file named '" & mSync.SyncTargetOriginName(sts_target_wbk) & "'" _
        Else SyncTargetsSource = cll(1).Path
    End If

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function SyncTargetOriginName(ByVal ston_wbk As Workbook) As String
' ----------------------------------------------------------------------------
' Returns serviced 'Sync-Target-Workbook's (ston_wbk) origin name.
' ----------------------------------------------------------------------------
    SyncTargetOriginName = Replace(ston_wbk.Name, SYNC_COPY_SUFFIX, vbNullString)
End Function

Public Function SyncTargetWorkingCopy(ByVal satc_wbk As Workbook) As Workbook
' ----------------------------------------------------------------------------
' When the Workbook (satc_wbk) is a Sync-Target-Origin-Workbook it is saved
' as a Sync-Target-Working-Copy and this Workbook is returned to the caller.
' When the Workbok (satc_wbk) already the Sync-Target-Working-Copy it is closed,
' deleted and the Sync-Target-Origin-Workbook is opened instead. All the
' remainingas is the same as above.
' ----------------------------------------------------------------------------
    
    Dim fso                     As New FileSystemObject
    Dim FileExtension           As String
    Dim sSyncTargetWorkingCopy  As String
    Dim sSyncTargetOrigin       As String
    Dim wbSyncTargetWrkngCpy    As Workbook
    Dim wbSyncTargetOrigin      As Workbook
    
    FileExtension = fso.GetExtensionName(satc_wbk.FullName)
    If InStr(satc_wbk.FullName, SYNC_COPY_SUFFIX) = 0 Then
        '~~ The opened Workbook is the origin Sync-Target-Workbook
        '~~ When the working copy is also already open it is closed without save
        sSyncTargetOrigin = satc_wbk.FullName
        sSyncTargetWorkingCopy = Replace(satc_wbk.FullName, "." & FileExtension, SYNC_COPY_SUFFIX & "." & FileExtension)
        If mWbk.IsOpen(sSyncTargetWorkingCopy, wbSyncTargetWrkngCpy) Then
            Application.EnableEvents = False
            wbSyncTargetWrkngCpy.Close False
            Application.EnableEvents = True
        End If
    
        '~~ Keep a record of the originally opend Sync-Target-Workbook
        wsService.SyncTargetOrigin = sSyncTargetOrigin
        
        '~~ Prepare for the working copy
        '~~  Delete any working copy still existing from a previous synchronization
        '~~ An already existing working copy is deleted
        If fso.FileExists(sSyncTargetWorkingCopy) Then
            fso.DeleteFile (sSyncTargetWorkingCopy)
        End If
        satc_wbk.SaveAs FileName:=sSyncTargetWorkingCopy, AccessMode:=xlExclusive  ' save Workbook under copy name
        
        '~~ Return the opened working copy
        Set SyncTargetWorkingCopy = mWbk.GetOpen(sSyncTargetWorkingCopy)                 ' return sync copy as new sync target
     
    ElseIf InStr(satc_wbk.FullName, SYNC_COPY_SUFFIX) <> 0 Then
        '~~ When the opened Workbook is the working copy of the Sync-Target-Workbook
        '~~ it is closed, deleted and the origin Sync-Target-Workbook is opened instead
        sSyncTargetWorkingCopy = Replace(satc_wbk.FullName, "." & FileExtension, SYNC_COPY_SUFFIX & "." & FileExtension)
        sSyncTargetOrigin = Replace(satc_wbk.FullName, SYNC_COPY_SUFFIX, vbNullString)
        sSyncTargetWorkingCopy = satc_wbk.FullName
        Application.EnableEvents = False
        satc_wbk.Close False
        mFile.Delete sSyncTargetWorkingCopy
        mWbk.GetOpen sSyncTargetOrigin
        Application.EnableEvents = True
    
         '~~ Keep a record of the originally opend Sync-Target-Workbook
        wsService.SyncTargetOrigin = sSyncTargetOrigin
        '~~ Prepare for the working copy
        '~~  Delete any working copy still existing from a previous synchronization
        '~~ An already existing working copy is deleted
        satc_wbk.SaveAs FileName:=sSyncTargetWorkingCopy, AccessMode:=xlExclusive  ' save Workbook under copy name
        
        '~~ Return the opened working copy
        Set SyncTargetWorkingCopy = mWbk.GetOpen(sSyncTargetWorkingCopy)                 ' return sync copy as new sync target
    End If
        
xt: Set fso = Nothing
    
End Function

Private Sub ClearSyncData()
    
    With wsService
        If CLng(.SyncDialogLeft) < 5 Then .SyncDialogLeft = 20
        If CLng(.SyncDialogTop) < 5 Then .SyncDialogTop = 20
    End With
    Application.ScreenUpdating = False
    wsSync.Clear
    
End Sub

Public Sub CollectSyncItems()
' ------------------------------------------------------------------------------
' Collects all synchronization issues/items - only for getting the max length
' of concerned sync items and item types and the number of involved items
' ------------------------------------------------------------------------------
    Const PROC = "CollectSyncItems"
    
    On Error GoTo eh
    Dim wbkTarget As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ClearSyncData
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    mSync.MonitorStep "Collecting References"
    mSyncRefs.CollectAllItems
    mSyncRefs.CollectNew
    mSyncRefs.CollectObsolete
    
    mSync.MonitorStep "Collecting Worksheets"
    mSyncSheets.CollectAllItems
    
    mSyncNames.CollectAllItems
    
'    mSync.MonitorStep "Collecting Named Ranges"
'    mSyncRanges.CollectAllItems
'
    mSync.MonitorStep "Collecting Sheet Shapes"
    mSyncShapes.CollectAllItems
    mSyncShapes.CollectNew
    
    mSync.MonitorStep "Collecting VB-Components"
    mSyncComps.CollectAllItems
    mSyncComps.CollectNew
    mSyncComps.CollectObsolete
    mSyncComps.CollectChanged
    

xt: Application.EnableEvents = True
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Public Function KeyControlName(ByVal s As String) As String
    KeyControlName = Split(s, SHEET_CONTROL_CONCAT)(1)
End Function

Public Function KeySheetControl(ByVal sheet_name As String, _
                                ByVal control_name As String) As String
    KeySheetControl = sheet_name & SHEET_CONTROL_CONCAT & control_name
End Function

Public Function KeySheetName(ByVal s As String) As String
    KeySheetName = Split(s, SHEET_CONTROL_CONCAT)(0)
End Function

Private Function NameExists(ByVal ne_wbk As Workbook, _
                            ByVal ne_nme As Name) As Boolean
    Dim nme As Name
    For Each nme In ne_wbk.Names
        NameExists = nme.Name = ne_nme.Name
        If NameExists Then Exit For
    Next nme
End Function


Private Sub RenameSheet(ByRef rs_wbk As Workbook, _
                        ByVal rs_old_name As String, _
                        ByVal rs_new_name As String)
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "RenameSheet"
    
    On Error GoTo eh
    Dim sh  As Worksheet
    For Each sh In rs_wbk.Worksheets
        If sh.Name = rs_old_name Then
            sh.Name = rs_new_name
            Log.Entry = "Sheet-Name changed to '" & rs_new_name & "'"
            Exit For
        End If
    Next sh

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub RenameWrkbkModule( _
                        ByRef rdm_wbk As Workbook, _
                        ByVal rdm_new_name As String)
' ---------------------------------------------------
' Renames in Workbook (rdm_wbk) the Workbook Module
' to (rdm_new_name).
' ---------------------------------------------------
    Const PROC = "RenameWrkbkModule"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    
    With rdm_wbk.SyncTargetithSource
        For Each vbc In .VBComponents
            If vbc.Type = vbext_ct_Document Then
                If mComp.IsWrkbkDocMod(vbc) Then
                    Log.ServicedItem = vbc
                    vbc.Name = rdm_new_name
                    Log.Entry = "Renamed to '" & rdm_new_name
                    Exit For
                End If
            End If
        Next vbc
    End With
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
    
Public Sub ClearSyncTargetExportFiles()
    Dim fso     As New FileSystemObject
    Dim sFolder As String
    Dim wbkTarget    As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    sFolder = wbkTarget.Path & "\" & mConfig.FolderExport
    With fso
        If .FolderExists(sFolder) Then .DeleteFolder sFolder
    End With
    Set fso = Nothing
End Sub

Public Sub UnloadSyncMessage(ByVal sm_title As String)
    With wsService
        .SyncDialogTop = mMsg.MsgInstance(sm_title).Top
        .SyncDialogLeft = mMsg.MsgInstance(sm_title).Left
    End With
    mMsg.MsgInstance sm_title, True

End Sub

Public Sub MonitorStep(ByVal ms_text As String)
    Dim s As String
    
    On Error Resume Next
    ActiveWindow.WindowState = xlMaximized
    s = "Synchronization (by " & ThisWorkbook.Name _
                          & ") for " _
                          & wsService.SyncSourceWorkbookName _
                          & ": " _
                          & ms_text
    mService.DsplyStatus s

End Sub

Public Sub RunSync()
' ------------------------------------------------------------------------------
' Performs all tasks to synchronizes all Types/Items in the the Sync-Target-
' Workbook with its corresponding (same named) item in the Sync-Source-Workbook.
'
' Attention! The order of synchronizations matters.
' ------------------------------------------------------------------------------
    Const PROC = "RunSync"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Select Case True
        Case Not wsSync.RefSyncDone
            mSync.MonitorStep "Synchronizing References"
            Set dctNew = mSyncRefs.CollectNew
            Set dctObsolete = mSyncRefs.CollectObsolete
            mSyncRefs.Sync dctNew, dctObsolete

        Case Not wsSync.WshSyncDone
            mSync.MonitorStep "Synchronizing Worksheets"
            Set dctNew = mSyncSheets.CollectNew
            Set dctObsolete = mSyncSheets.CollectObsolete
            Set dctChanged = mSyncSheets.CollectChanged
            mSyncSheets.Sync dctNew, dctObsolete, dctChanged

        Case Not wsSync.NmeSyncDone
            mSync.MonitorStep "Synchronizing Range Names"
            Set dctNew = mSyncNames.CollectNew
            Set dctObsolete = mSyncNames.CollectObsolete
            Set dctChanged = mSyncNames.CollectChanged
            mSyncNames.Sync dctNew, dctObsolete, dctChanged

'        Case Not wsSync.RngSyncDone
'            mSync.MonitorStep "Synchronizing Named Ranges"
'            Set dctChanged = mSyncRanges.CollectChanged
'            mSyncRanges.Sync dctChanged
        
        Case Not wsSync.ShpSyncDone
            mSync.MonitorStep "Synchronizing Sheet Shapes"
            Set dctNew = mSyncShapes.CollectNew
            Set dctObsolete = mSyncShapes.CollectObsolete
            UnloadSyncMessage TITLE_SYNC_SHEET_SHAPES
            mSyncShapes.Sync dctNew, dctObsolete

        Case Not mSyncComps.Done(dctNew, dctObsolete, dctChanged)
            mSync.MonitorStep "Synchronizing VB-Components"
            mSyncComps.Sync dctNew, dctObsolete, dctChanged
    
        Case mSyncComps.Done(dctNew, dctObsolete, dctChanged)
            mSync.MonitorStep "Done!"
            wsSync.VbcSyncDone = True
            mSync.ClearSyncTargetExportFiles
    End Select
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function WbkGetOpen(ByVal go_wbk_full_name As String) As Workbook
' ----------------------------------------------------------------------
' Returns an opened Workbook object named (go_wbk_full_name) or Nothing
' when a file named (go_wbk_full_name) not exists.
' ----------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If fso.FileExists(go_wbk_full_name) Then
        If WbkIsOpen(wbk_full_name:=go_wbk_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wbk_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wbk_full_name)
    End If
    
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsOpen( _
            Optional ByVal wbk_name As String = vbNullString, _
            Optional ByVal wbk_full_name As String = vbNullString) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE when a Workbook either identified by its BaseName (wbk_base_name)
' or by its full name (wbk_full_name) is open. When the BaseName is
' provided in the current Excel instance, else in any Excel instance.
' -------------------------------------------------------------------------
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim xlApp As Excel.Application
    
    If wbk_name = vbNullString And wbk_full_name = vbNullString Then GoTo xt
    
    With New FileSystemObject
        If wbk_full_name <> vbNullString Then
            '~~ With the full name the open test spans all application instances
            If Not .FileExists(wbk_full_name) Then GoTo xt
            If wbk_name = vbNullString Then wbk_name = .GetFileName(wbk_full_name)
            On Error Resume Next
            Set xlApp = VBA.GetObject(wbk_full_name).Application
            WbkIsOpen = Err.Number = 0
        Else
            On Error Resume Next
            wbk_name = Application.Workbooks(wbk_name).Name
            WbkIsOpen = Err.Number = 0
        End If
    End With

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

