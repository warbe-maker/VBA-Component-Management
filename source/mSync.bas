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

Public Const APP_RUN_ARG_BUTTON_CAPTION     As Long = 1
Public Const APP_RUN_ARG_SERVICE            As Long = 3
Public Const APP_RUN_ARG_SERVICE_ARG1       As Long = 4
Public Const APP_RUN_ARG_SERVICE_ARG2       As Long = 5
Public Const APP_RUN_ARG_SERVICE_ARG3       As Long = 6
Public Const APP_RUN_ARG_SERVICING_WORKBOOK As Long = 2
Public Const README_SYNC_CHAPTER            As String = "?#using-the-synchronization-service"
Public Const README_SYNC_CHAPTER_NEW_NAMES  As String = "?#synchronization-of-new-names"
Public Const README_SYNC_CHAPTER_SHEETS     As String = "?#worksheet-synchronization"
Public Const SYNC_ALL_BTTN                  As String = "Synchronize"                       ' Identifies the synchronization dialog
Public Const SYNC_TARGET_SUFFIX             As String = "_Target"                           ' suffix for the sync target working copy
Public Const TITLE_SYNC_NAMES               As String = "Synchronize Range Names"           ' Identifies the synchronization dialog
Public Const TITLE_SYNC_RANGES              As String = "Synchronize Named Ranges"          ' Identifies the synchronization dialog
Public Const TITLE_SYNC_REFS                As String = "Synchronize VB-Project References" ' Identifies the synchronization dialog
Public Const TITLE_SYNC_SHEETS              As String = "Synchronize Worksheets"            ' Identifies the synchronization dialog
Public Const TITLE_SYNC_SHEET_SHAPES        As String = "Synchronize Sheet Shapes"          ' Identifies the synchronization dialog
Public Const TITLE_SYNC_VBCOMPS             As String = "Synchronize VBComponents"          ' Identifies the synchronization dialog

' Synchronization open option buttons ---------------
Private AbortOpenDialogForPreparationCopy   As String
Private AbortOpenDlogForPreparationTarget   As String
Private ContinueSyncWithExstngWorkingCopy   As String
Private ContinueSyncWithOpenedWorkingCopy   As String
Private ContinueOpenWithTargetSynchronztn   As String
Private ContinueWithSynchAgainFromScratch   As String
' ---------------------------------------------------

Private dctNew                              As Dictionary
Private dctObsolete                         As Dictionary
Public dctObsoleteNames                     As Dictionary
Private dctChanged                          As Dictionary
Private dctOwnedByProject                   As Dictionary
Private wbkSyncSource                       As Workbook     ' the synchronization services global opened Sync-Source-Workbook (same name as the opened Sync-Target-Workbook)
Private wbkSyncTarget                       As Workbook     '  the synchronization services global opened Sync-Target-Workbook
Private wbkSyncTargetCopy                   As Workbook     '  the synchronization services global opened Sync-Target-Workbook's working copy
Private OpenDecisionMsgTitle                As String

Public Enum SyncMode
    Count = 1
    Confirm = 2
    Synchronize = 3
End Enum

Public Property Get source() As Workbook
    Dim s As String
    
    On Error Resume Next
    s = wbkSyncSource.name
    If Err.Number <> 0 Then
        Set wbkSyncSource = mWbk.GetOpen(wsService.SyncSourceFullName)
    End If
    Set source = wbkSyncSource

End Property

Public Property Let source(ByVal s_wbk As Workbook):    Set wbkSyncSource = s_wbk:  End Property

Public Property Get Target() As Workbook
    Dim s As String
    
    On Error Resume Next
    s = wbkSyncTarget.name
    If Err.Number <> 0 Then
        Set wbkSyncTarget = mWbk.GetOpen(wsService.ServicedWorkbookFullName)
    End If
    Set Target = wbkSyncTarget
    mService.WbkServiced = wbkSyncTarget
    
End Property

Public Property Let Target(ByVal t_wbk As Workbook):        Set wbkSyncTarget = t_wbk:      End Property

Public Property Get TargetCopy() As Workbook

    Dim s As String
    
    On Error Resume Next
    s = wbkSyncTarget.name
    If Err.Number <> 0 Then
        Set wbkSyncTargetCopy = mWbk.GetOpen(wsService.SyncTargetFullNameCopy)
    End If
    Set TargetCopy = wbkSyncTargetCopy
    mService.WbkServiced = wbkSyncTargetCopy
    
End Property

Public Property Let TargetCopy(ByVal l_wbk As Workbook):    Set wbkSyncTargetCopy = l_wbk:  End Property

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

Public Sub ClearSyncData()
    
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
    
    mBasic.BoP ErrSrc(PROC)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    
    mSync.MonitorStep "Collecting Sync References"
    mSyncRefs.CollectAllItems
    mSyncRefs.CollectNew
    mSyncRefs.CollectObsolete
    
    mSync.MonitorStep "Collecting Sync Worksheets"
    mSyncSheets.CollectAllItems
    
    mSync.MonitorStep "Collecting Sync Names"
    mSyncNames.CollectAllItems
    
    mSync.MonitorStep "Collecting Sync Sheet Shapes"
    mSyncShapes.CollectAllItems
    mSyncShapes.CollectNew
    
    mSync.MonitorStep "Collecting Sync VB-Components"
    mSyncComps.CollectAllItems
    mSyncComps.CollectNew
    mSyncComps.CollectObsolete
    mSyncComps.CollectChanged
    

xt: mBasic.EoP ErrSrc(PROC)
    Application.EnableEvents = True
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Public Sub Finalize()
' ------------------------------------------------------------------------------
' Displays a finalization dialog including a synchronization summary. When the
' finalization is confirmed the Sync-Target-Workbook's working copy is saved
' under its origin name and the working copy is deleted.
' ------------------------------------------------------------------------------
    Const Finalize              As String = "Finalize this Synchronization"
    Const ABORT_FINALIZATION    As String = "Abort the finalization"
    
    Dim MsgText     As TypeMsg
    Dim MsgButtons  As Collection
    
    Set MsgButtons = mMsg.Buttons(Finalize, vbLf, ABORT_FINALIZATION)
    With MsgText
        With .Section(1)
            .Label.Text = Finalize & ":"
            .Label.FontColor = rgbDarkGreen
            .Text.Text = "The Sync-Target-Workbook's working copy will be saved under its origin name and " & _
                         "deleted."
        End With
        With .Section(2)
            .Label.Text = "Please note:"
            .Label.FontColor = rgbDarkGreen
            .Text.Text = "A re-synchronization will still be possible when the most recent archive is " & _
                         "copied under its origin name and re-opened. However, any changes made in the production " & _
                         "version in the meantime will get lost!"
        End With
        With .Section(5)
            .Label.Text = "See: Using the Synchronization Service"
            .Label.FontColor = rgbBlue
            .Label.OpenWhenClicked = mCompMan.README_URL & mSync.README_SYNC_CHAPTER
            .Text.Text = "The chapter 'Using the Synchronization Service' will provide additional information"
        End With
    End With

    Select Case mMsg.Dsply(dsply_title:="Finalization of the synchronization" _
                         , dsply_msg:=MsgText _
                         , dsply_buttons:=MsgButtons)
        Case Finalize
            mSync.TargetCopy.SaveAs mSync.TargetOriginFullName(mSync.TargetCopy)
'            mSync.TargetCopyDelete
        Case ABORT_FINALIZATION
    End Select
                             
End Sub

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

Public Sub MonitorStep(ByVal ms_text As String)
    Dim s As String
    
    On Error Resume Next
    ActiveWindow.WindowState = xlMaximized
    s = "Synchronization (by " & ThisWorkbook.name _
                          & ") for " _
                          & mService.WbkServiced.name _
                          & ": " _
                          & ms_text
    Application.StatusBar = vbNullString
    mService.DsplyStatus s

End Sub

Private Function NameExists(ByVal ne_wbk As Workbook, _
                            ByVal ne_nme As name) As Boolean
    Dim nme As name
    For Each nme In ne_wbk.Names
        NameExists = nme.name = ne_nme.name
        If NameExists Then Exit For
    Next nme
End Function

Public Sub OpenDecision()
' ------------------------------------------------------------------------------
' Displays a mode-less dialog with open-decision-buttons.
' ------------------------------------------------------------------------------
    Const PROC  As String = "OpenDecision"
    
    On Error GoTo eh
    Dim MsgText     As TypeMsg
    Dim MsgButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    
    '~~ Prepare decision option button captions
    ContinueOpenWithTargetSynchronztn = "Synchronize"
    ContinueWithSynchAgainFromScratch = "Re-Synchronize"
    ContinueSyncWithExstngWorkingCopy = "Continue Synchronization with" & vbLf & _
                                        "existing Sync-Target-Workbook's working copy"
    ContinueSyncWithOpenedWorkingCopy = "Continue ongoing Synchronization with" & vbLf & _
                                        "opened Sync-Target-Workbook's working copy"
    AbortOpenDlogForPreparationTarget = "Stop for pre-synchronization preparations"
    
    If TargetCopyIsOpened _
    Then OpenDecisionMsgTitle = "Decide about the opened Sync-Target-Workbook's working copy" _
    Else OpenDecisionMsgTitle = "Decide about the opened Sync-Target-Workbook"

    With MsgText
        If OpenedIsTarget Then ' The Sync-Target-Wokbook is opened
            With .Section(1)
                If TargetCopyExists Then
                    '~~ Re-Synchronize (again from scratch)
                    .Label.Text = Replace(ContinueWithSynchAgainFromScratch, vbLf, " ") & ":"
                    .Label.FontColor = rgbDarkGreen
                    .Text.Text = "Although there's an ongoing, still un-finalized synchronization going on " & _
                                 "(indicated by an existing Sync-Target-Workbook's working copy) the " & _
                                 "synchronization will restarted from scratch. I.e. the Sync-Target-Workbook " & _
                                 "will again be archived and the already existing Sync-Target-Workbook's " & _
                                 "working copy will be ignored."
                    Set MsgButtons = mMsg.Buttons(ContinueWithSynchAgainFromScratch)
                    mMsg.ButtonAppRun AppRunArgs _
                                    , ContinueWithSynchAgainFromScratch _
                                    , ThisWorkbook _
                                    , "mSync.OpenDecisionReSyncFromScratch"
                Else
                    '~~ Synchronize
                    .Label.Text = Replace(ContinueOpenWithTargetSynchronztn, vbLf, " ") & ":"
                    .Label.FontColor = rgbDarkGreen
                    .Text.Text = "The opened Sync-Target-Workbook will be archived and its VB-Project will " & _
                                 "be synchronized with its corresponding Sync-Source-Workbook. For this - " & _
                                 "because Source- and Target-Workbook have the same name - the Sync-Target-Workbook " & _
                                 "is saved as Sync-Target-Workbook working copy and this one will be synchronized."
                    Set MsgButtons = mMsg.Buttons(ContinueOpenWithTargetSynchronztn)
                    mMsg.ButtonAppRun AppRunArgs _
                                    , ContinueOpenWithTargetSynchronztn _
                                    , ThisWorkbook _
                                    , "mSync.OpenDecisionTargetSync"
                End If
            End With
            If TargetCopyExists Then
                With .Section(2)
                        '~~ Continue with ongoing synchronization
                        .Label.Text = Replace(ContinueSyncWithExstngWorkingCopy, vbLf, " ") & ":"
                        .Label.FontColor = rgbDarkGreen
                        .Text.Text = "An already started but yet unfinished and/or not finalized synchronization " & _
                                     "will be continued. When all synchronizations had been done a dialog " & _
                                     "for its finalization will be displayed." & vbLf & _
                                     "A continuation may have become  appropriate when a synchronization hads been " & _
                                     "terminated for manual pre-synchronization work in the Sync-Target-Workbook's " & _
                                     "working copy. In that case it should be noted, that such modifications will get " & _
                                     "lost with a re-synchronization from scratch."
                End With
                Set MsgButtons = mMsg.Buttons(MsgButtons, vbLf, ContinueSyncWithExstngWorkingCopy)
                mMsg.ButtonAppRun AppRunArgs _
                                , ContinueSyncWithExstngWorkingCopy _
                                , ThisWorkbook _
                                , "mSync.OpenDecisionContinueWithExstngWorkingCopy"
            End If
            With .Section(3)
                .Label.Text = Replace(AbortOpenDlogForPreparationTarget, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The Sync-Target-Workbook is kept open for possibly still required " & _
                             "pre-synchronization preparations. When the Sync-Target-Workbook is " & _
                             "closed thereafter and re-opened any other decision may be taken."
            End With
            Set MsgButtons = mMsg.Buttons(MsgButtons, vbLf, AbortOpenDlogForPreparationTarget)
            mMsg.ButtonAppRun AppRunArgs _
                            , AbortOpenDlogForPreparationTarget _
                            , ThisWorkbook _
                            , "mSync.OpenDecisionPreparationTarget"
        
        ElseIf TargetCopyIsOpened Then
            '~~ When the Sync-Target-Workbook's working copy is opened this likely means that
            '~~ the synchronization had already be started and may have passed som steps already.
            '~~ The option to prepare the working copy is not recommendable and thus is not offered.
            With .Section(1)
                '~~ Continue with opened working copy
                .Label.Text = Replace(ContinueSyncWithOpenedWorkingCopy, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The opened Sync-Target-Workbook's working copy is used to continue a yet " & _
                             "unfinished (or just pending finalization) synchronization. A finalization " & _
                             "dialog will be displayed once all synchronizations are done."

            End With
            Set MsgButtons = mMsg.Buttons(ContinueSyncWithOpenedWorkingCopy)
            mMsg.ButtonAppRun AppRunArgs _
                            , ContinueSyncWithOpenedWorkingCopy _
                            , ThisWorkbook _
                            , "mSync.OpenDecisionContinueSyncWithOpenedWorkingCopy"
            
            With .Section(2)
                '~~ Synchronize again
                .Label.Text = Replace(ContinueWithSynchAgainFromScratch, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The synchronization will be performed again from scratch. The " & _
                             "Sync-Target-Workbook will again be archived and an already " & _
                             "existing Sync-Target-Workbook's working copy will be ignored. " & vbLf & _
                             "Re-synchronization will be appropriate when the Sync-Target-Workbook " & _
                             "had been modified (manual pre-synchronization work)." & vbLf & _
                             "Attention: Any manual pre-synchronization preparations mad in the " & _
                             "Sync-Target-Workbook's working copy will get lost!"
            End With
            Set MsgButtons = mMsg.Buttons(MsgButtons, vbLf, ContinueWithSynchAgainFromScratch)
            mMsg.ButtonAppRun AppRunArgs _
                            , ContinueWithSynchAgainFromScratch _
                            , ThisWorkbook _
                            , "mSync.OpenDecisionReSyncFromScratch"
            
            With .Section(3)
                '~~ Prepare Sync-Target-Workbook
                .Label.Text = Replace(AbortOpenDialogForPreparationCopy, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The Sync-Target-Workbook's working copy will closed without save and removed and " & _
                             "the Sync-Target-Workbook will be opened for a manual pre-sync-preparation instead. " & _
                             "When done the Sync-Target-Workbook will be closed and re-opened for being synchronized."
            End With
            
            Set MsgButtons = mMsg.Buttons(MsgButtons, vbLf, AbortOpenDlogForPreparationTarget)
            mMsg.ButtonAppRun AppRunArgs _
                            , AbortOpenDlogForPreparationTarget _
                            , ThisWorkbook _
                            , "mSync.OpenDecisionPreparationTarget"
            
        End If ' Target copy is opened
    
        With .Section(5)
            With .Label
                .Text = "See: Using the Synchronization Service"
                .FontColor = rgbBlue
                .OpenWhenClicked = mCompMan.README_URL & mSync.README_SYNC_CHAPTER
            End With
            .Text.Text = "The chapter 'Using the Synchronization Service' will provide additional information"
        End With
    End With
    
    '~~ Display the mode-less open decision dialog
    mMsg.Dsply dsply_title:=OpenDecisionMsgTitle _
             , dsply_msg:=MsgText _
             , dsply_buttons:=MsgButtons _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=AppRunArgs _
             , dsply_width_min:=45 _
             , dsply_pos:=wsService.SyncDialogTop & ";" & wsService.SyncDialogLeft
                            
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub OpenDecisionContinueSyncWithOpenedWorkingCopy()
' ------------------------------------------------------------------------------
' A Sync-Target-Workbook's working copy has directly been opened
' ------------------------------------------------------------------------------
    Const PROC = "OpenDecisionContinueSyncWithOpenedWorkingCopy"
    
    On Error GoTo xt
    Dim wbk             As Workbook
    Dim wbkTargetCopy   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpenDecisionMsgTitle, True
    mSync.Target = mWbk.GetOpen(wsService.ServicedWorkbookFullName)
    mSync.TargetCopyOpen False, wbkTargetCopy
    mSync.TargetCopy = wbkTargetCopy
    If mWbk.IsOpen(wsService.ServicedWorkbookFullName, wbk) Then
        mSync.TargetClose ' Close any still open Sync-Target-Workbook
    End If
    mSync.SourceOpen  ' Open the corresponding Sync-Source-Workbook
    mSync.RunSync

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub OpenDecisionContinueWithExstngWorkingCopy()
' ------------------------------------------------------------------------------
' A Sync-Target-Workbook has been opened with an already existing working copy
' and Continue had been decided with the open dialog.
' ------------------------------------------------------------------------------
    Const PROC = "OpenDecisionContinueWithExstngWorkingCopy"
    
    On Error GoTo xt
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpenDecisionMsgTitle, True ' Close/unload the mode-less open dialog
    mSync.TargetClose                           ' Close the opened Sync-Target-Workbook
    mSync.TargetCopyOpen False                  ' Opem the existing Sync-Target-Workbook's working copy
    mSync.SourceOpen                            ' Open the corresponding Sync-Source-Workbook
    mSync.TargetCopy.Activate
    mSync.RunSync

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub OpenDecisionTargetSync()
' ------------------------------------------------------------------------------
' Syncronize the opened Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "OpenDecisionTargetSync"
    
    On Error GoTo xt
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpenDecisionMsgTitle, True
    mSync.TargetArchive                         ' Archive the Sync-Target-Workbook
    mSync.TargetCopyOpen                        ' Establish a new Sync-Target-Workbook working copy
    mSync.SourceOpen                            ' Open the corresponding Sync-Source-Workbook
    mSync.ClearSyncData                         ' Clear any data from a previous synchronization
    mSync.TargetClearExportFiles
    mSync.CollectSyncItems
    mSync.RunSync

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub OpenDecisionPreparationTarget()
' ------------------------------------------------------------------------------
' When the Sync-Target-Workbook had been opened it simply will remain open.
' When the Sync-Target-Workbook's working copy had been opend:
' - it will be closed (without save!) and removed (a prepared Sync-Target-
'   Workbook will require a sync from scratch.
' - The Sync-Source-Workbook will be closed and the Sync-Target-Workbook will be
'   opened instead.
' ------------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    
    mMsg.MsgInstance OpenDecisionMsgTitle, True ' Close/unload the mode-less open dialog
    If OpenedIsTarget Then
        With fso
            If .FileExists(wsService.SyncTargetFullNameCopy) _
            Then .DeleteFile wsService.SyncTargetFullNameCopy
        End With
    ElseIf OpenedIsTargetCopy Then
        ActiveWorkbook.Close False
        mSync.SourceClose
        With fso
            If .FileExists(wsService.SyncTargetFullNameCopy) _
            Then .DeleteFile wsService.SyncTargetFullNameCopy
        End With
        Application.EnableEvents = False
        mSync.TargetOpen
        Application.EnableEvents = True
        
    End If

End Sub

Public Sub OpenDecisionReSyncFromScratch()
' ------------------------------------------------------------------------------
' Resynchronizer the opened Sync-Target-Workbook although a working copy already
' exists.
' ------------------------------------------------------------------------------
    Const PROC = "OpenDecisionTargetSync"
    
    On Error GoTo xt
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpenDecisionMsgTitle, True ' Close/teminate the mode-less open dialog
    mSync.SourceClose                           ' Close a possibly already open Sync-Source-Workbook
    mSync.TargetArchive                         ' Archive the Sync-Target-Workbook
    mSync.TargetCopyOpen                        ' Establish a new working copy
    mSync.SourceOpen                            ' Open the corresponding Sync-Source-Workbook
    mSync.ClearSyncData                         ' Clear any data from a previous synchronization
    mSync.TargetClearExportFiles
    mSync.CollectSyncItems
    mSync.RunSync

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
        If sh.name = rs_old_name Then
            sh.name = rs_new_name
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
                    vbc.name = rdm_new_name
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

Public Sub RunSync()
' ------------------------------------------------------------------------------
' Performs all tasks to synchronizes all Types/Items in the the Sync-Target-
' Workbook with the Sync-Source-Workbook.
'
' Attention! Synchronizing Workshhets is a prerequisite for the synchronization
'            of Range Names, and Worksheet Shapes.
' ------------------------------------------------------------------------------
    Const PROC = "RunSync"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mService.WbkServiced = mSync.TargetCopy
    
    Select Case True
        Case Not wsSync.WshSyncDone
            mSync.MonitorStep "Synchronizing Worksheets"
            Set dctNew = mSyncSheets.CollectNew
            Set dctObsolete = mSyncSheets.CollectObsolete
            Set dctChanged = mSyncSheets.CollectChanged
            Set dctOwnedByProject = mSyncSheets.CollectOwnedByProject
            mSyncSheets.Sync dctNew, dctObsolete, dctChanged, dctOwnedByProject

        Case Not wsSync.NmeSyncDone
            mSync.MonitorStep "Synchronizing Range Names"
            Set dctNew = mSyncNames.CollectNew
            Set dctObsolete = mSyncNames.CollectObsolete
            mSyncNames.Sync dctNew, dctObsolete

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
            mSync.TargetClearExportFiles
        
        Case Not wsSync.RefSyncDone
            mSync.MonitorStep "Synchronizing References"
            Set dctNew = mSyncRefs.CollectNew
            Set dctObsolete = mSyncRefs.CollectObsolete
            mSyncRefs.Sync dctNew, dctObsolete

    End Select
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SourceClose()
' ------------------------------------------------------------------------------
' Closes an open Sync-Source-Workbook by considering still unsaved change.
' ------------------------------------------------------------------------------
    Dim wbk As Workbook
    Dim sSourceFullName As String
    
    sSourceFullName = wsService.SyncSourceFullName
    If mWbk.IsOpen(sSourceFullName, wbk) Then
        If wbk.FullName = sSourceFullName Then
            Application.EnableEvents = False
            If wbk.Saved _
            Then wbk.Close False _
            Else wbk.Close True
            Application.EnableEvents = True
        End If
    End If

End Sub

Public Function SourceExists(ByVal se_wbk_opened As Workbook) As Boolean
' ------------------------------------------------------------------------------
' When the Sync-Source-Workbook unambigously exists in the Serviced-DevAndTest-
' Folder the function returns TRUE and saves the found Workbook's full name to
' wsService.SyncSourceFullName, else the function displays a corresponding
' message and returns FALSE.
' ------------------------------------------------------------------------------
    Dim cll             As Collection
    Dim sSourceFullName As String
    Dim sResult         As String
    Dim Msg             As TypeMsg
    Dim MsgTitle        As String
    Dim i               As Long
    Dim MsgButtons      As Collection
    
    mFso.Exists ex_folder:=wsConfig.FolderDevAndTest _
              , ex_file:=Replace(se_wbk_opened.name, SYNC_TARGET_SUFFIX & ".xls", ".xls") _
              , ex_result_files:=cll
    Select Case cll.Count
        Case 0
            MsgTitle = "No corresponding Sync-Source-Workbook found!"
            With Msg.Section(1)
                .Text.Text = "No correponding Sync-Source-Workbook for the opened Sync-Target-Workbook " & _
                             "(" & se_wbk_opened.name & ") could be found in the configured 'ServicedDevAndTest' folder " & _
                             "(" & wsConfig.FolderDevAndTest & ")."
            End With
        Case 1
            sSourceFullName = cll(1)
        Case Else
            MsgTitle = "Ambigous Sync-Source-Workbooks found!"
            With Msg
                .Section(1).Text.Text = "For the opened Sync-Target-Workbook (" & se_wbk_opened.name & ") ambigous " & _
                                        "corresponding Sync-Source-Workbooks had been found in the configured " & _
                                        "'ServicedDevAndTest' folder (" & wsConfig.FolderDevAndTest & "):"
                With .Section(2).Text
                    .MonoSpaced = True
                    .FontSize = 8
                    .Text = cll(1)
                    For i = 2 To cll.Count
                        .Text = .Text & vbLf & cll(i)
                    Next i
                End With
                .Section(3).Text.Text = "Terminate this synchronization trial, first remove the additional Workbook's or move them outside the " & _
                             wsConfig.FolderDevAndTest & " folder and re-open the Sync-Target-Workbook."
            End With
    End Select
            
    If sSourceFullName = vbNullString Then
        With Msg.Section(5)
            With .Label
                .Text = "See: Using the Synchronization Service"
                .FontColor = rgbBlue
                .OpenWhenClicked = mCompMan.README_URL & mSync.README_SYNC_CHAPTER
            End With
            .Text.Text = "The chapter 'Using the Synchronization Service' will provide additional information"
        End With
        mMsg.Dsply dsply_title:=MsgTitle _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=mMsg.Buttons("Terminate Synchronization") _
                 , dsply_width_min:=30
        wsService.SyncSourceFullName = vbNullString
    Else
        SourceExists = True
        wsService.SyncSourceFullName = sSourceFullName
    End If
    
End Function

Public Sub SourceOpen()
    
    Application.EnableEvents = False
    Set wbkSyncSource = mWbk.GetOpen(wsService.SyncSourceFullName)
    Application.EnableEvents = True

End Sub

Public Sub TargetArchive()
' ------------------------------------------------------------------------------
' Archives the opened Sync-Target-Workbook under the archiving name
' <workbook-base-name>-yy-mm-dd-nn.<extension> in a dedicated folder named
' <workbook-base-name>-yy-mm-dd-nn whereby nn is a number from 01 to 99 to
' distinguish a Sync-Target-Workbook archived more than once during a day.
' ------------------------------------------------------------------------------
    Dim fso                     As New FileSystemObject
    Dim sArchiveFolderRoot      As String
    Dim sArchiveFolderTarget    As String
    Dim l                       As Long
    Dim sArchivedWbkFullName    As String
    Dim sTargetWbkBaseName      As String
    Dim sTargetFullName         As String
    Dim sArchiveSuffix          As String
    Dim sExt                    As String
    
    sTargetFullName = wsService.ServicedWorkbookFullName
    With fso
        sExt = .GetExtensionName(sTargetFullName)
        '~~ Establish the archive root folder
        sArchiveFolderRoot = .GetParentFolderName(wsConfig.FolderSyncTarget) & "\SyncArchive"
        If Not .FolderExists(sArchiveFolderRoot) Then .CreateFolder sArchiveFolderRoot
    
        sTargetWbkBaseName = .GetBaseName(sTargetFullName)
        For l = 1 To 99
            sArchiveSuffix = "-" & Format(Now(), "yy-mm-dd-") & Format(l, "00")
            sArchiveFolderTarget = sArchiveFolderRoot & "\" & sTargetWbkBaseName & sArchiveSuffix
            If Not .FolderExists(sArchiveFolderTarget) Then
                .CreateFolder sArchiveFolderTarget
                sArchivedWbkFullName = sArchiveFolderTarget & "\" & sTargetWbkBaseName & sArchiveSuffix & "." & sExt
                mWbk.GetOpen(sTargetFullName).SaveCopyAs sArchivedWbkFullName
                Exit For
            End If
        Next l
    End With
    
End Sub

Public Sub TargetClearExportFiles()
    Dim fso             As New FileSystemObject
    Dim sFolder         As String
    Dim sTargetFullName As String
    
    sTargetFullName = wsService.ServicedWorkbookFullName
    With fso
        sFolder = .GetParentFolderName(sTargetFullName) & "\" & wsConfig.FolderExport
        If .FolderExists(sFolder) Then .DeleteFolder sFolder
    End With
    Set fso = Nothing
    
End Sub

Public Sub TargetOpen()
    Set wbkSyncTarget = mWbk.GetOpen(wsService.ServicedWorkbookFullName)
End Sub

Public Sub TargetClose()
' ------------------------------------------------------------------------------
' Closes an open Sync-Target-Workbook by considering possible changes made.
' ------------------------------------------------------------------------------
    Dim wbk                         As Workbook
    Dim Msg                         As TypeMsg
    Dim MsgTitle                    As String
    Dim BttnCloseWithoutSaving      As String
    Dim BttnCloseBySavingChanges    As String
    
    If mWbk.IsOpen(wsService.ServicedWorkbookFullName, wbk) Then
        BttnCloseWithoutSaving = "Close the open Sync-Target-Workbook" & vbLf & _
                                 "without saving the changes"
        BttnCloseBySavingChanges = "Close the open Sync-Target-Workbook" & vbLf & _
                                   "by saving the changes"
        
        If Not wbk.Saved Then
            MsgTitle = "Sync-Target-Workbook still open with unsaved changes!"
            With Msg.Section(1)
                With .Label
                    .Text = Replace(BttnCloseWithoutSaving, vbLf, " ")
                    .FontColor = rgbGreen
                End With
                With .Text
                    .Text = "In case the changes were manual pre-synchronisation modifications such " & _
                            "as insertions or removals of cells, row, or coloumns, these modifications" & vbLf & _
                            "w i l l   g e t   l o s t !"
                End With
            End With
            With Msg.Section(2)
                With .Label
                    .Text = Replace(BttnCloseBySavingChanges, vbLf, " ")
                    .FontColor = rgbGreen
                End With
                With .Text
                    .Text = "Attention!" & vbLf & _
                            "The changes made will become effective only when the Sync-Target-Workbook " & _
                            "is re-opened and a Re-Synchronization (from scratch) is chosen in the " & _
                            "displayed open decision dialog."
                End With
            End With
            Select Case mMsg.Dsply(dsply_title:=MsgTitle _
                                , dsply_msg:=Msg _
                                , dsply_buttons:=mMsg.Buttons(BttnCloseWithoutSaving, vbLf, BttnCloseBySavingChanges))
                Case BttnCloseWithoutSaving:    wbk.Close False
                Case BttnCloseBySavingChanges:  wbk.Close True
            End Select
        Else
            Application.EnableEvents = False
            wbk.Close
            Application.EnableEvents = True
        End If
    End If
End Sub

Public Sub TargetCopyClose()
' ------------------------------------------------------------------------------
' Closes a still open Sync-Target-Workbook's working copy by considering
' possible synchronizations or manual changes made.
' ------------------------------------------------------------------------------
    Dim wbk                         As Workbook
    Dim Msg                         As TypeMsg
    Dim MsgTitle                    As String
    Dim BttnCloseWithoutSaving      As String
    Dim BttnCloseBySavingChanges    As String
    Dim sTargetCopyFullName         As String
    
    sTargetCopyFullName = wsService.SyncTargetFullNameCopy
    If mWbk.IsOpen(sTargetCopyFullName, wbk) Then
        BttnCloseWithoutSaving = "Close the open Sync-Target-Workbook's working copy" & vbLf & _
                                 "without saving the changes"
        BttnCloseBySavingChanges = "Close the open Sync-Target-Workbook's working copy" & vbLf & _
                                   "by saving the changes"
        
        If Not wbk.Saved Then
            MsgTitle = "The Sync-Target-Workbook's working copy is still open with unsaved changes!"
            With Msg.Section(1)
                With .Label
                    .Text = Replace(BttnCloseWithoutSaving, vbLf, " ")
                    .FontColor = rgbGreen
                End With
                With .Text
                    .Text = "In case the changes result from an aborted synchronization, saving without changes " & _
                            "will not do any harm. With the next open all already made synchronizations will just " & _
                            "be made again. In case the changes were manual pre-synchronisation modifications " & _
                            "such as insertions or removals of cells, row, or coloumns, these modification " & vbLf & _
                            "w i l l  g e t   l o s t !"
                End With
            End With
            With Msg.Section(2)
                With .Label
                    .Text = Replace(BttnCloseBySavingChanges, vbLf, " ")
                    .FontColor = rgbGreen
                End With
                With .Text
                    .Text = "Attention!" & vbLf & _
                            "Any changes made by an aborted synchronization or manual pre-synchronisation modifications " & _
                            "will become effective when the Sync-Target-Workbook (or the Sync-Target-Workbook's working " & _
                            "copy directly) is re-opened and 'Continue ongoing Synchronization' is chosen in the displayed open decision " & _
                            "dialog."
                End With
            End With
            Select Case mMsg.Dsply(dsply_title:=MsgTitle _
                                , dsply_msg:=Msg _
                                , dsply_buttons:=mMsg.Buttons(BttnCloseWithoutSaving, vbLf, BttnCloseBySavingChanges))
                Case BttnCloseWithoutSaving:    wbk.Close False
                Case BttnCloseBySavingChanges:  wbk.Close True
            End Select
        Else
            Application.EnableEvents = False
            wbk.Close
            Application.EnableEvents = True
        End If
    End If
End Sub

Public Function TargetCopyDelete() As Boolean
    Dim fso                 As New FileSystemObject
    Dim sTargetCopyFullName As String
    
    sTargetCopyFullName = wsService.SyncTargetFullNameCopy
    With fso
        If .FileExists(sTargetCopyFullName) Then .DeleteFile sTargetCopyFullName
    End With
    Set fso = Nothing

End Function

Private Function TargetCopyExists() As Boolean
    Dim fso As New FileSystemObject
    Dim sTargetCopyFullName As String
    
    sTargetCopyFullName = wsService.SyncTargetFullNameCopy
    TargetCopyExists = fso.FileExists(sTargetCopyFullName)
    Set fso = Nothing

End Function

Public Function TargetCopyFullName(ByVal tcfn_wbk As Variant) As String
' ----------------------------------------------------------------------------
' Returns the Sync-Target-Workbook's working copy full name derived from the
' provided argument (tcfn_wbk) which may be a Workbook object or a string and
' regardless the provided argument identifies a Sync-Target-Workbook or the
' Sync-Target-Workbook's copy already.
' ----------------------------------------------------------------------------
    Dim str As String
    
    If TypeName(tcfn_wbk) = "Workbook" _
    Then str = tcfn_wbk.FullName _
    Else: str = tcfn_wbk
        
    If InStr(str, SYNC_TARGET_SUFFIX & ".xls") = 0 _
    Then TargetCopyFullName = Replace(str, ".xls", SYNC_TARGET_SUFFIX & ".xls") _
    Else TargetCopyFullName = str

End Function

Private Function TargetCopyIsOpened() As Boolean
    TargetCopyIsOpened = ActiveWorkbook.name Like "*" & SYNC_TARGET_SUFFIX & ".*"
End Function

Public Sub TargetCopyOpen(Optional ByVal tco_new As Boolean = True, _
                          Optional ByRef tco_wbk_result As Workbook)
' ----------------------------------------------------------------------------
' Establish a Sync-Target-Workbook's working copy, i.e. a copy of the
' Sync-Target-Workbook (asumed the ActiveWorkbook) with a SYNC_TARGET_SUFFIX
' in its name.
' When a Sync-Target-Workbook's working copy already exists it is deleted.
' The procedure is exclusively used to start a new synchronization!
' when the ActiveWorkbook is a Sync-Target-Workbook's working copy the
' procedure ends without any action.
' ----------------------------------------------------------------------------
    Dim fso         As New FileSystemObject
    Dim sTargetCopy As String
    Dim sTarget     As String
    Dim wbk         As Workbook
    
    sTarget = wsService.ServicedWorkbookFullName
    sTargetCopy = wsService.SyncTargetFullNameCopy
    
    If tco_new Then
        '~~ Delete an already existing working copy
        If fso.FileExists(sTargetCopy) Then
            If mWbk.IsOpen(sTargetCopy, wbk) Then
                wbk.Close False
            End If
            fso.DeleteFile (sTargetCopy)
        End If
        ' Save the provided Workbook under the copy name and set it as the current working copy
        Application.EnableEvents = False
        mWbk.GetOpen(sTarget).SaveAs FileName:=sTargetCopy, AccessMode:=xlExclusive
        Set tco_wbk_result = ActiveWorkbook
        Application.EnableEvents = True
    Else
        '~~ Continue with existing working copy
        Application.EnableEvents = False
        Set tco_wbk_result = mWbk.GetOpen(sTargetCopy)
        tco_wbk_result.Activate
        Application.EnableEvents = True
    End If
    
xt: Set fso = Nothing
    
End Sub

Private Function OpenedIsTarget() As Boolean
    OpenedIsTarget = Not ActiveWorkbook.name Like "*" & SYNC_TARGET_SUFFIX & ".*"
End Function

Private Function OpenedIsTargetCopy() As Boolean
    OpenedIsTargetCopy = ActiveWorkbook.name Like "*" & SYNC_TARGET_SUFFIX & ".*"
End Function

Public Function TargetOriginFullName(ByVal tofn_wbk As Variant) As String
' ----------------------------------------------------------------------------
' Returns the Sync-Target-Workbook's full name derived from the provided
' argument (tofn_wbk) which may be a Workbook object or a string and
' regardless the provided argument identifies a Sync-Target-Workbook or a
' Sync-Target-Workbook's copy.
' ----------------------------------------------------------------------------
    Dim str As String
    
    If TypeName(tofn_wbk) = "Workbook" _
    Then str = tofn_wbk.FullName _
    Else str = tofn_wbk
        
    If InStr(str, SYNC_TARGET_SUFFIX & ".xls") <> 0 _
    Then TargetOriginFullName = Replace(str, SYNC_TARGET_SUFFIX & ".xls", ".xls") _
    Else TargetOriginFullName = str

End Function

Public Sub UnloadSyncMessage(ByVal sm_title As String)
    With wsService
        .SyncDialogTop = mMsg.MsgInstance(sm_title).Top
        .SyncDialogLeft = mMsg.MsgInstance(sm_title).Left
    End With
    mMsg.MsgInstance sm_title, True

End Sub

