Attribute VB_Name = "mSyncTest"
Option Explicit

Private Const TEST_SYNC_SOURCE = "E:\Ablage\Excel VBA\DevAndTest\Common-VBA-Excel-Component-Management-Services\SyncTest\SyncSource\CompManSyncTest.xlsb"
Private Const TEST_SYNC_TARGET = "E:\Ablage\Excel VBA\CompManSyncService\CompManSyncTest-Target\CompManSyncTest.xlsb"

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mSyncTest." & sProc
End Function

Public Sub Test_SyncRefs()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_SyncRefs"
    Const TEST_REF_1    As String = "Microsoft Visual Basic for Applications Extensibility 5.6"

    On Error GoTo eh
    Dim wbSource    As Workbook
    Dim TestRef1    As Reference
    
    mBasic.BoP ErrSrc(PROC)
    For Each TestRef1 In ThisWorkbook.VBProject.References
        If TestRef1.Description Like "*Extensibility*" Then Exit For
    Next TestRef1
    If TestRef1 Is Nothing Then Stop ' Error with test-Setup !!!
    
    mSync.SyncBackup TEST_SYNC_SOURCE
    mSync.SyncRestore TEST_SYNC_TARGET
    mSync.SyncBackup TEST_SYNC_TARGET
    
    '~~ Prepare the 'Synchronization-Source-Workbook for this test
    Application.EnableEvents = False
    Set wbSource = mWbk.GetOpen(TEST_SYNC_SOURCE)
    
    If Not mSyncRefs.RefExists(wbSource, "Microsoft Visual Basic for Applications Extensibility 5.6") Then
        wbSource.VBProject.References.AddFromGuid GUID:=TestRef1.GUID, Major:=TestRef1.Major, Minor:=TestRef1.Minor
    End If
    wbSource.Close
    Application.EnableEvents = True
    
    '~~ Prepare the 'Synchronization-Source-Workbook for this test
    
    '~~ Run the test
    Application.Run ThisWorkbook.Name & "!mCompMan." & SERVICE_SYNCHRONIZE, mWbk.GetOpen(TEST_SYNC_TARGET)

    mSync.SyncRestore TEST_SYNC_SOURCE

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_SyncColWidth()
' ---------------------------------------------------------------------
' Attention: This test preserves the target Workbook by a backup before
' and a restore after the synch test. The target Workbook thus will not
' show the synch result unless the terst procedire is stopped.
' ---------------------------------------------------------------------
    Const PROC = "Test_SyncColWidth"
    
    On Error GoTo eh
    Dim wbSource    As Workbook
    Dim WbTarget    As Workbook
    Dim ws          As Worksheet
    Dim sSheetName  As String
    
   
    mSync.SyncRestore TEST_SYNC_TARGET
    mSync.SyncBackup TEST_SYNC_TARGET
    
    Set WbTarget = mCompMan.WbkGetOpen(TEST_SYNC_TARGET)
    Set wbSource = mCompMan.WbkGetOpen(TEST_SYNC_SOURCE)
    
    For Each ws In wbSource.Worksheets
        If mSyncSheets.SheetExists(wb:=WbTarget _
                                 , sh1_name:=ws.Name _
                                 , sh1_code_name:=ws.CodeName _
                                 , sh2_name:=sSheetName _
                                  ) _
        Then
            mSyncRanges.SyncNamedColumnsWidth ws_source:=ws _
                                            , ws_target:=WbTarget.Worksheets(sSheetName)
        
        End If
    Next ws
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
 
Public Sub Test_SyncVBProjects()
' ---------------------------------------------------------------------
' Attention: This test preserves the target Workbook by a backup before
' and a restore after the synch test. The target Workbook thus will not
' show the synch result unless the terst procedire is stopped.
' ---------------------------------------------------------------------
    Const PROC = "Test_SyncVBProjects"
    
    On Error GoTo eh
                   
    Set mService.Serviced = mCompMan.WbkGetOpen(TEST_SYNC_TARGET)
    mSync.SyncRestore TEST_SYNC_TARGET
    
    mService.SyncVBProjects wb_target:=mCompMan.WbkGetOpen(TEST_SYNC_TARGET) _
                          , wb_source_name:=TEST_SYNC_SOURCE _
                          , restricted_sheet_rename_asserted:=True _
                          , design_rows_cols_added_or_deleted:=False
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_Synch_CompsChanged()
    Const PROC = "Test_Synch_CompsChanged"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim BttnDelete  As String
    Dim BttnKeep    As String
    Dim flLog       As File
    Dim LastModif   As Date
    Dim bttns       As Collection
    
    Set Stats = New clsStats
    Set Sync = New clsSync
    Set Log = New clsLog
    
    Set mService.Serviced = mCompMan.WbkGetOpen(TEST_SYNC_TARGET)
    Set Sync.TargetWb = mService.Serviced
    Set Sync.SourceWb = mCompMan.WbkGetOpen(TEST_SYNC_SOURCE)
    Log.File = mFile.Temp(mService.Serviced.Path, ".log")
    
    Sync.CollectAllSyncItems
    
    LastModif = fso.GetFile(TEST_SYNC_TARGET).DateLastModified
    
    mSync.SyncBackup TEST_SYNC_TARGET
    mSyncComps.SyncCodeChanges
    mSync.SyncRestore TEST_SYNC_TARGET
    Application.EnableEvents = False ' The open service UpdateOutdatedCommonComponents would start with a new log-file otherwise
    mCompMan.WbkGetOpen TEST_SYNC_TARGET
    Application.EnableEvents = True

xt: With New FileSystemObject
        If .FileExists(Log.File) Then
            Set flLog = .GetFile(Log.File)
            BttnDelete = "Delete Log-File" & vbLf & .GetFileName(Log.File)
            BttnKeep = "Keep Log-File" & vbLf & .GetFileName(Log.File)
            Set bttns = mMsg.Buttons(BttnDelete, BttnKeep)
            If mMsg.Box(Title:=PROC & " Log-File" _
                      , Prompt:=mFile.txt(.GetFile(Log.File)) _
                      , box_monospaced:=True _
                      , Buttons:=bttns) = BttnDelete _
            Then .DeleteFile flLog
        End If
    End With
    Set Log = Nothing

    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_Synch_RangesFormating()
    Const PROC = "Test_Synch_RangesFormating"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim BttnDelete  As String
    Dim BttnKeep    As String
    Dim flLog       As File
    Dim LastModif   As Date
    Dim bttns       As Collection
    
    Set Log = New clsLog
    Log.File = mFile.Temp(ThisWorkbook.Path, ".log")
    
    Set Stats = New clsStats
    Set Sync = New clsSync
    
    Set mService.Serviced = mCompMan.WbkGetOpen(TEST_SYNC_TARGET)
    Set Sync.TargetWb = mService.Serviced
    Set Sync.SourceWb = mCompMan.WbkGetOpen(TEST_SYNC_SOURCE)
    
    Sync.CollectAllSyncItems
    
    LastModif = fso.GetFile(TEST_SYNC_TARGET).DateLastModified
    
    mSync.SyncBackup TEST_SYNC_TARGET
    mSyncRanges.SyncFormating
    mSync.SyncRestore TEST_SYNC_TARGET
    Application.EnableEvents = False ' The open service UpdateOutdatedCommonComponents would start with a new log-file otherwise
    mCompMan.WbkGetOpen TEST_SYNC_TARGET
    Application.EnableEvents = True

xt: With New FileSystemObject
        If .FileExists(Log.File) Then
            Set flLog = .GetFile(Log.File)
            BttnDelete = "Delete Log-File" & vbLf & .GetFileName(Log.File)
            BttnKeep = "Keep Log-File" & vbLf & .GetFileName(Log.File)
            Set bttns = mMsg.Buttons(BttnDelete, BttnKeep)
            If mMsg.Box(Title:=PROC & " Log-File" _
                      , Prompt:=mFile.txt(.GetFile(Log.File)) _
                      , box_monospaced:=True _
                      , Buttons:=bttns) = BttnDelete _
            Then .DeleteFile flLog
        End If
    End With
    Set Log = Nothing

    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


