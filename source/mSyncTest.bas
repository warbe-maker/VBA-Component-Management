Attribute VB_Name = "mSyncTest"
Option Explicit

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mSyncTest." & sProc
End Function

Public Sub Test_SyncColWidth()
' ---------------------------------------------------------------------
' Attention: This test preserves the target Workbook by a backup before
' and a restore after the synch test. The target Workbook thus will not
' show the synch result unless the terst procedire is stopped.
' ---------------------------------------------------------------------
    Const PROC = "Test_SyncColWidth"
    
    On Error GoTo eh
    Dim sSource     As String
    Dim sTarget     As String
    Dim wbSource    As Workbook
    Dim wbTarget    As Workbook
    Dim ws          As Worksheet
    Dim sSheetName  As String
    
    sTarget = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncTarget\SyncTarget.xlsb"
    sSource = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncSource\SyncSource.xlsb"
    
    mSync.SyncRestore sTarget
    mSync.SyncBackup sTarget
    
    Set wbTarget = mCompMan.WbkGetOpen(sTarget)
    Set wbSource = mCompMan.WbkGetOpen(sSource)
    
    For Each ws In wbSource.Worksheets
        If mSyncSheets.SheetExists(wb:=wbTarget _
                                 , sh1_name:=ws.name _
                                 , sh1_code_name:=ws.CodeName _
                                 , sh2_name:=sSheetName _
                                  ) _
        Then
            mSyncRanges.SyncNamedColumnsWidth ws_source:=ws _
                                            , ws_target:=wbTarget.Worksheets(sSheetName)
        
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
    Dim sSource     As String
    Dim sTarget     As String
        
    sTarget = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncTarget\SyncTarget.xlsb"
    sSource = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncSource\SyncSource.xlsb"
           
    Set mService.Serviced = mCompMan.WbkGetOpen(sTarget)
    mSync.SyncRestore sTarget
    
    mService.SyncVBProjects wb_target:=mCompMan.WbkGetOpen(sTarget) _
                          , wb_source_name:=sSource _
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
    Dim sTarget     As String
    Dim sSource     As String
    Dim BttnDelete  As String
    Dim BttnKeep    As String
    Dim flLog       As File
    Dim LastModif   As Date
    Dim bttns       As Collection
    
    Set Stats = New clsStats
    Set Sync = New clsSync
    Set Log = New clsLog

    sTarget = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncTarget\SyncTarget.xlsb"
    sSource = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncSource\SyncSource.xlsb"
    
    Set mService.Serviced = mCompMan.WbkGetOpen(sTarget)
    Set Sync.Target = mService.Serviced
    Set Sync.source = mCompMan.WbkGetOpen(sSource)
    Log.File = mFile.Temp(mService.Serviced.Path, ".log")
    
    Sync.CollectAllSyncItems
    
    LastModif = fso.GetFile(sTarget).DateLastModified
    
    mSync.SyncBackup sTarget
    mSyncComps.SyncCodeChanges
    mSync.SyncRestore sTarget
    Application.EnableEvents = False ' The open service UpdateOutdatedCommonComponents would start with a new log-file otherwise
    mCompMan.WbkGetOpen sTarget
    Application.EnableEvents = True

xt: With New FileSystemObject
        If .FileExists(Log.File) Then
            Set flLog = .GetFile(Log.File)
            BttnDelete = "Delete Log-File" & vbLf & .GetFileName(Log.File)
            BttnKeep = "Keep Log-File" & vbLf & .GetFileName(Log.File)
            mMsg.Buttons bttns, BttnDelete, BttnKeep
            If mMsg.Box(Title:=PROC & " Log-File" _
                      , Prompt:=mFile.Txt(.GetFile(Log.File)) _
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
    Dim sTarget     As String
    Dim sSource     As String
    Dim BttnDelete  As String
    Dim BttnKeep    As String
    Dim flLog       As File
    Dim LastModif   As Date
    Dim bttns       As Collection
    
    Set Log = New clsLog
    Log.File = mFile.Temp(ThisWorkbook.Path, ".log")
    
    Set Stats = New clsStats
    Set Sync = New clsSync
    
    sTarget = "E:\Ablage\Excel VBA\DevAndTest\Test-Sync-Target-Project\Test_Sync_Target.xlsb"
    sSource = "E:\Ablage\Excel VBA\DevAndTest\Test-Sync-Source-Project\Test_Sync_Source.xlsb"
    
    Set mService.Serviced = mCompMan.WbkGetOpen(sTarget)
    Set Sync.Target = mService.Serviced
    Set Sync.source = mCompMan.WbkGetOpen(sSource)
    
    Sync.CollectAllSyncItems
    
    LastModif = fso.GetFile(sTarget).DateLastModified
    
    mSync.SyncBackup sTarget
    mSyncRanges.SyncFormating
    mSync.SyncRestore sTarget
    Application.EnableEvents = False ' The open service UpdateOutdatedCommonComponents would start with a new log-file otherwise
    mCompMan.WbkGetOpen sTarget
    Application.EnableEvents = True

xt: With New FileSystemObject
        If .FileExists(Log.File) Then
            Set flLog = .GetFile(Log.File)
            BttnDelete = "Delete Log-File" & vbLf & .GetFileName(Log.File)
            BttnKeep = "Keep Log-File" & vbLf & .GetFileName(Log.File)
            mMsg.Buttons bttns, BttnDelete, BttnKeep
            If mMsg.Box(Title:=PROC & " Log-File" _
                      , Prompt:=mFile.Txt(.GetFile(Log.File)) _
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


