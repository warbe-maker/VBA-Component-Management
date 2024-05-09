Attribute VB_Name = "mSync"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSync: Provides all services and means for the synchroni-
'                        zation of a Sync-Target- with its corresponding Sync-
'                        Source-Workbook.
' Public services:
' - AppRunSyncAll
' - ClearSyncData
' - Finalize
' - Initialize
' - MonitorStep
' - OpenDecision                 .
' - OpnDcsnPreparationTarget     .
' - OpnDcsnReSyncFromScratch     .
' - OpnDcsnSyncOpenWorkingCopy   .
' - OpnDcsnTargetSync            .
' - OpnDcsnWithExstngWorkingCopy .
' - RunSync                      Invoked via Application.Run
' - Source               Get/Let .
' - SourceClose                  .
' - SourceExists                 .
' - SourceOpen                   .
' - Target               Get/Let .
' - TargetArchive
' - TargetClearExportFiles
' - TargetClose
' - TargetOpen
' - TargetOriginFullName
' - TargetWorkingCopy Get/Let
' - TargetWorkingCopyClose
' - TargetWorkingCopyDelete
' - TargetWorkingCopyFullName
'
' W. Rauschenberger, Berlin Dec 2022
' ----------------------------------------------------------------------------
Public Const SYNC_ALL_BTTN                  As String = "Synchronize"                               ' Identifies the synchronization dialog
Public Const SYNC_TARGET_SUFFIX             As String = "_TargetWorkingCopy"                        ' suffix for the sync target working copy
Public Const SYNC_ACTION_REMOVE_OBSOLETE    As String = "Remove obsolete"
Public Const SYNC_ACTION_ADD_NEW            As String = "Add new"
Public Const SYNC_ACTION_CHANGE_CODE        As String = "Change Code "
Public Const SYNC_ACTION_CHANGE_NAME        As String = "Change Name "
Public Const SYNC_ACTION_CHANGE_CODENAME    As String = "Change CodeName "
Public Const SYNC_ACTION_CHANGE_SCOPE       As String = "Change Scope "
Public Const SYNC_ACTION_MULTIPLE_SOURCE    As String = "Multiple source "
Public Const SYNC_ACTION_MULTIPLE_TARGET    As String = "Multiple target "
Public Const SYNC_ACTION_OWNED_BY_PROJECT   As String = "Owned by VB-Project"
Public Const SYNC_ID_SEPARATOR              As String = " | "
Private Const TITLE_SYNC_ALL                As String = "Synchronize the VB-Project (by one click)"

Public cllAction                            As Collection
Public cllComment                           As Collection
Public cllDirection                         As Collection
Public cllId                                As Collection
Public cllKind                              As Collection
Public DueSyncKindOfObjects                 As New clsQ
Public lSyncMode                            As enSyncOption

Private AbortOpenDialogForPreparationCopy   As String
Private AbortOpenDlogForPreparationTarget   As String
Private cllDueSyncs                         As Collection
Private ContinueOpenWithTargetSynchronztn   As String
Private ContinueSyncOpenWorkingCopy         As String
Private ContinueSyncWithExstngWorkingCopy   As String
Private ContinueWithSynchAgainFromScratch   As String
Private OpnDcsnMsgTitle                     As String
Private wbkSyncSource                       As Workbook     ' the synchronization services global opened Sync-Source-Workbook (same name as the opened Sync-Target-Workbook)
Private wbkSyncTarget                       As Workbook     '  the synchronization services global opened Sync-Target-Workbook
Private wbkSyncTargetWorkingCopy            As Workbook     '  the synchronization services global opened Sync-Target-Workbook's working copy

Public Enum enSyncStep
    enSyncStepCollecting = 1
    enSyncStepSyncing
End Enum

Public Enum enSyncAction
    enSyncActionAddNew = 1
    enSyncActionChangeCode
    enSyncActionChangeCodeName
    enSyncActionChanged
    enSyncActionChangeName
    enSyncActionChangeNameAndScope
    enSyncActionChangeRefTo
    enSyncActionChangeRefToAndScope
    enSyncActionChangeScope
    enSyncActionChangeShapePrprty
    enSyncActionMultiple
    enSyncActionMultipleAdd
    enSyncActionMultipleRemove
    enSyncActionMultipleSource
    enSyncActionMultipleTarget
    enSyncActionOwnedByPrjct
    enSyncActionOwnedByPrjctObsolete
    enSyncActionRemoveObsolete
End Enum

Public Enum enSyncKindOfObject
    enSyncObjectKindReference = 1
    enSyncObjectKindWorksheet = 2
    enSyncObjectKindName = 3
    enSyncObjectKindVBComponent = 4
    enSyncObjectKindShape = 5
End Enum

Public Enum enSyncOption
    SyncSummarized
    SyncByKind
End Enum

Public Property Get Source() As Workbook
    Dim s As String
    
    On Error Resume Next
    s = wbkSyncSource.Name
    If Err.Number <> 0 Then
        Set wbkSyncSource = mWbk.GetOpen(wsService.SyncSourceFullName)
    End If
    Set Source = wbkSyncSource

End Property

Public Property Let Source(ByVal wbk As Workbook)
    Set wbkSyncSource = wbk
    wsService.SyncSourceFullName = wbk.FullName
End Property

Public Property Get Target() As Workbook
    Dim s As String
    
    On Error Resume Next
    s = wbkSyncTarget.Name
    If Err.Number <> 0 Then
        Set wbkSyncTarget = mWbk.GetOpen(wsService.CurrentServicedWorkbookFullName)
    End If
    Set Target = wbkSyncTarget
    Services.ServicedWbk = wbkSyncTarget
    
End Property

Public Property Let Target(ByVal wbk As Workbook):        Set wbkSyncTarget = wbk:      End Property

Public Property Get TargetWorkingCopy() As Workbook

    Dim s As String
    
    On Error Resume Next
    s = wbkSyncTarget.Name
    If Err.Number <> 0 Then
        Set wbkSyncTargetWorkingCopy = mWbk.GetOpen(wsService.SyncTargetFullNameCopy)
    End If
    Set TargetWorkingCopy = wbkSyncTargetWorkingCopy
    Services.ServicedWbk = wbkSyncTargetWorkingCopy
    
End Property

Public Property Let TargetWorkingCopy(ByVal wbk As Workbook):    Set wbkSyncTargetWorkingCopy = wbk:  End Property

Private Function AllDueSyncs(ByRef a_refs As Boolean, _
                             ByRef a_names As Boolean, _
                             ByRef a_sheets As Boolean, _
                             ByRef a_comps As Boolean, _
                             ByRef a_shapes As Boolean) As String
' ------------------------------------------------------------------------------
' Returns all currently known (collected) due synchronizations as string for
' ------------------------------------------------------------------------------
    Const PROC = "AllDueSyncs"
    
    On Error GoTo eh
    Dim s       As String
    Dim sRefs   As String
    Dim sComps  As String
    Dim sSheets As String
    Dim sNames  As String
    Dim sShapes As String
    
    sRefs = DueSyncs(enSyncObjectKindReference):    a_refs = sRefs <> vbNullString
    sSheets = DueSyncs(enSyncObjectKindWorksheet):  a_sheets = sSheets <> vbNullString
    sNames = DueSyncs(enSyncObjectKindName):        a_names = sNames <> vbNullString
    sComps = DueSyncs(enSyncObjectKindVBComponent): a_comps = sComps <> vbNullString
    sShapes = DueSyncs(enSyncObjectKindShape):      a_shapes = sShapes <> vbNullString
    
    s = sRefs & vbLf & sSheets & vbLf & sNames & vbLf & sComps
    While Left(s, 1) = vbLf
        s = Right(s, Len(s) - 1)
    Wend
    AllDueSyncs = s

xt:   Exit Function

eh:   Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AllDueSyncsDone(ByVal a_wbk_source As Workbook, _
                                 ByVal a_wbk_target As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when all all still due synchronizations are done.
' Note: The AllDone services remove the corresponding enumerated sync object
'       from the due syncs queue (cllDueSyncQ).
' ------------------------------------------------------------------------------
    Const PROC = "AllDueSyncsDone"
    
    mBasic.BoP ErrSrc(PROC)
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference) Then
        If Not mSyncRefs.AllDone(a_wbk_source, a_wbk_target) Then GoTo xt
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet) Then
        If Not mSyncSheets.AllDone(a_wbk_source, a_wbk_target) Then GoTo xt
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindName) Then
        If Not mSyncNames.AllDone(a_wbk_source, a_wbk_target) Then GoTo xt
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape) Then
        If Not mSyncShapes.AllDone(a_wbk_source, a_wbk_target) Then GoTo xt
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent) Then
        If Not mSyncComps.AllDone(a_wbk_source, a_wbk_target) Then GoTo xt
    End If
    AllDueSyncsDone = True

xt:   mBasic.EoP ErrSrc(PROC)
    Exit Function

End Function

Public Function AppRunNewIds(ByVal a_kind As enSyncKindOfObject) As String
    AppRunNewIds = DueSyncIdsByAction(a_kind, enSyncActionAddNew)
End Function

Public Function AppRunObsoleteIds(ByVal a_kind As enSyncKindOfObject) As String
    AppRunObsoleteIds = DueSyncIdsByAction(a_kind, enSyncActionRemoveObsolete)
End Function

Private Sub AppRunSyncAll()
' ------------------------------------------------------------------------------
' All due synchronizations are collected and confirmed.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunSyncAll"
    
    On Error GoTo eh
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkSource = mSync.Source
    Set wbkTarget = mSync.TargetWorkingCopy
    Services.ServicedWbk = wbkTarget
    Services.MessageUnload TITLE_SYNC_ALL ' allow watching the sync log
    
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference) Then
        mSyncRefs.AppRunSyncAll
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindName) Then
        mSyncNames.AppRunSyncAll
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet) Then
        mSyncSheets.AppRunSyncAll
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape) Then
        mSyncShapes.AppRunSyncAll
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent) Then
        mSyncComps.AppRunSyncAll
    End If
    
xt:   mBasic.EoP ErrSrc(PROC)
    mSync.RunSync
    Exit Sub

eh:   Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ClearSyncData()
    
    If CLng(Services.DialogLeft) < 5 Then Services.DialogLeft = 20
    If CLng(Services.DialogTop) < 5 Then Services.DialogTop = 20
    Application.ScreenUpdating = False
    wsSyncLog.Clear
    
End Sub

Private Sub CollectAll(ByVal c_wbk_source As Workbook, _
                       ByVal c_wbk_target As Workbook, _
              Optional ByRef c_terminated As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "CollectAll"
    
    mBasic.BoP ErrSrc(PROC)
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference) Then mSyncRefs.Collect c_wbk_source, c_wbk_target:            If c_terminated Then GoTo xt
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet) Then mSyncSheets.Collect c_wbk_source, c_wbk_target:          If c_terminated Then GoTo xt
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindName) Then mSyncNames.Collect c_wbk_source, c_wbk_target, c_terminated:  If c_terminated Then GoTo xt
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape) Then mSyncShapes.Collect c_wbk_source, c_wbk_target:              If c_terminated Then GoTo xt
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent) Then mSyncComps.Collect c_wbk_source, c_wbk_target:         If c_terminated Then GoTo xt
    mBasic.EoP ErrSrc(PROC)

xt:  Exit Sub
End Sub

Private Sub CollectDueKind(ByVal c_wbk_source As Workbook, _
                           ByVal c_wbk_target As Workbook, _
              Optional ByRef c_terminated As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "CollectDueKind"
    
    mBasic.BoP ErrSrc(PROC)
    Select Case True
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference):   mSyncRefs.Collect c_wbk_source, c_wbk_target
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet): mSyncSheets.Collect c_wbk_source, c_wbk_target
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindName):  mSyncNames.Collect c_wbk_source, c_wbk_target, c_terminated
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape): mSyncShapes.Collect c_wbk_source, c_wbk_target
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent):  mSyncComps.Collect c_wbk_source, c_wbk_target
    End Select
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Function Collected(ByVal c_kind As enSyncKindOfObject, _
                           ByVal c_action As enSyncAction) As Long
    c_kind = c_kind
    c_action = c_action
    
End Function

Public Sub CollectKnown(ByRef c_dct As Dictionary, _
                        ByVal c_id As String, _
               Optional ByVal c_v As Variant = Nothing)
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    If c_dct Is Nothing Then Set c_dct = New Dictionary
    If Not c_dct.Exists(c_id) Then
        If c_v Is Nothing _
        Then c_dct.Add c_id, c_id _
        Else c_dct.Add c_id, c_v
    End If
End Sub

Private Function DueSyncGet(ByVal d_index As Long, _
                  Optional ByRef d_sequence As String, _
                  Optional ByRef d_kind As enSyncKindOfObject, _
                  Optional ByRef d_action As enSyncAction, _
                  Optional ByRef d_direction As String, _
                  Optional ByRef d_id As String, _
                  Optional ByRef d_comment As String) As String
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    Const PROC = "DueSyncGet"
    
    On Error GoTo eh
    Static lMaxLenAction    As Long
    Static lMaxLenDir       As Long
    Static lMaxLenKind      As Long
    Static lMaxLenIdPart1   As Long
    Static lMaxLenIdPart2   As Long
    Dim cll                 As Collection
    Dim sNbsp               As String
    Dim v                   As Variant
    Dim s                   As String
    
    sNbsp = Services.NonBreakingSpace
    If d_index = 1 Then DueSyncMaxLenghts lMaxLenKind, lMaxLenAction, lMaxLenDir, lMaxLenIdPart1, lMaxLenIdPart2
    
    Set cll = cllDueSyncs(d_index)
    d_sequence = cll(1)
    d_kind = cll(2)
    d_action = cll(3)
    d_direction = cll(4)
    d_id = cll(5)
    d_comment = cll(6)
    
    s = mBasic.Align(SyncActionString(d_action), lMaxLenAction, AlignLeft, , sNbsp)
    s = s & mBasic.Align(d_direction, lMaxLenDir, AlignLeft, , sNbsp)
    s = s & mBasic.Align(SyncKindString(d_kind), lMaxLenKind, AlignLeft) & " "
    v = Split(d_id, SYNC_ID_SEPARATOR)
    s = s & mBasic.Align(v(0), lMaxLenIdPart1, AlignLeft) & "  "
    If UBound(v) >= 1 Then s = s & mBasic.Align(v(1), lMaxLenIdPart2, AlignLeft) & "  "
    If UBound(v) >= 2 Then s = s & v(2)

xt:  DueSyncGet = s
    Exit Function

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function
                 
Public Function DueSyncIdsByAction(ByVal d_kind As enSyncKindOfObject, _
                                   ByVal d_action As enSyncAction, _
                          Optional ByVal d_direction As String = vbNullString) As String
' ------------------------------------------------------------------------
' Returns a cooma delimited string of those SyncIds which match the
' provided kind (d_kind) - Name, Worksheet, etc. - and the provided
' action (d_action) - 'Remove obsolete' for instance.
' ------------------------------------------------------------------------
    Dim cll         As Collection
    Dim s           As String
    Dim v           As Variant
    Dim enAction    As enSyncAction
    Dim sComment    As String
    Dim sDirection  As String
    Dim sId         As String
    Dim enKind      As enSyncKindOfObject
    Dim sSequence   As String
    Dim sDelim      As String
    
    For Each v In cllDueSyncs
        Set cll = v
        sSequence = cll(1)
        enKind = cll(2)
        enAction = cll(3)
        sDirection = cll(4)
        sId = cll(5)
        sComment = cll(6)
        
        If enKind = d_kind Then
            If enAction = d_action Then
                s = s & sDelim & sId
                sDelim = ","
            ElseIf d_action = enSyncActionChanged And SyncActionIsChange(enAction) Then
                '~~ The action retrieved from the Collection (enAction) correlates with the requested action (d_action)
                If Trim(sDirection) = Trim(d_direction) Then
                    s = s & sDelim & sId
                    sDelim = ","
                End If
            End If
        End If
    Next v
    DueSyncIdsByAction = s

End Function

Private Function DuesyncIdsByThisAction(ByVal d_action_requested As enSyncAction, _
                                        ByVal d_action_retrieved As enSyncAction) As Boolean
' ----------------------------------------------------------------------------
' Retuns TRUE when the restieved sync action meets the requested sync action.
' ----------------------------------------------------------------------------
    DuesyncIdsByThisAction = d_action_retrieved = d_action_requested
    If Not DuesyncIdsByThisAction Then
        DuesyncIdsByThisAction = d_action_requested = enSyncActionChanged And SyncActionIsChange(d_action_retrieved)
    End If
End Function

Public Sub DueSyncLet(Optional ByVal d_sequence As String = vbNullString, _
                      Optional ByVal d_kind As enSyncKindOfObject = 0, _
                      Optional ByVal d_action As enSyncAction, _
                      Optional ByVal d_direction As String = vbNullString, _
                      Optional ByVal d_id As String, _
                      Optional ByVal d_comment As String = vbNullString)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "DueSyncLet"
    
    On Error GoTo eh
    Dim cll     As New Collection
    
    If cllDueSyncs Is Nothing Then Set cllDueSyncs = New Collection
    cll.Add d_sequence
    cll.Add d_kind
    cll.Add d_action
    cll.Add d_direction
    cll.Add d_id
    cll.Add d_comment
    cllDueSyncs.Add cll
    Set cll = Nothing
    
    Select Case d_kind
        Case enSyncObjectKindName
            Select Case d_action
                Case enSyncActionAddNew:           mSyncNames.KnownNew(d_id, SyncKindString(d_kind)) = True
                Case enSyncActionRemoveObsolete:   mSyncNames.KnownObsolete(d_id, SyncKindString(d_kind)) = True
                Case enSyncActionChangeName _
                   , enSyncActionChangeScope _
                   , enSyncActionChangeNameAndScope
                                                mSyncNames.KnownChanged(d_id, SyncKindString(d_kind)) = True
            End Select
        Case enSyncObjectKindReference
            Select Case d_action
                Case enSyncActionAddNew:            mSyncRefs.KnownNew(d_id) = True
                Case enSyncActionRemoveObsolete:    mSyncRefs.KnownObsolete(d_id) = True
            
            End Select
        Case enSyncObjectKindWorksheet
            Select Case d_action
                Case enSyncActionRemoveObsolete:    mSyncSheets.KnownObsolete(d_id) = True
                Case enSyncActionAddNew:            mSyncSheets.KnownNew(d_id) = True
                Case enSyncActionChangeName _
                   , enSyncActionChangeCodeName
                                                mSyncSheets.KnownChanged(d_id) = True
                Case enSyncActionOwnedByPrjct:    mSyncSheets.KnownOwnedByPrjct(d_id) = True
            End Select
        Case enSyncObjectKindShape
            Select Case d_action
                Case enSyncActionRemoveObsolete:    mSyncShapes.KnownObsolete(d_id) = True
                Case enSyncActionAddNew:            mSyncShapes.KnownNew(d_id) = True
                Case enSyncActionChangeName _
                   , enSyncActionChangeCodeName
                                                mSyncShapes.KnownChanged(d_id) = True
            End Select
        Case enSyncObjectKindVBComponent
            Select Case d_action
                Case enSyncActionRemoveObsolete:    mSyncComps.KnownObsolete(d_id) = True
                Case enSyncActionAddNew:            mSyncComps.KnownNew(d_id) = True
                Case enSyncActionChangeCode:        mSyncComps.KnownChanged(d_id) = True
            End Select
    End Select

xt:  Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub DueSyncMaxLenghts(ByRef d_max_length_kind As Long, _
                              ByRef d_max_length_action, _
                              ByRef d_max_length_dir As Long, _
                              ByRef d_max_length_id_part1 As Long, _
                              ByRef d_max_length_id_part2 As Long)
' ------------------------------------------------------------------------
' Returns the max lenghts of all due sync items.
' ------------------------------------------------------------------------
    Const PROC = "DueSyncMaxLenghts"
    
    On Error GoTo eh
    Dim i           As Long
    Dim enAction    As enSyncAction
    Dim sComment    As String
    Dim sDirection  As String
    Dim sId         As String
    Dim enKind      As enSyncKindOfObject
    Dim sSequence   As String
    Dim v           As Variant
    Dim cll         As Collection
    
    d_max_length_kind = 0
    d_max_length_action = 0
    d_max_length_dir = 0
    d_max_length_id_part1 = 0
    d_max_length_id_part2 = 0
    
    For i = 1 To cllDueSyncs.Count
        Set cll = cllDueSyncs(i)
        sSequence = cll(1)
        enKind = cll(2)
        enAction = cll(3)
        sDirection = cll(4)
        sId = cll(5)
        sComment = cll(6)
        
        d_max_length_kind = mBasic.Max(d_max_length_kind, Len(SyncKindString(enKind)) + 2)
        d_max_length_action = mBasic.Max(d_max_length_action, Len(SyncActionString(enAction)) + 1)
        d_max_length_dir = mBasic.Max(d_max_length_dir, Len(sDirection) + 2)
        v = Split(sId, SYNC_ID_SEPARATOR)
        d_max_length_id_part1 = mBasic.Max(d_max_length_id_part1, Len(v(0)))
        If UBound(v) >= 1 Then d_max_length_id_part2 = mBasic.Max(d_max_length_id_part2, Len(v(1)))
    Next i
    
xt:  Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function DueSyncKindOfObject(ByVal en As enSyncKindOfObject) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the enumerated kind of object is still in the
' DueSyncKindOfObjects queue.
' ------------------------------------------------------------------------------
    DueSyncKindOfObject = DueSyncKindOfObjects.IsQueued(en)
End Function

Public Function DueSyncs(ByVal d_kind As enSyncKindOfObject, _
                Optional ByRef d_number As Long) As String
' ------------------------------------------------------------------------------
' Returns the due synchronizations as a list for being displayed. when a kind
' (d_kind) is provided only of the kind, when an action (d_action) is provided
' only of that action
' ------------------------------------------------------------------------------
    Const PROC = "DueSyncs"
    
    On Error GoTo eh
    Dim i           As Long
    Dim s           As String
    Dim LineBreak   As String
    Dim sDueSync    As String
    Dim enKind      As enSyncKindOfObject
    Dim enAction    As enSyncAction
    Dim sDirection  As String
    
    mBasic.BoP ErrSrc(PROC)
    LineBreak = vbNullString
    For i = 1 To cllDueSyncs.Count
        sDueSync = DueSyncGet(i, , enKind, enAction, sDirection)
        If d_kind <> 0 Then
            If enKind <> d_kind Then GoTo nx
        End If
        If Trim(sDirection) = vbNullString Or Trim(sDirection) = "from" Then d_number = d_number + 1
        s = s & LineBreak & sDueSync
        LineBreak = vbLf
nx:  Next i
    
xt:  DueSyncs = s
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Private Sub Finalize()
' ------------------------------------------------------------------------------
' Displays a finalization dialog including a synchronization summary. When the
' finalization is confirmed the Sync-Target-Workbook's working copy is saved
' under its origin name and the working copy is deleted.
' ------------------------------------------------------------------------------
    Const PROC          As String = "Finalize"
    Const BTTN_FINALIZE As String = "Finalize the VB-Project-Synchronization"
    Const BTTN_ABORT    As String = "Terminate without finalization"
    
    On Error GoTo eh
    Dim MsgText         As mMsg.udtMsg
    Dim MsgButtons      As Collection
    Dim wbkSource       As Workbook
    Dim wbkTarget       As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.Source
    
    Set MsgButtons = mMsg.Buttons(BTTN_FINALIZE, vbLf, BTTN_ABORT)
    With MsgText
        With .Section(1)
            With .Label
                .Text = Replace(BTTN_FINALIZE, vbLf, " ")
                .FontColor = rgbDarkGreen
                .FontBold = True
            End With
            .Text.Text = "The Sync-Source-Workbook will be closed, the Sync-Target-Workbook's working copy " & _
                         "('" & wbkTarget.Name & "') will be copied under its origin name ('" & mSync.Target.Name & "') " & _
                         "and deleted. I.e. the productive Workbook originally moved to the configured synchronization " & _
                         "target folder ('" & wsConfig.FolderSyncTarget & "') is ready for being moved back to " & _
                         "production." & vbLf & _
                         "Note: A re-synchronization will still be possible when the most recent archive is " & _
                         "copied under its origin name and re-opened. However, changes like required manual " & _
                         "synchronizations will need to be done again."
        End With
        With .Section(2)
            With .Label
                .Text = Replace(BTTN_ABORT, vbLf, " ")
                .FontColor = rgbDarkGreen
                .FontBold = True
            End With
            .Text.Text = "With then termination of this dialog without finalizing the synchronization a re-open " & _
                         "of the Sync-Target-Workbook is required. When the open option 'continue ...' is selected " & _
                         "this finalization dialog will be displayed again."
        End With
        With .Section(8).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization service (GitHub README):"
            .FontColor = rgbBlue
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_SYNC_CHAPTER
        End With
    End With

    Select Case mMsg.Dsply(dsply_title:="Finalization of the synchronization" _
                         , dsply_msg:=MsgText _
                         , dsply_buttons:=MsgButtons)
        Case BTTN_FINALIZE
            FSo.DeleteFile mSync.TargetOriginFullName(mSync.TargetWorkingCopy)
            wbkTarget.SaveAs FileName:=mSync.TargetOriginFullName(mSync.TargetWorkingCopy)
            mSync.TargetWorkingCopyDelete
        Case BTTN_ABORT
            wbkSource.Close
            mSync.TargetWorkingCopy.Close False
    End Select
    
xt:  mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub InitDueSyncs()
    Set cllDueSyncs = Nothing:  Set cllDueSyncs = New Collection
End Sub

Public Sub Initiate(Optional ByVal i_sync_refs As Boolean = True, _
                    Optional ByVal i_sync_sheets As Boolean = True, _
                    Optional ByVal i_sync_names As Boolean = True, _
                    Optional ByVal i_sync_shapes As Boolean = False, _
                    Optional ByVal i_sync_comps As Boolean = True)
' ------------------------------------------------------------------------------
' Initializes the service after a Sync-Target-Workbook had been opened.
' Cases distinguished:
' - An origin Sync-Target-Workbook has been opened:
'   When a version suffixed _SyncTarget_WorkingCopy does not exist, the
'   Sync-Target-Workbook is saved as a version suffixed _SyncTarget_WorkingCopy,
'   else a desision dialog is opened providing the choices "re-sync from scratch",
'   "continue with ongoing synch", "terminate in order to enable manual
'   modifications on the Sync-Target-Workbook.
' - A suffixed _SyncTarget_WorkingCopy:
'   When identical with last performed - and inerrupted
'                            sync continue it
' ------------------------------------------------------------------------------
    With LogServiced
        .AlignmentItems "|L|L|L|L|L|"
        .MaxItemLengths 10, Comps.MaxLenServicedType, 90, 15, 150
        .Headers "|  Due  | Item | Item | Sync | Comment  |"
        .Headers "| Sync  | Type | Id   |Result| |"
    End With
    
    Set DueSyncKindOfObjects = Nothing
    Set DueSyncKindOfObjects = New clsQ
    
    '~~ Determining the sequence in which the synchronizations are performed by enqueing.
    If i_sync_refs Then DueSyncKindOfObjects.EnQueue enSyncObjectKindReference
    If i_sync_sheets Then DueSyncKindOfObjects.EnQueue enSyncObjectKindWorksheet
    If i_sync_names Then DueSyncKindOfObjects.EnQueue enSyncObjectKindName
    If i_sync_shapes Then DueSyncKindOfObjects.EnQueue enSyncObjectKindShape
    If i_sync_comps Then DueSyncKindOfObjects.EnQueue enSyncObjectKindVBComponent
    
    wsService.SyncComponents = DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent)
    wsService.SyncNames = DueSyncKindOfObjects.IsQueued(enSyncObjectKindName)
    wsService.SyncReferences = DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference)
    wsService.SyncShapes = DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape)
    wsService.SyncWorksheets = DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet)
    
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference) Then
        wsSyncLog.SummaryDone("Reference") = "due"
    Else
        wsSyncLog.SummaryDone("Reference") = "'-"
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet) Then
        wsSyncLog.SummaryDone("Worksheet") = "due"
        mSyncSheets.Initialize
    Else
        wsSyncLog.SummaryDone("Worksheet") = "'-"
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindName) Then
        wsSyncLog.SummaryDone(enSyncObjectKindName) = "due"
        mSyncNames.Initialize
    Else
        wsSyncLog.SummaryDone(enSyncObjectKindName) = "'-"
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape) Then
        wsSyncLog.SummaryDone("Shape") = "due"
        mSyncShapes.Initialize
    Else
        wsSyncLog.SummaryDone("Shape") = "'-"
    End If
    If DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent) Then
        wsSyncLog.SummaryDone("VBComponent") = "due"
        mSyncComps.Initialize
    Else
        wsSyncLog.SummaryDone("VBComponent") = "'-"
    End If
    
End Sub

Public Sub MonitorStep(ByVal ms_text As String)
    
    On Error Resume Next
    ActiveWindow.WindowState = xlMaximized
    mCompManClient.Progress p_service_name:=mCompManClient.SRVC_SYNCHRONIZE _
                          , p_by_servicing_wbk_name:=ThisWorkbook.Name _
                          , p_serviced_wbk_name:=Services.ServicedWbk.Name _
                          , p_service_info:=ms_text
    
End Sub

Public Sub OpenDecision()
' ------------------------------------------------------------------------------
' Displays a mode-less dialog with open-decision-buttons.
' ------------------------------------------------------------------------------
    Const PROC  As String = "OpenDecision"
    
    On Error GoTo eh
    Dim MsgText     As mMsg.udtMsg
    Dim MsgButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    
    '~~ Prepare decision option button captions
    ContinueOpenWithTargetSynchronztn = "Synchronize"
    ContinueWithSynchAgainFromScratch = "Re-Synchronize"
    ContinueSyncWithExstngWorkingCopy = "Continue Synchronization with" & vbLf & _
                                        "existing Sync-Target-Workbook's working copy"
    ContinueSyncOpenWorkingCopy = "Continue ongoing Synchronization with" & vbLf & _
                                        "opened Sync-Target-Workbook's working copy"
    AbortOpenDlogForPreparationTarget = "Stop for pre-synchronization preparations"
    
    If TargetWorkingCopyIsOpened _
    Then OpnDcsnMsgTitle = "Decide about the opened Sync-Target-Workbook's working copy" _
    Else OpnDcsnMsgTitle = "Decide about the opened Sync-Target-Workbook"

    With MsgText
        If OpenedIsTarget Then ' The Sync-Target-Wokbook is opened
            With .Section(1)
                If TargetWorkingCopyExists Then
                    '~~ Re-Synchronize (again from scratch)
                    .Label.Text = Replace(ContinueWithSynchAgainFromScratch, vbLf, " ") & ":"
                    .Label.FontColor = rgbDarkGreen
                    .Text.Text = "Although there's an ongoing, still un-finalized synchronization going on " & _
                                 "(indicated by an existing Sync-Target-Workbook's working copy) the " & _
                                 "synchronization will restarted from scratch. I.e. the Sync-Target-Workbook " & _
                                 "will again be archived and the already existing Sync-Target-Workbook's " & _
                                 "working copy will be ignored."
                    Set MsgButtons = mMsg.Buttons(ContinueWithSynchAgainFromScratch)
                    mMsg.BttnAppRun AppRunArgs _
                                  , ContinueWithSynchAgainFromScratch _
                                  , ThisWorkbook _
                                  , "mSync.OpnDcsnReSyncFromScratch"
                Else
                    '~~ Synchronize
                    .Label.Text = Replace(ContinueOpenWithTargetSynchronztn, vbLf, " ") & ":"
                    .Label.FontColor = rgbDarkGreen
                    .Text.Text = "The opened Sync-Target-Workbook will be archived and its VB-Project will " & _
                                 "be synchronized with its corresponding Sync-Source-Workbook. For this - " & _
                                 "because Source- and Target-Workbook have the same name - the Sync-Target-Workbook " & _
                                 "is saved as Sync-Target-Workbook working copy and this one will be synchronized."
                    Set MsgButtons = mMsg.Buttons(ContinueOpenWithTargetSynchronztn)
                    mMsg.BttnAppRun AppRunArgs _
                                  , ContinueOpenWithTargetSynchronztn _
                                  , ThisWorkbook _
                                  , "mSync.OpnDcsnTargetSync"
                End If
            End With
            If TargetWorkingCopyExists Then
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
                mMsg.BttnAppRun AppRunArgs _
                                , ContinueSyncWithExstngWorkingCopy _
                                , ThisWorkbook _
                                , "mSync.OpnDcsnWithExstngWorkingCopy"
            End If
            With .Section(3)
                .Label.Text = Replace(AbortOpenDlogForPreparationTarget, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The Sync-Target-Workbook is kept open for possibly still required " & _
                             "pre-synchronization preparations. When the Sync-Target-Workbook is " & _
                             "closed thereafter and re-opened any other decision may be taken."
            End With
            Set MsgButtons = mMsg.Buttons(MsgButtons, vbLf, AbortOpenDlogForPreparationTarget)
            mMsg.BttnAppRun AppRunArgs _
                            , AbortOpenDlogForPreparationTarget _
                            , ThisWorkbook _
                            , "mSync.OpnDcsnPreparationTarget"
        
        ElseIf TargetWorkingCopyIsOpened Then
            '~~ When the Sync-Target-Workbook's working copy is opened this likely means that
            '~~ the synchronization had already be started and may have passed som steps already.
            '~~ The option to prepare the working copy is not recommendable and thus is not offered.
            With .Section(1)
                '~~ Continue with opened working copy
                .Label.Text = Replace(ContinueSyncOpenWorkingCopy, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The opened Sync-Target-Workbook's working copy is used to continue a yet " & _
                             "unfinished (or just pending finalization) synchronization. A finalization " & _
                             "dialog will be displayed once all synchronizations are done."

            End With
            Set MsgButtons = mMsg.Buttons(ContinueSyncOpenWorkingCopy)
            mMsg.BttnAppRun AppRunArgs _
                            , ContinueSyncOpenWorkingCopy _
                            , ThisWorkbook _
                            , "mSync.OpnDcsnSyncOpenWorkingCopy"
            
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
            mMsg.BttnAppRun AppRunArgs _
                            , ContinueWithSynchAgainFromScratch _
                            , ThisWorkbook _
                            , "mSync.OpnDcsnReSyncFromScratch"
            
            With .Section(3)
                '~~ Prepare Sync-Target-Workbook
                .Label.Text = Replace(AbortOpenDialogForPreparationCopy, vbLf, " ") & ":"
                .Label.FontColor = rgbDarkGreen
                .Text.Text = "The Sync-Target-Workbook's working copy will closed without save and removed and " & _
                             "the Sync-Target-Workbook will be opened for a manual pre-sync-preparation instead. " & _
                             "When done the Sync-Target-Workbook will be closed and re-opened for being synchronized."
            End With
            
            Set MsgButtons = mMsg.Buttons(MsgButtons, vbLf, AbortOpenDlogForPreparationTarget)
            mMsg.BttnAppRun AppRunArgs _
                            , AbortOpenDlogForPreparationTarget _
                            , ThisWorkbook _
                            , "mSync.OpnDcsnPreparationTarget"
            
        End If ' Target copy is opened
    
        With .Section(8).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization service (GitHub README):"
            .FontColor = rgbBlue
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_SYNC_CHAPTER
        End With
    End With
    
    '~~ Display the mode-less open decision dialog
    mMsg.Dsply dsply_title:=OpnDcsnMsgTitle _
             , dsply_msg:=MsgText _
             , dsply_Label_spec:="R70" _
             , dsply_buttons:=MsgButtons _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=AppRunArgs _
             , dsply_width_min:=45 _
             , dsply_pos:=Services.DialogTop & ";" & Services.DialogLeft
                            
xt:  Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function OpenedIsTarget() As Boolean
    OpenedIsTarget = Not ActiveWorkbook.Name Like "*" & SYNC_TARGET_SUFFIX & ".*"
End Function

Private Function OpenedIsTargetWorkingCopy() As Boolean
    OpenedIsTargetWorkingCopy = ActiveWorkbook.Name Like "*" & SYNC_TARGET_SUFFIX & ".*"
End Function

Private Sub OpnDcsnPreparationTarget()
' ------------------------------------------------------------------------------
' When the Sync-Target-Workbook had been opened it simply will remain open.
' When the Sync-Target-Workbook's working copy had been opend:
' - it will be closed (without save!) and removed (a prepared Sync-Target-
'   Workbook will require a sync from scratch.
' - The Sync-Source-Workbook will be closed and the Sync-Target-Workbook will be
'   opened instead.
' ------------------------------------------------------------------------------
    Const PROC = "OpnDcsnPreparationTarget"
        
    mMsg.MsgInstance OpnDcsnMsgTitle, True ' Close/unload the mode-less open dialog
    If OpenedIsTarget Then
        With FSo
            If .FileExists(wsService.SyncTargetFullNameCopy) _
            Then .DeleteFile wsService.SyncTargetFullNameCopy
        End With
    ElseIf OpenedIsTargetWorkingCopy Then
        ActiveWorkbook.Close False
        mSync.SourceClose
        With FSo
            If .FileExists(wsService.SyncTargetFullNameCopy) _
            Then .DeleteFile wsService.SyncTargetFullNameCopy
        End With
        mCompManClient.Events ErrSrc(PROC), False
        mSync.TargetOpen
        mCompManClient.Events ErrSrc(PROC), True
        
    End If

End Sub

Private Sub OpnDcsnReSyncFromScratch()
' ------------------------------------------------------------------------------
' Resynchronizer the opened Sync-Target-Workbook although a working copy already
' exists.
' ------------------------------------------------------------------------------
    Const PROC = "OpnDcsnTargetSync"
    
    On Error GoTo xt
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpnDcsnMsgTitle, True ' Close/teminate the mode-less open dialog
    mSync.SourceClose                           ' Close a possibly already open Sync-Source-Workbook
    mSync.TargetArchive                         ' Archive the Sync-Target-Workbook
    mSync.TargetWorkingCopyOpen                        ' Establish a new working copy
    mSync.SourceOpen                            ' Open the corresponding Sync-Source-Workbook
    mSync.ClearSyncData                         ' Clear any data from a previous synchronization
    mSync.TargetClearExportFiles
    mSync.SyncMode
    mSync.RunSync

xt:  mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OpnDcsnSyncOpenWorkingCopy()
' ------------------------------------------------------------------------------
' A Sync-Target-Workbook's working copy has directly been opened
' ------------------------------------------------------------------------------
    Const PROC = "OpnDcsnSyncOpenWorkingCopy"
    
    On Error GoTo xt
    Dim wbk             As Workbook
    Dim wbkTargetWorkingCopy   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpnDcsnMsgTitle, True
    mSync.Target = mWbk.GetOpen(wsService.CurrentServicedWorkbookFullName)
    mSync.TargetWorkingCopyOpen False, wbkTargetWorkingCopy
    mSync.TargetWorkingCopy = wbkTargetWorkingCopy
    If mWbk.IsOpen(wsService.CurrentServicedWorkbookFullName, wbk) Then
        mSync.TargetClose ' Close any still open Sync-Target-Workbook
    End If
    mSync.SourceOpen  ' Open the corresponding Sync-Source-Workbook
    mSync.SyncMode
    mSync.RunSync

xt:  mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OpnDcsnTargetSync()
' ------------------------------------------------------------------------------
' Syncronize the opened Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "OpnDcsnTargetSync"
    
    On Error GoTo xt
    
    mBasic.BoP ErrSrc(PROC)
    mMsg.MsgInstance OpnDcsnMsgTitle, True
    mSync.TargetArchive                         ' Archive the Sync-Target-Workbook
    mSync.TargetWorkingCopyOpen                        ' Establish a new Sync-Target-Workbook working copy
    mSync.SourceOpen                            ' Open the corresponding Sync-Source-Workbook
    mSync.ClearSyncData                         ' Clear any data from a previous synchronization
    mSync.TargetClearExportFiles
    mSync.SyncMode
    mSync.RunSync

xt:  mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Progress(ByVal p_kind As enSyncKindOfObject, _
                    ByVal p_sync_step As enSyncStep, _
                    ByVal p_sync_action As enSyncAction, _
                    ByVal p_count As Long)
' ------------------------------------------------------------------------------
' Returns a string indcation the sync progress
' <kind> <action> [<sub-action> (n of m)] [ ] ...
' Example 1: Name Collect New(n of m), Obsolete(n of m), Changes(n of m), Multiple(n of m)
' Example 1: Name Sync New(n of m), Obsolete(n of m), Changed(n of m), Multiple(n of m)
' ------------------------------------------------------------------------------
    Const PROC = "Progress"
    
    On Error GoTo eh
    Static enKind           As enSyncKindOfObject
    Static enStep           As enSyncStep
    Static enAction         As enSyncAction
    Static sNew             As String
    Static sObsolete        As String
    Static sChanged         As String
    Static sMultiple        As String
    Static lCountNew        As Long
    Static lCountObsolete   As Long
    Static lCountChanged    As Long
    Static lCountMultiple   As Long
    
    Dim lOf                 As Long
    Dim sStepDetails        As String
    Dim sProgressMsg        As String
    
    mBasic.BoP ErrSrc(PROC)
    If p_kind <> enKind Or p_sync_step <> enStep Or p_sync_action <> enAction Then
        lOf = ProgressOf(p_kind, p_sync_step, p_sync_action)
    End If
    
    If p_sync_step <> enAction Then
        sProgressMsg = vbNullString
        sNew = vbNullString
        sObsolete = vbNullString
        sChanged = vbNullString
        sMultiple = vbNullString
        enAction = p_sync_step
    End If
    
    If p_kind <> enKind Then
        sNew = vbNullString
        sObsolete = vbNullString
        sChanged = vbNullString
        sMultiple = vbNullString
        enKind = p_kind
    End If
        
    Select Case True
        Case p_sync_action = enSyncActionAddNew:            lCountNew = p_count:        sNew = "new (" & p_count & " of " & lOf & ")"
        Case p_sync_action = enSyncActionRemoveObsolete:    lCountObsolete = p_count:   sObsolete = "obsolete (" & p_count & " of " & lOf & ")"
        Case SyncActionIsChange(p_sync_action):             lCountChanged = p_count:    sChanged = "changed (" & p_count & " of " & lOf & ")"
        Case SyncActionIsMultiple(p_sync_action):           lCountMultiple = p_count:   sMultiple = "multiple (" & p_count & " of " & lOf & ")"
    End Select
        
    sStepDetails = Trim(sChanged & ", " & sNew & ", " & sObsolete & ", " & sMultiple)
    While Left(sStepDetails, 1) = ","
        sStepDetails = Trim(Right(sStepDetails, Len(sStepDetails) - 1))
    Wend
    While Right(sStepDetails, 1) = ","
        sStepDetails = Trim(Left(sStepDetails, Len(sStepDetails) - 1))
    Wend
    sProgressMsg = SyncStepString(p_sync_step) & " " & SyncKindString(enKind) & " " & sStepDetails
    Application.StatusBar = sProgressMsg
    
xt:  mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ProgressOf(ByVal p_kind As enSyncKindOfObject, _
                            ByVal p_sync_step As enSyncStep, _
                            ByVal p_sync_action As enSyncAction) As Long
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "ProgressOf"
    
    On Error GoTo eh
    Dim lOfSource   As Long
    Dim lOfTarget   As Long
    
    mBasic.BoP ErrSrc(PROC)
    Select Case p_sync_step
        Case enSyncStepCollecting
            Select Case p_kind
                Case enSyncObjectKindName:          lOfSource = mSync.Source.Names.Count
                                                    lOfTarget = mSync.TargetWorkingCopy.Names.Count
                Case enSyncObjectKindWorksheet:     lOfSource = mSync.Source.Worksheets.Count
                                                    lOfTarget = mSync.TargetWorkingCopy.Worksheets.Count
                Case enSyncObjectKindVBComponent:   lOfSource = mSync.Source.VBProject.VBComponents.Count
                                                    lOfTarget = mSync.TargetWorkingCopy.VBProject.VBComponents.Count
                Case enSyncObjectKindReference:     lOfSource = mSync.Source.VBProject.References.Count
                                                    lOfTarget = mSync.TargetWorkingCopy.VBProject.References.Count
                Case enSyncObjectKindShape:         lOfSource = mSyncShapes.NoOfShapes(mSync.Source)
                                                    lOfTarget = mSyncShapes.NoOfShapes(mSync.TargetWorkingCopy)
            End Select
            
            Select Case p_sync_action
                Case enSyncActionAddNew:    ProgressOf = lOfSource
                Case Else:                  ProgressOf = lOfTarget
            End Select
        
        Case enSyncStepSyncing
                                            ProgressOf = Collected(p_kind, p_sync_action)
    End Select
    
    If SyncActionIsChange(p_sync_action) And p_kind <> enSyncObjectKindVBComponent Then
        ProgressOf = ProgressOf / 2
    End If
    
xt:  mBasic.EoP ErrSrc(PROC)
    Exit Function

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub RunSync()
' ------------------------------------------------------------------------------
' Performs all due synchronizations for due synchronization objects (those in
' the cllDueSyncQ. Because the individual object type specific synchronizations
' are completely independent their processing sequence is arbitrary.
'
' ------------------------------------------------------------------------------
    Const PROC = "RunSync"
    
    On Error GoTo eh
    Dim bTerminated As Boolean
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkSource = mSync.Source
    Set wbkTarget = mSync.TargetWorkingCopy
    
    Services.ServicedWbk = wbkTarget
    Set cllDueSyncs = Nothing: Set cllDueSyncs = New Collection
    
    If AllDueSyncsDone(wbkSource, wbkTarget) Then
        mSync.Finalize
        GoTo xt
    End If
                
    If lSyncMode = SyncByKind Then
        CollectDueKind wbkSource, wbkTarget, bTerminated
        SyncKind wbkSource, wbkTarget, bTerminated
    Else
        '~~ Perform all due synchronizations summarized. I.e. all are displayed
        '~~ and performed by one click
        CollectAll wbkSource, wbkTarget, bTerminated
        If bTerminated Then GoTo xt
        Application.ScreenUpdating = False
        SyncAll
    End If
    
xt:  mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AppRunInit()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    wsSyncLog.Activate
End Sub

Public Sub AppRunTerminate()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    wsSyncLog.Range("rngDone").EntireColumn.AutoFit
    Application.ScreenUpdating = False
End Sub

Private Sub SourceClose()
' ------------------------------------------------------------------------------
' Closes an open Sync-Source-Workbook by considering still unsaved change.
' ------------------------------------------------------------------------------
    Const PROC = "SourceClose"
    
    Dim wbk As Workbook
    Dim sSourceFullName As String
    
    sSourceFullName = wsService.SyncSourceFullName
    If mWbk.IsOpen(sSourceFullName, wbk) Then
        If wbk.FullName = sSourceFullName Then
            mCompManClient.Events ErrSrc(PROC), False
            If wbk.Saved _
            Then wbk.Close False _
            Else wbk.Close True
            mCompManClient.Events ErrSrc(PROC), True
        End If
    End If

End Sub

Public Function SourceExists(ByVal se_wbk_opened As Workbook) As Boolean
' ------------------------------------------------------------------------------
' When the Sync-Source-Workbook unambigously exists in the Serviced-CompManRoot-
' Folder the function returns TRUE and saves the found Workbook's full name to
' wsService.SyncSourceFullName, else the function displays a corresponding
' message and returns FALSE.
' ------------------------------------------------------------------------------
    Dim cll             As Collection
    Dim sSourceFullName As String
    Dim Msg             As mMsg.udtMsg
    Dim MsgTitle        As String
    Dim i               As Long
    
    mFso.Exists x_folder:=wsConfig.FolderCompManRoot _
              , x_file:=Replace(se_wbk_opened.Name, SYNC_TARGET_SUFFIX & ".xls", ".xls") _
              , x_result_files:=cll
    Select Case cll.Count
        Case 0
            MsgTitle = "No corresponding Sync-Source-Workbook found!"
            With Msg.Section(1)
                .Text.Text = "No correponding Sync-Source-Workbook for the opened Sync-Target-Workbook " & _
                             "(" & se_wbk_opened.Name & ") could be found in the configured 'ServicedCompManRoot' folder " & _
                             "(" & wsConfig.FolderCompManRoot & ")."
            End With
        Case 1
            sSourceFullName = cll(1)
        Case Else
            MsgTitle = "Ambigous Sync-Source-Workbooks found!"
            With Msg
                .Section(1).Text.Text = "For the opened Sync-Target-Workbook (" & se_wbk_opened.Name & ") ambigous " & _
                                        "corresponding Sync-Source-Workbooks had been found in the configured " & _
                                        "'ServicedCompManRoot' folder (" & wsConfig.FolderCompManRoot & "):"
                With .Section(2).Text
                    .MonoSpaced = True
                    .FontSize = 8
                    .Text = cll(1)
                    For i = 2 To cll.Count
                        .Text = .Text & vbLf & cll(i)
                    Next i
                End With
                .Section(3).Text.Text = "Terminate this synchronization trial, first remove the additional Workbook's or move them outside the " & _
                             wsConfig.FolderCompManRoot & " folder and re-open the Sync-Target-Workbook."
            End With
    End Select
            
    If sSourceFullName = vbNullString Then
        With Msg.Section(8).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization service (GitHub README):"
            .FontColor = rgbBlue
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_SYNC_CHAPTER
        End With
        mMsg.Dsply dsply_title:=MsgTitle _
                 , dsply_msg:=Msg _
                 , dsply_Label_spec:="R70" _
                 , dsply_buttons:=mMsg.Buttons("Terminate Synchronization") _
                 , dsply_width_min:=30
        wsService.SyncSourceFullName = vbNullString
    Else
        SourceExists = True
        wsService.SyncSourceFullName = sSourceFullName
    End If
    
End Function

Private Sub SourceOpen()
    Const PROC = "SourceOpen"
    
    mCompManClient.Events ErrSrc(PROC), False
    Set wbkSyncSource = mWbk.GetOpen(wsService.SyncSourceFullName)
    mCompManClient.Events ErrSrc(PROC), True

End Sub

Public Function SyncActionIsChange(ByVal s_sync_action As enSyncAction) As Boolean
    Select Case s_sync_action
        Case enSyncActionChangeCode, _
             enSyncActionChangeCodeName, _
             enSyncActionChanged, _
             enSyncActionChangeName, _
             enSyncActionChangeNameAndScope, _
             enSyncActionChangeRefTo, _
             enSyncActionChangeRefToAndScope, _
             enSyncActionChangeScope, _
             enSyncActionChangeShapePrprty
             SyncActionIsChange = True
    End Select
End Function

Public Function SyncActionIsMultiple(ByVal s_sync_action) As Boolean
    Select Case s_sync_action
        Case enSyncActionMultiple _
           , enSyncActionMultipleAdd _
           , enSyncActionMultipleRemove _
           , enSyncActionMultipleSource _
           , enSyncActionMultipleTarget
        SyncActionIsMultiple = True
    End Select
End Function

Public Function SyncActionString(ByVal en As enSyncAction) As String
    Select Case en
        Case enSyncActionAddNew:                SyncActionString = "Add new"
        Case enSyncActionChangeCode:            SyncActionString = "Change Code"
        Case enSyncActionChangeCodeName:        SyncActionString = "Change CodeName"
        Case enSyncActionChanged:               SyncActionString = "Changed"
        Case enSyncActionChangeName:            SyncActionString = "Change Name"
        Case enSyncActionChangeNameAndScope:    SyncActionString = "Change Name and Scope "
        Case enSyncActionChangeRefTo:           SyncActionString = "Change RefersTo"
        Case enSyncActionChangeRefToAndScope:   SyncActionString = "Change RefersTo and Scope"
        Case enSyncActionChangeScope:           SyncActionString = "Change Scope"
        Case enSyncActionChangeShapePrprty:     SyncActionString = "Change Property"
        Case enSyncActionMultiple:              SyncActionString = "Multiple"
        Case enSyncActionMultipleAdd:           SyncActionString = "Multiple add"
        Case enSyncActionMultipleRemove:        SyncActionString = "Multiple remove"
        Case enSyncActionMultipleSource:        SyncActionString = "Multiple source"
        Case enSyncActionMultipleTarget:        SyncActionString = "Multiple target"
        Case enSyncActionOwnedByPrjct:          SyncActionString = "Owned by VB-Project"
        Case enSyncActionOwnedByPrjctObsolete:  SyncActionString = "Owned by VB-Project obsolete"
        Case enSyncActionRemoveObsolete:        SyncActionString = "Remove obsolete"
    End Select
End Function

Private Sub SyncAll()
' ------------------------------------------------------------------------------
' Displays all collected due synchronizations for being executed by one click -
' which calls AppRunSyncAll.
' ------------------------------------------------------------------------------
    Const PROC = "SyncAll"
    
    On Error GoTo eh
    Dim fSync       As fMsg
    Dim bDueRefs    As Boolean
    Dim bDueNames   As Boolean
    Dim bDueSheets  As Boolean
    Dim bDueComps   As Boolean
    Dim bDueShapes  As Boolean
    Dim lSection    As Long
    Dim Msg         As mMsg.udtMsg
    Dim Bttn1       As String
    Dim Bttn2       As String
    Dim cllButtons  As Collection
    Dim AppRunArgs  As New Dictionary
    Dim sDueSyncs   As String
    
    mBasic.BoP ErrSrc(PROC)
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_ALL)
    sDueSyncs = AllDueSyncs(a_refs:=bDueRefs _
                          , a_names:=bDueNames _
                          , a_sheets:=bDueSheets _
                          , a_comps:=bDueComps _
                          , a_shapes:=bDueShapes)
    lSection = 1
    With Msg.Section(lSection)
        .Text.MonoSpaced = True
        .Text.Text = sDueSyncs
    End With
         
    If bDueRefs Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = "References synchronization:"
            End With
            .Text.Text = "Due because at least one Reference is new or obsolete."
        End With
    End If
        
    If bDueSheets Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = "Worksheet synchronization:"
            End With
            .Text.Text = "Worksheets synchronization is due because at least of one " & _
                         "the properties (mainly the Name or the CodeName) has changed."
        End With
    End If
    
    If bDueNames Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = "Names synchronization:"
            End With
            .Text.Text = "Due because at least one Name is new, obsolete, or has modified properties (e.g. the Scope)."
        End With
    End If
        
    If bDueComps Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = "VBComponent synchronization:"
            End With
            .Text.Text = "Due because at least one " & _
                         "is new, is obsolete, or has a modified code."
        End With
    End If
    
    Bttn1 = "Perform all due synchronizations" & vbLf & "listed above"
    Bttn2 = "Terminate this synchronization"
    Set cllButtons = mMsg.Buttons(Bttn1, Bttn2)
    mMsg.BttnAppRun AppRunArgs, Bttn1 _
                                , ThisWorkbook _
                                , "mSync.AppRunSyncAll"
    mMsg.BttnAppRun AppRunArgs, Bttn2 _
                                , ThisWorkbook _
                                , "mSync.SyncTerminate"
    
    If bDueRefs Or bDueNames Or bDueSheets Or bDueComps Then
        lSection = lSection + 1
        With Msg.Section(lSection).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization"
            .FontColor = rgbBlue
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_SYNC_CHAPTER
        End With
        
        '~~ Display the mode-less dialog for the Names synchronization to run
        mMsg.Dsply dsply_title:=TITLE_SYNC_ALL _
                 , dsply_msg:=Msg _
                 , dsply_Label_spec:="R70" _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=Services.DialogTop & ";" & Services.DialogLeft
        DoEvents
    End If
        
xt: mBasic.EoP ErrSrc(PROC)
    mCompManClient.Events ErrSrc(PROC), True
    Exit Sub

eh:  Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SyncKind(ByVal s_wbk_source As Workbook, _
                     ByVal s_wbk_target As Workbook, _
            Optional ByRef s_teminated As Boolean = False)
' ------------------------------------------------------------------------------
' Synchronizes the next still due kind of synchronization object.
' ------------------------------------------------------------------------------
    Select Case True
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindReference):      mSyncRefs.SyncKind s_wbk_source, s_wbk_target
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindName):           mSyncNames.SyncKind s_wbk_source, s_wbk_target
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindWorksheet):      mSyncSheets.SyncKind s_wbk_source, s_wbk_target
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindShape):          mSyncShapes.SyncKind
        Case DueSyncKindOfObjects.IsQueued(enSyncObjectKindVBComponent):    mSyncComps.SyncKind s_wbk_source, s_wbk_target
    End Select
    
End Sub

Private Function SyncKindString(ByVal en As enSyncKindOfObject) As String
    Select Case en
        Case enSyncObjectKindVBComponent:    SyncKindString = "VBComponent"
        Case enSyncObjectKindName:    SyncKindString = "Name"
        Case enSyncObjectKindReference:     SyncKindString = "Reference"
        Case enSyncObjectKindShape:   SyncKindString = "Shape"
        Case enSyncObjectKindWorksheet:   SyncKindString = "Worksheet"
    End Select
End Function
                             
Public Sub SyncMode()
    
    Dim sMsg    As mMsg.udtMsg
    Dim Bttn1   As String:  Bttn1 = "Synchronize" & vbLf & "summarized"
    Dim Bttn2   As String:  Bttn2 = "Synchronize" & vbLf & "by kind"
    
    With sMsg.Section(1)
        With .Label
            .Text = Replace(Bttn1, vbLf, " ")
            .FontBold = True
        End With
        .Text.Text = "All required synchronizations are listed and can be performed by one single " & _
                     "click - except the synchronization of VB-Components, which can only be " & _
                     "performed one at a time due to technical constraints. This is the fastest " & _
                     "way of synchronizing the VB-Projects of two Workbooks however."
    End With
    With sMsg.Section(2)
        With .Label
            .Text = Replace(Bttn2, vbLf, " ")
            .FontBold = True
        End With
        .Text.Text = "All required synchronizations of one kind (Reference, Worksheet, Name, Shape, " & _
                     "VBComponent) are listed and performed by a single click - except the " & _
                     "synchronization of VB-Components, which can only be performed one at a time " & _
                     "due to technical constraints. This is a medium fast way of synchronizing the " & _
                     "VB-Projects of two Workbooks still with some more control over what is synchronized."
    End With
    Select Case mMsg.Dsply(dsply_title:="Decide on kind of synchronization mode" _
                         , dsply_msg:=sMsg _
                         , dsply_buttons:=mMsg.Buttons(Bttn1, Bttn2))
        Case Bttn1: lSyncMode = SyncSummarized
        Case Bttn2: lSyncMode = SyncByKind
    End Select
    
End Sub

Private Function SyncStepString(ByVal en As enSyncStep) As String
    Select Case en
        Case enSyncStepCollecting:   SyncStepString = "Collecting"
        Case enSyncStepSyncing:      SyncStepString = "Syncing"
    End Select
End Function

Private Sub SyncTerminate()
    MsgBox "Synchronization teminated"
End Sub

Private Sub TargetArchive()
' ------------------------------------------------------------------------------
' Archives the opened Sync-Target-Workbook under the archiving name
' <workbook-base-name>-yy-mm-dd-nn.<extension> in a dedicated folder named
' <workbook-base-name>-yy-mm-dd-nn whereby nn is a number from 01 to 99 to
' distinguish a Sync-Target-Workbook archived more than once during a day.
' ------------------------------------------------------------------------------
    Dim sArchiveFolderRoot      As String
    Dim sArchiveFolderTarget    As String
    Dim l                       As Long
    Dim sArchivedWbkFullName    As String
    Dim sTargetWbkBaseName      As String
    Dim sTargetFullName         As String
    Dim sArchiveSuffix          As String
    Dim sExt                    As String
    
    sTargetFullName = wsService.CurrentServicedWorkbookFullName
    With FSo
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

Private Sub TargetClearExportFiles()
    Dim sFolder         As String
    Dim sTargetFullName As String
    
    sTargetFullName = wsService.CurrentServicedWorkbookFullName
    With FSo
        sFolder = .GetParentFolderName(sTargetFullName) & "\" & wsConfig.FolderExport
        If .FolderExists(sFolder) Then .DeleteFolder sFolder
    End With
    
End Sub

Private Sub TargetClose()
' ------------------------------------------------------------------------------
' Closes an open Sync-Target-Workbook by considering possible changes made.
' ------------------------------------------------------------------------------
    Const PROC = "TargetClose"
    
    Dim wbk                         As Workbook
    Dim Msg                         As mMsg.udtMsg
    Dim MsgTitle                    As String
    Dim BttnCloseWithoutSaving      As String
    Dim BttnCloseBySavingChanges    As String
    
    If mWbk.IsOpen(wsService.CurrentServicedWorkbookFullName, wbk) Then
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
            mCompManClient.Events ErrSrc(PROC), False
            wbk.Close
            mCompManClient.Events ErrSrc(PROC), True
        End If
    End If
End Sub

Private Sub TargetOpen()
    Set wbkSyncTarget = mWbk.GetOpen(wsService.CurrentServicedWorkbookFullName)
End Sub

Private Function TargetOriginFullName(ByVal tofn_wbk As Variant) As String
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

Private Sub TargetWorkingCopyClose()
' ------------------------------------------------------------------------------
' Closes a still open Sync-Target-Workbook's working copy by considering
' possible synchronizations or manual changes made.
' ------------------------------------------------------------------------------
    Const PROC = "TargetWorkingCopyClose"
    
    Dim wbk                         As Workbook
    Dim Msg                         As mMsg.udtMsg
    Dim MsgTitle                    As String
    Dim BttnCloseWithoutSaving      As String
    Dim BttnCloseBySavingChanges    As String
    Dim sTargetWorkingCopyFullName         As String
    
    sTargetWorkingCopyFullName = wsService.SyncTargetFullNameCopy
    If mWbk.IsOpen(sTargetWorkingCopyFullName, wbk) Then
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
            mCompManClient.Events ErrSrc(PROC), False
            wbk.Close
            mCompManClient.Events ErrSrc(PROC), True
        End If
    End If
End Sub

Private Function TargetWorkingCopyDelete() As Boolean
    Dim sTargetWorkingCopyFullName As String
    
    sTargetWorkingCopyFullName = wsService.SyncTargetFullNameCopy
    With FSo
        If .FileExists(sTargetWorkingCopyFullName) Then .DeleteFile sTargetWorkingCopyFullName
    End With

End Function

Private Function TargetWorkingCopyExists() As Boolean
    Dim sTargetWorkingCopyFullName As String
    
    sTargetWorkingCopyFullName = wsService.SyncTargetFullNameCopy
    TargetWorkingCopyExists = FSo.FileExists(sTargetWorkingCopyFullName)

End Function

Public Function TargetWorkingCopyFullName(ByVal tcfn_wbk As Variant) As String
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
    Then TargetWorkingCopyFullName = Replace(str, ".xls", SYNC_TARGET_SUFFIX & ".xls") _
    Else TargetWorkingCopyFullName = str

End Function

Private Function TargetWorkingCopyIsOpened() As Boolean
    TargetWorkingCopyIsOpened = ActiveWorkbook.Name Like "*" & SYNC_TARGET_SUFFIX & ".*"
End Function

Private Sub TargetWorkingCopyOpen(Optional ByVal tco_new As Boolean = True, _
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
    Const PROC = "TargetWorkingCopyOpen"
    
    Dim sTargetWorkingCopy As String
    Dim sTarget     As String
    Dim wbk         As Workbook
    
    sTarget = wsService.CurrentServicedWorkbookFullName
    sTargetWorkingCopy = wsService.SyncTargetFullNameCopy
    mCompManClient.Events ErrSrc(PROC), False
    
    If tco_new Then
        '~~ Delete an already existing working copy
        If FSo.FileExists(sTargetWorkingCopy) Then
            If mWbk.IsOpen(sTargetWorkingCopy, wbk) Then
                wbk.Close False
            End If
            FSo.DeleteFile (sTargetWorkingCopy)
        End If
        ' Save the provided Workbook under the copy name and set it as the current working copy
        mWbk.GetOpen(sTarget).SaveAs FileName:=sTargetWorkingCopy, AccessMode:=xlExclusive
        Set tco_wbk_result = ActiveWorkbook
    Else
        '~~ Continue with existing working copy
        Set tco_wbk_result = mWbk.GetOpen(sTargetWorkingCopy)
        tco_wbk_result.Activate
    End If
    
xt:  mCompManClient.Events ErrSrc(PROC), True
    
End Sub

