Attribute VB_Name = "mSyncSheets"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mSyncSheets: Provides all services for the synchroni-
'                              sation of New, Obsolete, or Property
'                              Changed Worksheets.
' Public services:
' - AllDone                 Returns TRUE when all Worksheets are in sync
' - ClearLinksToSource      1. Retrieves all links in the provided
'                              Workbook and breaks them, i.e. in all used
'                              formulas.
'                           2. Retrieves all Names referring to a range
'                              in a provided Sync-Source-Workbook and
'                              removes it from the ReferredTo property.
' - CollectChangedCodeName  Returns a collection of those Worksheets in
'                           the Sync-Source-Workbook which do not exist in
'                           the Sync-Target-Workbook under the Name but
'                           with a different CodeName
' - CollectChanged          Returns a collection of those Worksheets in
'                           the Sync-Target-Workbook which do not exist in
'                           the Sync-Source-Workbook under a different
'                           Name or CodeName.
' - CollectNew              Returns a collection of those Worksheets in
'                           the Sync-Source-Workbook which do not exist in
'                           the Sync-Target-Workbook under the Name and
'                           CodeName.
' - CollectObsolete         Returns a collection of those Worksheets in
'                           the Sync-Target-Workbook which are regarded
'                           obsolete because they do not exist in the
'                           Sync-Source-Workbook under its Name nor
'                           CodeName.
' - Corresponding           Returns the Worksheet (sync_wsh_result) in the Workbook (sync_wbk) which
'                           corresponds to the Worksheet (sync_wsh) whereby corresponding is defined by
'                           either an equal Name or an equal CodeName or both equal.
' - RunSyncCodeName         Called via Application.Run by CommonButton: Changes the CodeName of
'                           the provided Sheet in the provided Sync-Target-Workbook to the
'                           CodeName of the sheet in the provided Sync-Source-Workbook.
' - AppRunChanged               Called via Application.Run by CommonButton: Changes the name of the
'                           provided Sheet in the provided Sync-Target-Workbook to the Name of
'                           the sheet in the provided Sync-Source-Workbook.
' - AppAppRunNew               Called via Application.Run by CommonButton: Copies a provided sheet
'                           from a provided Sync-Source-Workbook to a provided Sync-Target-
'                           Workbook.
' - AppAppRunObsolete       Called via Application.Run by CommonButton: Removes a provided Sheet
'                           from a provided Workbook
' - Sync                    Called by mSync.RunSync when there are (still) any outstanding Sheet
'                           synchronizations to be done. Displays them in a mode-less dialog for
'                           being confirmed one by one.
' - SyncOrder               Called by mSync.RunSync: Syncronizes the sheet's order in the Sync-
'                           Target-Workbook to appear in the same order as in the Sync-Source-
'                           Workbook.
'
' W. Rauschenberger Berlin, Dec 2022
' ------------------------------------------------------------------------
Private Const TITLE_SYNC_SHEETS       As String = "VB-Project Synchronization: Worksheets"

Private dctKnownInSync          As Dictionary
Private dctKnownNew             As Dictionary
Private dctKnownObsolete        As Dictionary
Private dctKnownChanged         As Dictionary
Private dctKnownOwnedByPrjct    As Dictionary

Public Enum enCorrespondingSheetsQuality
    enOrNameCodeName
    enAndNameCodeName
End Enum

Public Property Get KnownChanged(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownChanged Is Nothing _
    Then KnownChanged = dctKnownChanged.Exists(nk_id)
End Property

Public Property Let KnownChanged(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownChanged, nk_id
End Property

Private Property Get KnownInSync(Optional ByVal is_id As String) As Boolean
    If Not dctKnownInSync Is Nothing _
    Then KnownInSync = dctKnownInSync.Exists(is_id)
End Property

Private Property Let KnownInSync(Optional ByVal is_id As String, _
                                                 ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownInSync, is_id
End Property

Public Property Get KnownNew(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownNew Is Nothing _
    Then KnownNew = dctKnownNew.Exists(nk_id)
End Property

Public Property Let KnownNew(Optional ByVal nk_id As String, _
                                          ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownNew, nk_id
End Property

Public Property Get KnownObsolete(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownObsolete Is Nothing _
    Then KnownObsolete = dctKnownObsolete.Exists(nk_id)
End Property

Public Property Let KnownObsolete(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownObsolete, nk_id
End Property

Public Property Get KnownOwnedByPrjct(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownOwnedByPrjct Is Nothing _
    Then KnownOwnedByPrjct = dctKnownOwnedByPrjct.Exists(nk_id)
End Property

Public Property Let KnownOwnedByPrjct(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownOwnedByPrjct, nk_id
End Property

Public Sub AppRunSyncAll()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all Sheets by
' removing obsolete, adding new, and changing the Name/CodeName.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunSyncAll"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If dctKnownChanged.Count <> 0 Then mSyncSheets.AppRunChanged                ' Synchronize Worksheets the Name or CodeName changed
    If dctKnownNew.Count <> 0 Then mSyncSheets.AppRunNew                        ' Synchronize new Worksheets
    If dctKnownObsolete.Count <> 0 Then mSyncSheets.AppRunObsolete              ' Synchronize obsolete Worksheets
    If dctKnownOwnedByPrjct.Count <> 0 Then mSyncSheets.AppRunOwnedByPrjctIds   ' Synchronize sheets regarded "owned-by-the-VB-Project"

xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        Services.MessageUnload TITLE_SYNC_SHEETS
        mSync.RunSync
    End If
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AppRunChanged()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Changes the Name property of the
' Name object in the Sync-Target-Workbook which refers to the Name's (rc_nme)
' Range to the Name property of the Name object (rc_nme).
' ------------------------------------------------------------------------------
    Const PROC = "AppRunChanged"
    
    On Error GoTo eh
    Dim sOldName        As String
    Dim sOldCodeName    As String
    Dim wbkSource       As Workbook
    Dim wbkTarget       As Workbook
    Dim wshSource       As Worksheet
    Dim wshTarget       As Worksheet
    Dim bName           As Boolean
    Dim bCodeName       As Boolean
    Dim vSource         As Variant
    Dim vTarget         As Variant
    Dim i               As Long
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    vSource = Split(AppRunChangedIdsSource, ",")
    vTarget = Split(AppRunChangedIdsTarget, ",")
    mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                 , p_sync_step:=enSyncStepSyncing _
                 , p_sync_action:=enSyncActionChanged _
                 , p_count:=0
    
    mSync.AppRunInit
    
    For i = LBound(vSource) To UBound(vSource)
        GetSheet wbkSource, vSource(i), wshSource
        GetSheet wbkTarget, vTarget(i), wshTarget
        PropertiesDiffer pd_wsh_source:=wshSource _
                       , pd_wsh_target:=wshTarget _
                       , pd_name_changed:=bName _
                       , pd_codename_changed:=bCodeName
        If bName Then
            sOldName = wshTarget.Name
            wshTarget.Name = wshSource.Name
            wsSyncLog.Done "change", "Worksheet", SyncId(wshTarget), "changed", "Name changed from " & sOldName & " to " & wshTarget.Name
        ElseIf bCodeName Then
            sOldCodeName = wshTarget.CodeName
            mWsh.ChangeCodeName wbkTarget, sOldCodeName, wshSource.CodeName
            wsSyncLog.Done "change", "Worksheet", SyncId(wshTarget), "changed", "CodeName changed from " & sOldCodeName & " to " & wshTarget.CodeName
        End If
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                     , p_sync_step:=enSyncStepSyncing _
                     , p_sync_action:=enSyncActionChanged _
                     , p_count:=i + 1
    Next i
    
    dctKnownChanged.RemoveAll ' indicates done
    mSync.AppRunTerminate
    
xt: Services.MessageUnload TITLE_SYNC_SHEETS
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppRunChangedIdsSource() As String
    AppRunChangedIdsSource = DueSyncIdsByAction(enSyncObjectKindWorksheet, enSyncActionChanged, "to")
End Function

Private Function AppRunChangedIdsTarget() As String
    AppRunChangedIdsTarget = DueSyncIdsByAction(enSyncObjectKindWorksheet, enSyncActionChanged, "from")
End Function

Private Function AppRunOwnedByPrjctIds() As String
    AppRunOwnedByPrjctIds = DueSyncIdsByAction(enSyncObjectKindWorksheet, enSyncActionOwnedByPrjct)
End Function

Private Sub AppRunNew()
' ------------------------------------------------------------------------------
' Synchronize new and changed Worksheets and synchronize the concerned Names'
' properties.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunNew"
    
    On Error GoTo eh
    Dim i           As Long
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    Dim va          As Variant
    Dim dctNames    As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    mSyncNames.AllNames wbkTarget, dctNames
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    va = Split(AppRunNewIds(enSyncObjectKindWorksheet), ",")
    mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                 , p_sync_step:=enSyncStepSyncing _
                 , p_sync_action:=enSyncActionAddNew _
                 , p_count:=0
    
    mSync.AppRunInit
    
    For i = LBound(va) To UBound(va)
        GetSheet wbkSource, va(i), wshSource
        CloneSheetToTargetWorkbook wshSource, wbkTarget, wshTarget
        ClearLinksToSource wshSource, wshTarget
        RemoveDuplicatesCreatedBySheetClone r_in_dct:=dctNames _
                                          , r_wbk_source:=wbkSource _
                                          , r_wbk_target:=wbkTarget
        
        wsSyncLog.Done "new", "Worksheet", va(i), "added", "New! Added/cloned to Sync-Target-Workbook (working copy)"
        mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                     , p_sync_step:=enSyncStepSyncing _
                     , p_sync_action:=enSyncActionAddNew _
                     , p_count:=i + 1
    Next i
    
    dctKnownNew.RemoveAll ' indicates that all new Names had been added
    mSync.AppRunTerminate

xt: Services.MessageUnload TITLE_SYNC_SHEETS
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AppRunObsolete()
' ------------------------------------------------------------------------------
' Synchronize obsolete Worksheets
' ------------------------------------------------------------------------------
    Const PROC      As String = "AppRunObsolete"
    
    On Error GoTo eh
    Dim i           As Long
    Dim wsh         As Worksheet
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim va          As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    Application.DisplayAlerts = False
    va = Split(AppRunObsoleteIds(enSyncObjectKindWorksheet), ",")
    mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                 , p_sync_step:=enSyncStepSyncing _
                 , p_sync_action:=enSyncActionRemoveObsolete _
                 , p_count:=0
    mSync.AppRunInit
    
    For i = LBound(va) To UBound(va)
        GetSheet wbkTarget, va(i), wsh
        If Not wsh Is Nothing Then
            Services.ServicedItem = wsh
            Delete wsh
            wsSyncLog.Done "obsolete", "Worksheet", va(i), "removed", "Obsolete! Removed from Sync-Target-Workbook (working copy)"
        Else
            Debug.Print "The Worksheet '" & va(i) & "' no longer exists!"
        End If
        mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                     , p_sync_step:=enSyncStepSyncing _
                     , p_sync_action:=enSyncActionRemoveObsolete _
                     , p_count:=i + 1
    Next i
    
    dctKnownObsolete.RemoveAll ' indicates that all removals had been done
    mSync.AppRunTerminate
    
xt: Application.DisplayAlerts = True
    Services.MessageUnload TITLE_SYNC_SHEETS
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

Private Sub ClearLinksToSource(ByVal cls_wsh_source As Worksheet, _
                               ByVal cls_wsh_target As Worksheet)
' -----------------------------------------------------------------------------
' Required when a new Worksheet is copied/cloned from the Sync-Source-Workbook
' to the Sync-Target-Workbook.
' 1. Retrieves all links in the provided Workbook (sync_wbk_target) and breaks
'    them, i.e. in all used formulas
' 2. Retrieves all Names referring to a range in the Sync-Source-Workbook
'    (sync_wbk_source) and removes the refer-back from the ReferedTo property
' 3. Removes all Names referring back to the Sync-Source-Workbook to another
'    but the source sheet (cls_wsh_source)
' 4. Clear back-links to the Sync-Source-Workbook (sync_wbk_source) in Shape's
'    OnAction property
' -----------------------------------------------------------------------------
    Const PROC = "ClearLinksToSource"
    
    On Error GoTo eh
    Dim aLinks      As Variant
    Dim v           As Variant
    Dim nme         As Name
    Dim wsh         As Worksheet
    Dim shp         As Shape
    Dim sRefersTo   As String
    Dim sOnAction   As String
    
    mBasic.BoP ErrSrc(PROC)
    aLinks = mSync.TargetWorkingCopy.LinkSources(xlExcelLinks)
    If mBasic.ArrayIsAllocated(aLinks) Then
        For Each v In aLinks
            Services.ServicedItem = cls_wsh_target
            mSync.TargetWorkingCopy.BreakLink v, xlLinkTypeExcelLinks
            DoEvents
            LogServiced.Entry "obsolete", "Worksheet", SyncId(cls_wsh_target), "removed", "(Back)Link to '" & Split(v, "\")(UBound(Split(v, "\"))) & "' cleared"
        Next v
    End If
        
    '~~ Clear Link to source Workbook in Range-Names
    For Each nme In mSync.TargetWorkingCopy.Names
        If InStr(nme.RefersTo, mSync.source.Name) <> 0 Then
            Services.ServicedItem = cls_wsh_target
            sRefersTo = nme.RefersTo
            If InStr(nme.RefersTo, "[" & mSync.source.Name & "]") <> 0 Then
                If InStr(nme.RefersTo, "]" & cls_wsh_source.Name & "!") <> 0 Then
                    '~~ Names which had just been copied from the Sync-Source-Workbook allong with the cloning of the source sheet
                    '~~ are nor handled by this sheet synchronization but with the subsequent Names synchronization - in order
                    '~~ to enable full control of them.
                    nme.Delete
                Else
                    nme.RefersTo = Replace(nme.RefersTo, "[" & mSync.source.Name & "]", vbNullString)
                    LogServiced.Entry "invalid", "Worksheet", Services.ServicedItemName, "changed", "Back-Reference '" & sRefersTo & "' changed to '" & nme.RefersTo & "'"
                End If
            End If
        End If
    Next nme
    
    '~~ Clear Link to source in any shapes OnAction property
    For Each wsh In mSync.TargetWorkingCopy.Sheets
        Services.ServicedItem = wsh
        For Each shp In wsh.Shapes
            On Error Resume Next
            sOnAction = shp.OnAction
            If Err.Number = 0 Then ' shape has an OnAction property
                If InStr(sOnAction, mSync.source.Name) <> 0 Then
                    Services.ServicedItem = wsh
                    shp.OnAction = Replace(sOnAction, mSync.source.Name, mSync.TargetWorkingCopy.Name)
                    LogServiced.Entry "invalid", "Worksheet", Services.ServicedItemName, "changed", "Back-Link to source '" & sOnAction & "' changed to '" & shp.OnAction & "'"
                End If
            End If
        Next shp
    Next wsh

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CloneSheetToTargetWorkbook(ByVal sync_wsh_source As Worksheet, _
                                       ByVal SYNC_TEST_WBK_TARGET As Workbook, _
                                       ByRef sync_wsh_target As Worksheet)
' -----------------------------------------------------------------------------
' Copies the sheet (sync_wsh_source) from the Sync-Source-Workbook to the
' Sync-Target-Workbook and returns the target sheet as object.
' -----------------------------------------------------------------------------
    Const PROC = "CloneSheetToTargetWorkbook"
    
    On Error GoTo eh
    Dim wshTarget   As Worksheet
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    
    If Exists(wbkTarget, sync_wsh_source.Name, , wshTarget) Then
        Application.DisplayAlerts = False
        mWsh.Delete wshTarget
        Application.DisplayAlerts = True
    End If
    sync_wsh_source.Copy After:=SYNC_TEST_WBK_TARGET.Sheets(SYNC_TEST_WBK_TARGET.Worksheets.Count)
    Set sync_wsh_target = SYNC_TEST_WBK_TARGET.Sheets(SYNC_TEST_WBK_TARGET.Worksheets.Count)
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Collect(ByVal c_wbk_source As Workbook, _
                   ByVal c_wbk_target As Workbook)
' ------------------------------------------------------------------------------
' Returns a collection of all those sheet controls which exist in the source
' but not in the target Workbook, for each Shape the Application.Run arguments:
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim wshSource       As Worksheet
    Dim wshTarget       As Worksheet
    Dim sId             As String
    Dim bNameChange     As Boolean
    Dim bCodeNameChange As Boolean
    Dim sCorresponding  As String
    
    mBasic.EoP ErrSrc(PROC)
    If mSyncSheets.AllDone(c_wbk_source, c_wbk_target) Then GoTo xt
    If lSyncMode = SyncByKind Then mSync.InitDueSyncs
    
    '~~ Collect changed
    If DueCollect("Changed") Then
        mSync.Progress p_kind:=enSyncObjectKindWorksheet _
             , p_sync_step:=enSyncStepCollecting _
             , p_sync_action:=enSyncActionChanged _
             , p_count:=0

        For Each wshSource In mSync.source.Sheets
            sId = SyncId(wshSource)
            If Not KnownSource(sId) Then
                '~~ Note: A sheet already known new may be an 'owned-by-project' sheet which is re-cloned by default
                If mSyncSheets.Corresponding(c_wsh:=wshSource _
                                           , c_wbk:=c_wbk_target _
                                           , c_quality:=enOrNameCodeName _
                                           , c_name_change:=bNameChange _
                                           , c_code_name_change:=bCodeNameChange _
                                           , c_wsh_result:=wshTarget) = "1:1" Then
                    If bNameChange Or bCodeNameChange Then
                        If Not IsOwnedByPrjct(wshSource) Then
                            CollectChanged wshSource, wshTarget
                        End If
                    End If
                End If
            End If
            mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                 , p_sync_step:=enSyncStepCollecting _
                 , p_sync_action:=enSyncActionChanged _
                 , p_count:=dctKnownChanged.Count
        Next wshSource
    End If

    '~~ Collect new
    If DueCollect("New") Then
        mSync.Progress p_kind:=enSyncObjectKindWorksheet _
             , p_sync_step:=enSyncStepCollecting _
             , p_sync_action:=enSyncActionAddNew _
             , p_count:=0
        
        For Each wshSource In c_wbk_source.Worksheets
            sId = SyncId(wshSource)
            If Not KnownInSync(sId) Then
                Select Case mSyncSheets.Corresponding(c_wsh:=wshSource _
                                                    , c_wbk:=c_wbk_target _
                                                    , c_quality:=enOrNameCodeName)
                    Case "1:0"
                        ' Source sheet without a corresponding target sheet
                        If Not KnownNew(sId) Then
                            mSync.DueSyncLet , enSyncObjectKindWorksheet, enSyncActionAddNew, , sId
                        End If
                    Case "1:1"
                        If IsOwnedByPrjct(wshSource) Then
                            ' Source sheet considered 'owned-by-project' with a corresponding target sheet
                            If Not KnownOwnedByPrjct(sId) Then
                                mSync.DueSyncLet , enSyncObjectKindWorksheet, enSyncActionOwnedByPrjct, , sId
                            End If
                        End If
                End Select
            End If
            mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                 , p_sync_step:=enSyncStepCollecting _
                 , p_sync_action:=enSyncActionAddNew _
                 , p_count:=dctKnownNew.Count
        Next wshSource
    End If
        
    '~~ Collect obsolete
    If DueCollect("Obsolete") Then
            mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                 , p_sync_step:=enSyncStepCollecting _
                 , p_sync_action:=enSyncActionRemoveObsolete _
                 , p_count:=0
        
        For Each wshTarget In c_wbk_target.Worksheets
            sId = SyncId(wshTarget)
            If Not KnownInSync(sId) Then
                sCorresponding = mSyncSheets.Corresponding(c_wsh:=wshTarget _
                                                         , c_wbk:=c_wbk_source _
                                                         , c_quality:=enOrNameCodeName _
                                                         , c_wsh_result:=wshSource)
                Select Case sCorresponding
                    Case "1:0"
                        ' Target sheet without a corresponding source sheet
                        mSync.DueSyncLet , enSyncObjectKindWorksheet, enSyncActionRemoveObsolete, , sId
                End Select
            End If
            mSync.Progress p_kind:=enSyncObjectKindWorksheet _
                 , p_sync_step:=enSyncStepCollecting _
                 , p_sync_action:=enSyncActionRemoveObsolete _
                 , p_count:=dctKnownObsolete.Count
        Next wshTarget
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Collected(ByVal c_action As enSyncAction) As Long
    Select Case True
        Case mSync.SyncActionIsChange(c_action):    Collected = dctKnownChanged.Count
        Case c_action = enSyncActionRemoveObsolete: Collected = dctKnownObsolete.Count
        Case c_action = enSyncActionAddNew:         Collected = dctKnownNew.Count
        Case c_action = enSyncActionOwnedByPrjct:   Collected = dctKnownOwnedByPrjct.Count
    End Select
End Function

Private Sub CollectChanged(ByVal cc_wsh_source As Worksheet, _
                           ByVal cc_wsh_target As Worksheet)
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    Const PROC = "CollectChanged"
    
    On Error GoTo eh
    Dim bName       As Boolean
    Dim bCodeName   As Boolean
    Dim enAction    As enSyncAction
    Dim sIdSource   As String:      sIdSource = SyncId(cc_wsh_source)
    Dim sIdTarget   As String:      sIdTarget = SyncId(cc_wsh_target)
    
    PropertiesDiffer pd_wsh_source:=cc_wsh_source _
                   , pd_wsh_target:=cc_wsh_target _
                   , pd_name_changed:=bName _
                   , pd_codename_changed:=bCodeName
    If Not bName And Not bCodeName Then GoTo xt
    
    Select Case True
        Case bName:     enAction = enSyncActionChangeName
        Case bCodeName: enAction = enSyncActionChangeCodeName
    End Select
    
    mSync.DueSyncLet , enSyncObjectKindWorksheet, enAction, "from", sIdTarget
    mSync.DueSyncLet , enSyncObjectKindWorksheet, enAction, "to", sIdSource
    KnownChanged(sIdSource) = True
    KnownChanged(sIdTarget) = True
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function CollectInSync(ByVal c_wbk_target As Workbook, _
                               ByVal c_wbk_source As Workbook) As Long
' ----------------------------------------------------------------------------
' Collects all the already in sync Worksheets (in dctKnownInSync) and
' returns their number.
' ----------------------------------------------------------------------------
    Const PROC = "CollectInSync"
    
    On Error GoTo eh
    Dim wsh As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    Set dctKnownInSync = Nothing
    Set dctKnownInSync = New Dictionary
    For Each wsh In c_wbk_source.Worksheets
        If mSyncSheets.Corresponding(c_wsh:=wsh _
                                   , c_wbk:=c_wbk_target _
                                   , c_quality:=enAndNameCodeName) = "1:1" Then
            mSync.CollectKnown dctKnownInSync, SyncId(wsh)
        End If
    Next wsh

xt: CollectInSync = dctKnownInSync.Count
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Corresponding(ByVal c_wsh As Worksheet, _
                              ByVal c_wbk As Workbook, _
                              ByVal c_quality As enCorrespondingSheetsQuality, _
                     Optional ByRef c_name_change As Boolean, _
                     Optional ByRef c_code_name_change As Boolean, _
                     Optional ByRef c_wsh_result As Worksheet) As String
' ------------------------------------------------------------------------------
' Returns
' - the Worksheet (c_wsh_result) in the Workbook (c_wbk) which corresponds
'   to the Worksheet (c_wsh)
' - A string indicating the kind of correspondence
'   - 1:1 The sheet (c_wsh) exists in the Workbook (c_wbk),
'         when c_quality = enStrong with Name and CodeName equal,
'         when c_quality = enWeak (default) with both or one equal
'   - 1:0 The sheet (c_wsh) does not exists in the Workbook (c_wbk)
'
' Corresponding is defined:
' - When c_quality = enWeak, either the Names or the CodeNames or both equal
' - When c_quality = enStrong, the Name and the CodeName are equal.
'
' ------------------------------------------------------------------------------
    Const PROC = "Corresponding"
    
    On Error GoTo eh
    Dim wsh             As Worksheet
    Dim bSameName       As Boolean
    Dim bSameCodeName   As Boolean
    Dim sId             As String
    
    mBasic.BoP ErrSrc(PROC)
    Set c_wsh_result = Nothing
    For Each wsh In c_wbk.Worksheets
        sId = SyncId(wsh)
        bSameName = wsh.Name = c_wsh.Name
        bSameCodeName = wsh.CodeName = c_wsh.CodeName
        If bSameName Or bSameCodeName Then
            If c_quality = enAndNameCodeName And bSameName And bSameCodeName _
            Or c_quality = enOrNameCodeName And (bSameName Or bSameCodeName) Then
                Corresponding = "1:1"
                Set c_wsh_result = wsh
                c_name_change = Not bSameName
                c_code_name_change = Not bSameCodeName
                Exit For
            End If
        End If
    Next wsh
    If c_wsh_result Is Nothing Then
        Corresponding = "1:0"
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub Delete(ByVal d_wsh As Worksheet, _
          Optional ByRef d_protected As Boolean)
' ------------------------------------------------------------------------------
' Provides a 'clean deletion' of a Worksheet by removing all relevant Name
' objects beforehand in order to prevent invalid Name objects reulting from the
' deletion
' ------------------------------------------------------------------------------
    Const PROC = "Delete"
    
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    d_protected = d_wsh.ProtectContents
    If d_protected Then d_wsh.Unprotect
    mSyncNames.RemoveAllOfSheet d_wsh
    Application.DisplayAlerts = False
    On Error Resume Next
    d_wsh.Delete
    Application.DisplayAlerts = True
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function AllDone(ByVal d_wbk_source As Workbook, _
                        ByVal d_wbk_target As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the number of in-sync Names plus any sipped/2 are equal
' the Names in the Sync-Source- and Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "AllDone"
    
    On Error GoTo eh
    Dim lInSync    As Long
    Dim lTarget    As Long
    Dim lSource    As Long
    
    mBasic.BoP ErrSrc(PROC)
    lInSync = CollectInSync(d_wbk_source, d_wbk_target)
    lTarget = d_wbk_target.Worksheets.Count
    lSource = d_wbk_source.Worksheets.Count
    
    If lTarget = lSource _
    And lSource = lInSync Then
        AllDone = True
        mMsg.MsgInstance TITLE_SYNC_SHEETS, True
        mSyncSheets.SyncOrder d_wbk_source, d_wbk_target
        wsSyncLog.SummaryDone("Worksheet") = True
        mSync.DueSyncKindOfObjects.DeQueue , enSyncObjectKindWorksheet
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function DueCollect(ByVal d_c As String) As Boolean
    Select Case d_c
        Case "New":         DueCollect = dctKnownNew Is Nothing
                            If DueCollect Then
                                Set dctKnownNew = New Dictionary
                                Set dctKnownOwnedByPrjct = New Dictionary
                            End If
        Case "Changed":     DueCollect = dctKnownChanged Is Nothing
                            If DueCollect Then Set dctKnownChanged = New Dictionary
        Case "Obsolete":    DueCollect = dctKnownObsolete Is Nothing
                            If DueCollect Then Set dctKnownObsolete = New Dictionary
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncSheets." & s
End Function

Private Function Exists(ByVal se_wbk As Workbook, _
               Optional ByRef se_wsh_name As String = vbNullString, _
               Optional ByRef se_wsh_code_name As String = vbNullString, _
               Optional ByRef se_wsh_result As Worksheet) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the sheet identified by the provided name (se_wsh_name)
' and/or the provided CodeName (se_wsh_code_name) exists in the Workbook
' (se_wbk) under the provided Name/CodeName. The non-provided is returned when
' the Sheet exists.
' -----------------------------------------------------------------------------
    Const PROC = "Exists"
                             
    On Error GoTo eh
    Dim wsh As Worksheet
    
    If se_wsh_name = vbNullString And se_wsh_code_name = vbNullString _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Neither a Sheet's Name nor a Sheet's CodeName is provided!"
    
    For Each wsh In se_wbk.Worksheets
        Select Case True
            Case se_wsh_name <> vbNullString And se_wsh_code_name = vbNullString
                If wsh.Name = se_wsh_name Then
                    se_wsh_code_name = wsh.CodeName
                    Set se_wsh_result = wsh
                    Exists = True
                    GoTo xt
                End If
            Case se_wsh_name = vbNullString And se_wsh_code_name <> vbNullString
                If wsh.CodeName = se_wsh_code_name Then
                    se_wsh_name = wsh.Name
                    Set se_wsh_result = wsh
                    Exists = True
                    GoTo xt
                End If
            Case se_wsh_name <> vbNullString And se_wsh_code_name <> vbNullString
                If wsh.Name = se_wsh_name And wsh.CodeName = se_wsh_code_name Then
                    Set se_wsh_result = wsh
                    Exists = True
                    GoTo xt
                End If
        End Select
    Next wsh
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function GetSheet(ByVal gs_wbk As Workbook, _
                          ByVal gs_id As String, _
                 Optional ByRef gs_wsh As Worksheet) As Worksheet
    Dim wsh As Worksheet
    
    For Each wsh In gs_wbk.Worksheets
        If SyncId(wsh) = gs_id Then
            Set GetSheet = wsh
            Set gs_wsh = wsh
            Exit For
        End If
    Next wsh

End Function

Private Function HasUnlockedRange(ByVal hur_wsh As Worksheet) As Boolean
' -----------------------------------------------------------------------------
' Returns TRUE when the sheet (hur_wsg) has any unlocked range - which means
' that the user is able to change a value.
' -----------------------------------------------------------------------------
    Dim Rng As Range
    
    For Each Rng In hur_wsh.UsedRange.Cells
        If Rng.Locked = False Then
            HasUnlockedRange = True
            Exit Function
        End If
    Next Rng
    
End Function

Public Sub Initialize()
    Set dctKnownInSync = Nothing
    Set dctKnownChanged = Nothing
    Set dctKnownNew = Nothing
    Set dctKnownObsolete = Nothing
    Set dctKnownOwnedByPrjct = Nothing
End Sub

Private Function IsOwnedByPrjct(ByVal obp_wsh As Worksheet) As Boolean
    IsOwnedByPrjct = obp_wsh.ProtectContents And Not HasUnlockedRange(obp_wsh) Or Not obp_wsh.Visible
End Function

Private Function KnownSource(ByVal k_id As String) As Boolean
    KnownSource = KnownInSync(k_id) _
               Or KnownNew(k_id) _
               Or KnownChanged(k_id) _
               Or KnownOwnedByPrjct(k_id)
End Function

Private Sub PropertiesDiffer(ByVal pd_wsh_source As Worksheet, _
                             ByVal pd_wsh_target As Worksheet, _
                    Optional ByRef pd_name_changed As Boolean, _
                    Optional ByRef pd_codename_changed As Boolean, _
                    Optional ByRef pd_none_changed As Boolean)
' ----------------------------------------------------------------------------
' Returns TRUE for those Properties which differ (pd_name_changed,
' pd_codename_changed) between the provided Worksheet objects
' (pd_wsh_source, pd_wsh_target).
' ----------------------------------------------------------------------------
    pd_name_changed = pd_wsh_source.Name <> pd_wsh_target.Name
    pd_codename_changed = pd_wsh_source.CodeName <> pd_wsh_target.CodeName
    pd_none_changed = Not pd_name_changed And Not pd_codename_changed
End Sub

Private Function SyncId(ByVal wsh As Worksheet) As String
' -----------------------------------------------------------------------------
' Returns a unique id for the provided Worksheet (wsh) analogous to the VBE.
' -----------------------------------------------------------------------------
    SyncId = "CodeName=" & wsh.CodeName & " " & "Name=" & wsh.Name
End Function

Public Sub SyncKind(ByVal s_wbk_source As Workbook, _
                    ByVal s_wbk_target As Workbook)
' ----------------------------------------------------------------------------
' Displays a dialog to perform all required action to synchronize the Work-
' sheets of the Sync-Target-Workbook with thos in the Sync-source-Workbook.
' ----------------------------------------------------------------------------
    Const PROC          As String = "SyncKind"
    
    On Error GoTo eh
    Dim fSync           As fMsg
    Dim Msg             As mMsg.udtMsg
    Dim cllButtons      As New Collection
    Dim AppRunArgs      As New Dictionary
    Dim Bttn1           As String
    Dim sDueSyncs       As String
    Dim lSection        As Long
    Dim bDueObsolete    As Boolean
    Dim bDueNew         As Boolean
    Dim bDueChanged     As Boolean
    Dim bDueObP         As Boolean
    
    mCompManClient.Events ErrSrc(PROC), False
    mBasic.BoP ErrSrc(PROC)
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_SHEETS)
    mSync.MonitorStep "Synchronizing Worksheets"
    mSyncSheets.Collect s_wbk_source, s_wbk_target
    Bttn1 = "Perform all Worksheets synchronization actions" & vbLf & "listed above"
    
    sDueSyncs = mSync.DueSyncs(enSyncObjectKindWorksheet)
    With Msg.Section(1)
        .Text.MonoSpaced = True
        .Text.Text = sDueSyncs
    End With
    bDueObsolete = InStr(sDueSyncs, SyncActionString(enSyncActionRemoveObsolete)) <> 0
    bDueNew = InStr(sDueSyncs, SyncActionString(enSyncActionAddNew)) <> 0
    bDueChanged = InStr(sDueSyncs, SyncActionString(enSyncActionChangeName)) <> 0 _
               Or InStr(sDueSyncs, SyncActionString(enSyncActionChangeCodeName)) <> 0
    bDueObP = InStr(sDueSyncs, SyncActionString(enSyncActionOwnedByPrjct)) <> 0
    
    lSection = 1
    If bDueObsolete Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_REMOVE_OBSOLETE & ":"
            End With
            .Text.Text = "Considered because it has no corresponding source sheet, or the corresponding source " & _
                         "sheet is considered an 'owned-by-project' sheet which will replace it."
        End With
    End If
    If bDueNew Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_ADD_NEW & ":"
            End With
            .Text.Text = "Considered new because no corresponding sheet exists in the Sync-Target-Workbook, " & _
                         "or because the sheet is considered 'owned-by-project', i.e. the sheet is protected " & _
                         "and has no unlocked (input) cells and thus is about to replace the correspnding " & _
                         "target sheet which is removed."
        End With
    End If
    If bDueChanged Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_CHANGE_NAME & ", " & SYNC_ACTION_CHANGE_CODENAME & ":"
            End With
            .Text.Text = "Either the Name or the CodeName changed."
        End With
    End If
    
    If bDueObP Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_OWNED_BY_PROJECT & ":"
            End With
            .Text.Text = "The target sheet - if one exists - is removed and the source sheet is re-cloned." & vbLf & _
                         "Note: A Worksheet which is protected and does not have any unlocked, i.e. input, " & _
                         "cells is regarded 'owned by the VB-Project and fully syncronized assuming that " & _
                         "there might have been changes."
        End With
    End If
    
    mMsg.BttnAppRun AppRunArgs, Bttn1 _
                                , ThisWorkbook _
                                , "mSyncSheets.AppRunSyncAll"
    Set cllButtons = mMsg.Buttons(Bttn1)
    
    If bDueObsolete Or bDueNew Or bDueChanged Or bDueObP Then
        With Msg.Section(8).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization, section Name Synchronization:"
            .FontColor = rgbBlue
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_SYNC_CHAPTER_NAMES
        End With
        
        '~~ Display the mode-less dialog for the Names synchronization to run
        mMsg.Dsply dsply_title:=TITLE_SYNC_SHEETS _
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

Private Sub SyncOrder(ByVal so_wbk_source As Workbook, _
                      ByVal so_wbk_target As Workbook)
' ----------------------------------------------------------------------------
' Syncronizes the sheet's order in the Sync-Target-Workbook (sync_wbk_target)
' to appear in the same order as in the Sync-Source-Workbook (sync_wbk_source).
' Precondition: All sheet synchronizations had been done.
' ----------------------------------------------------------------------------
    Const PROC = "SyncOrder"
    
    On Error GoTo eh
    Dim i           As Long
    Dim lSheets     As Long
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    Dim sSourceName As String
    
    Application.ScreenUpdating = False
    lSheets = so_wbk_target.Worksheets.Count
    With so_wbk_source
        For i = 1 To .Worksheets.Count
            Set wshSource = .Worksheets(i)
            sSourceName = wshSource.Name
            With so_wbk_target
                For Each wshTarget In .Worksheets
                    If wshTarget.Name = sSourceName Then
                        wshTarget.Move After:=.Worksheets(lSheets)
                        Exit For
                    End If
                Next wshTarget
            End With
            Services.ServicedItem = wshTarget
        Next i
    End With
    Application.ScreenUpdating = True
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AppRunOwnedByPrjct()
' ------------------------------------------------------------------------------
' Synchronize "owned-by-the-VB-Project Worksheets by removing and re-cloning
' them.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunOwnedByPrjct"
    
    On Error GoTo eh
    Dim bProtected      As Boolean
    Dim dctNames        As Dictionary
    Dim i               As Long
    Dim va              As Variant
    Dim wbkSource       As Workbook
    Dim wbkTarget       As Workbook
    Dim wshSource       As Worksheet
    Dim wshTarget       As Worksheet
    Dim sCorresponding  As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    mSyncNames.AllNames wbkTarget, dctNames
    va = Split(AppRunOwnedByPrjctIds(), ",")
    mSync.Progress enSyncObjectKindWorksheet, enSyncStepSyncing, enSyncActionOwnedByPrjctObsolete, 0
    
    For i = LBound(va) To UBound(va)
        GetSheet wbkSource, va(i), wshSource
        '~~ Remove the existing if any
        sCorresponding = mSyncSheets.Corresponding(c_wsh:=wshSource _
                                                 , c_wbk:=wbkTarget _
                                                 , c_quality:=enOrNameCodeName _
                                                 , c_wsh_result:=wshTarget)
        If sCorresponding = "1:1" Then
            Delete wshTarget, bProtected
        End If
        ' Re-Clone it
        CloneSheetToTargetWorkbook wshSource, wbkTarget, wshTarget
        ClearLinksToSource wshSource, wshTarget
        RemoveDuplicatesCreatedBySheetClone r_in_dct:=dctNames _
                                          , r_wbk_source:=wbkSource _
                                          , r_wbk_target:=wbkTarget
        If bProtected Then wshTarget.Protect
        '~~ Synchronize RefersTo and the Scope of all Names which refer to a range in
        '~~ the cloned/added Worksheet with those of the corresponding Name in the Sync-Source-Workbook
        wsSyncLog.Done "new", "Worksheet", SyncId(wshSource), "added", "re-cloned, considered 'owned-by-project'"
        mSync.Progress enSyncObjectKindWorksheet, enSyncStepSyncing, enSyncActionOwnedByPrjctObsolete, i + 1
    Next i
             
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

