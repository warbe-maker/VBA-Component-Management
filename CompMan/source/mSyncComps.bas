Attribute VB_Name = "mSyncComps"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSyncComps: Provides the services to synchronize
'                             Sync-Target-Workbook VB-Components with
'                             Sync-Source-Workbook VB-Components.
' Public services:
' - AllDone         Returns TRUE when all VBComponents are in sync.
' - AppRunChanged   Called via Application.Run by CommonButton: Synchronizes a
'                   code change in a Sync-Source-Workbook's VBComponent with
'                   the corresponding Sync-Target-Workbook's VBComponent.
' - AppRunObsolete  Called via Application.Run by CommonButton: Removes
'                   VBComponent from the Sync-Target-Workbook's VBProject.
' - CollectChanged  Returns a Collection with all changed VBComponents
' - CollectNew      Returns a Collection with all new VBComponents
' - CollectObsolete Returns a Collection with all obsolete VBComponents.
' - RunNew          Called via Application.Run by CommonButton: Adds a
'                   VBComponent to the provided Workbook's VBProject.
' - Sync            Called by mSync.RunSync when there are still outstanding
'                   VBComponents to be synchronized.
'
' W. Rauschenberger Berlin Dec 2022
' ----------------------------------------------------------------------------
Private Const TITLE_SYNC_COMPS  As String = "VB-Project Synchronization: VB-Components"

Private dctKnownInSync      As Dictionary
Private dctKnownNew         As Dictionary
Private dctKnownObsolete    As Dictionary
Private dctKnownChanged     As Dictionary
Private dctKnownSkppdSource As Dictionary
Private dctKnownSkppdTarget As Dictionary

Public Property Get KnownChanged(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownChanged Is Nothing _
    Then KnownChanged = mSyncComps.dctKnownChanged.Exists(nk_id)
End Property

Public Property Let KnownChanged(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownChanged, nk_id
End Property

Private Property Get KnownInSync(Optional ByVal nis_id As String) As Boolean
    If Not dctKnownInSync Is Nothing _
    Then KnownInSync = dctKnownInSync.Exists(nis_id)
End Property

Private Property Let KnownInSync(Optional ByVal nis_id As String, _
                                         ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownInSync, nis_id
End Property

Public Property Get KnownNew(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownNew Is Nothing _
    Then KnownNew = mSyncComps.dctKnownNew.Exists(nk_id)
End Property

Public Property Let KnownNew(Optional ByVal nk_id As String, _
                                          ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownNew, nk_id
End Property

Public Property Get KnownObsolete(Optional ByVal nk_id As String) As Boolean
    KnownObsolete = mSyncComps.dctKnownObsolete.Exists(nk_id)
End Property

Public Property Let KnownObsolete(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownObsolete, nk_id
End Property

Private Property Get KnownSkppdSource(Optional ByVal k_vbc_id As String) As Boolean
    If Not dctKnownSkppdSource Is Nothing _
    Then KnownSkppdSource = dctKnownSkppdSource.Exists(k_vbc_id)
End Property

Private Property Let KnownSkppdSource(Optional ByVal k_vbc_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownSkppdSource, k_vbc_id
End Property

Private Property Get KnownSkppdTarget(Optional ByVal k_nme_id As String) As Boolean
    If Not dctKnownSkppdTarget Is Nothing _
    Then KnownSkppdTarget = dctKnownSkppdTarget.Exists(k_nme_id)
End Property

Private Property Let KnownSkppdTarget(Optional ByVal k_nme_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownSkppdTarget, k_nme_id
End Property

Public Function AllDone(ByVal d_wbk_source As Workbook, _
                        ByVal d_wbk_target As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Updates the collection of the VBComponents in sync and returns TRUE when
' there are no more VBComponent synchronizations outstanding by considering
' the number of source and target components as well as thos skipped.
' ----------------------------------------------------------------------------
    Const PROC = "AllDone"
    
    Dim lInSync         As Long
    Dim lSkippedSource  As Long
    Dim lSkippedTarget  As Long
    Dim lSourceComps    As Long
    Dim lTargetComps    As Long
    
    mBasic.BoP ErrSrc(PROC)
    lInSync = CollectInSync(d_wbk_source, d_wbk_target)
    If Not dctKnownSkppdSource Is Nothing And lSyncMode = SyncByKind Then lSkippedSource = dctKnownSkppdSource.Count
    If Not dctKnownSkppdTarget Is Nothing And lSyncMode = SyncByKind Then lSkippedTarget = dctKnownSkppdTarget.Count
    lSourceComps = d_wbk_source.VBProject.VBComponents.Count
    lTargetComps = d_wbk_target.VBProject.VBComponents.Count
    
    If lTargetComps = lInSync + lSkippedTarget _
    And lSourceComps = lInSync + lSkippedSource Then
        AllDone = True
        mMsg.MsgInstance TITLE_SYNC_COMPS, True ' unload the mode-less displayed message window
        wsSyncLog.SummaryDone("VBComponent") = True
        DueSyncKindOfObjects.DeQueue , enSyncObjectKindVBComponent
    End If
    mBasic.EoP ErrSrc(PROC)

End Function

Private Sub AppRunChanged()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes a code change in the
' VBComponent (sync_vbc) in the Sync-Source-Workbook with the
' corresponding VBComponent in the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunChanged"
    
    On Error GoTo eh
    Dim i           As Long
    Dim SourceComp  As New clsComp
    Dim TargetComp  As New clsComp
    Dim va          As Variant
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim sId         As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkSource = mSync.Source
    Set wbkTarget = mSync.TargetWorkingCopy
    mSync.AppRunInit
    va = Split(AppRunIdsChanged(enSyncObjectKindVBComponent), ",")
    mSync.Progress enSyncObjectKindVBComponent, enSyncStepSyncing, enSyncActionChangeCode, 0
    
    For i = LBound(va) To UBound(va)
        Set SourceComp.Wrkbk = wbkSource
        Set TargetComp.Wrkbk = wbkTarget
        With SourceComp
            .Wrkbk = wbkSource
        End With
        With TargetComp
            .Wrkbk = wbkTarget
        End With
        sId = va(i)
        If SourceComp.VBComp.Type <> vbext_ct_MSForm Then
            mUpdate.ByCodeReplace b_source:=SourceComp _
                                , b_target:=TargetComp
            wsSyncLog.Done "modified", "VBComponent", sId, "updated", "VBComponent's code lines replaced with code lines of source component"
        Else
            mUpdate.ByReImport b_comp:=SourceComp.VBComp.Name _
                             , b_source:=SourceComp.ExpFileFullName
            wsSyncLog.Done "modified", "VBComponent", sId, "updated", "VBComponent removed and source VBComponent's ExportFile re-imported"
        End If
        Set TargetComp = Nothing
        Set SourceComp = Nothing
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                     , p_sync_step:=enSyncStepSyncing _
                     , p_sync_action:=enSyncActionChangeCode _
                     , p_count:=i + 1
    Next i
    
    dctKnownChanged.RemoveAll
    mSync.AppRunTerminate
    
xt: Services.MessageUnload TITLE_SYNC_COMPS
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppRunIdsChanged(ByVal a_kind As enSyncKindOfObject) As String
    AppRunIdsChanged = DueSyncIdsByAction(a_kind, enSyncActionChanged)
End Function

Private Sub AppRunNew()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Adds the VBComponent (sync_vbc)
' to the provided Workbook's (sync_wbk_target) VBProject.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunNew"
    
    On Error GoTo eh
    Dim i           As Long
    Dim SourceComp  As clsComp
    Dim va          As Variant
    Dim vbc         As VBComponent
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim sId         As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkSource = mSync.Source
    Set wbkTarget = mSync.TargetWorkingCopy
    va = Split(AppRunNewIds(enSyncObjectKindVBComponent), ",")
    mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                 , p_sync_step:=enSyncStepSyncing _
                 , p_sync_action:=enSyncActionAddNew _
                 , p_count:=0
    
    mSync.AppRunInit
    
    For i = LBound(va) To UBound(va)
        Set SourceComp = New clsComp
        Set vbc = GetComp(wbkSource, va(i))
        sId = SyncId(vbc, wbkSource)
        With SourceComp
            .Wrkbk = wbkSource
            wbkTarget.VBProject.VBComponents.Import .ExpFileFullName
            wsSyncLog.Done "new", "VBComponent", sId, "added", "by import of Sync-Source-Workbook's corresonding Export-File"
        End With
        Set SourceComp = Nothing
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                 , p_sync_step:=enSyncStepSyncing _
                 , p_sync_action:=enSyncActionAddNew _
                 , p_count:=i + 1

    Next i
    
    dctKnownNew.RemoveAll
    mSync.AppRunTerminate
    
xt: Services.MessageUnload TITLE_SYNC_COMPS
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

                     
Private Sub AppRunObsolete()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes the VBComponent (sync_vbc)
' from the provided Sync-Target-Workbook's VBProject.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunObsolete"
    
    On Error GoTo eh
    Dim i           As Long
    Dim va          As Variant
    Dim vbc         As VBComponent
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim sId         As String
    
    mBasic.BoP ErrSrc(PROC)
    Services.MessageUnload TITLE_SYNC_COMPS ' for the next display
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.Source
    va = Split(AppRunObsoleteIds(enSyncObjectKindVBComponent), ",")
    mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                 , p_sync_step:=enSyncStepSyncing _
                 , p_sync_action:=enSyncActionRemoveObsolete _
                 , p_count:=0
    mSync.AppRunInit
    
    For i = LBound(va) To UBound(va)
        Set vbc = GetComp(wbkTarget, va(i))
        sId = SyncId(vbc, wbkTarget)
        wbkTarget.VBProject.VBComponents.Remove vbc
        wsSyncLog.Done "obsolete", "VBComponent", sId, "removed"
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                     , p_sync_step:=enSyncStepSyncing _
                     , p_sync_action:=enSyncActionRemoveObsolete _
                     , p_count:=i + 1
    Next i
    dctKnownObsolete.RemoveAll
    
xt: Services.MessageUnload TITLE_SYNC_COMPS
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AppRunSyncAll()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all VB-Components
' in the Sync-Target-Workbook with those in the Sync-Source-Workbook by
' removing obsolete, adding new, and updating changed VB-Components.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunSyncAll"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If dctKnownChanged.Count > 0 Then mSyncComps.AppRunChanged     ' Synchronize changed VB-Components
    If dctKnownNew.Count > 0 Then mSyncComps.AppRunNew             ' Synchronize new VB-Components
    If dctKnownObsolete.Count > 0 Then mSyncComps.AppRunObsolete   ' Synchronize obsolete VB-Components

xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        Services.MessageUnload TITLE_SYNC_COMPS
        mSync.RunSync
    End If
    Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function CodeDiffers(ByVal d_vbc_source As VBComponent, _
                             ByVal d_vbc_target As VBComponent) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the code in the provided VBComponents is different.
' ------------------------------------------------------------------------------
    Const PROC = "CodeDiffers"
    
    On Error GoTo eh
    Dim vSource As Variant
    Dim vTarget As Variant
    
    vSource = GetCodeAsArray(d_vbc_source)
    vTarget = GetCodeAsArray(d_vbc_target)
    If mBasic.ArryIsAllocated(vSource) And mBasic.ArryIsAllocated(vTarget) Then
        CodeDiffers = mBasic.ArryDiffers(vSource, vTarget, True)
    End If
    
xt: Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Collect(ByVal c_wbk_source As Workbook, _
                   ByVal c_wbk_target As Workbook)
' ------------------------------------------------------------------------------
' Collects obsolete, new and code changed VB-Components.
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim vbcSource   As VBComponent
    Dim vbcTarget   As VBComponent
    Dim sId         As String
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    If mSyncComps.AllDone(c_wbk_source, c_wbk_target) Then GoTo xt
    If lSyncMode = SyncByKind Then mSync.InitDueSyncs
    If Not dctKnownSkppdSource Is Nothing Then dctKnownSkppdSource.RemoveAll
    If Not dctKnownSkppdTarget Is Nothing Then dctKnownSkppdTarget.RemoveAll
    
    '~~ Collect VB-Components the code has changed
    If DueCollect("Changed") Then
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                     , p_sync_step:=enSyncStepSyncing _
                     , p_sync_action:=enSyncActionChangeCode _
                     , p_count:=0
        For Each vbcSource In c_wbk_source.VBProject.VBComponents
            sId = SyncId(vbcSource, c_wbk_source)
            If Not KnownSource(sId) Then
                If Corresponding(c_this_vbc:=vbcSource _
                               , c_this_wbk:=c_wbk_source _
                               , c_that_wbk:=c_wbk_target _
                               , c_vbc_result:=vbcTarget) = "1:1" _
                Then
                    If CodeDiffers(vbcSource, vbcTarget) Then
                        mSync.DueSyncLet , enSyncObjectKindVBComponent, enSyncActionChangeCode, , sId
                    End If
                    If vbcSource.Name <> vbcTarget.Name Then
                        If IsWrkbkDocMod(vbcSource, c_wbk_source) Then
                            If IsWrkbkDocMod(vbcSource, c_wbk_source) Then Stop
                            mSync.DueSyncLet , enSyncObjectKindVBComponent, enSyncActionChangeName, , sId
                        Else
                            KnownSkppdSource(sId) = True
                            KnownSkppdTarget(SyncId(vbcTarget, c_wbk_target)) = True
                        End If
                    End If
                End If
            End If
            mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                         , p_sync_step:=enSyncStepCollecting _
                         , p_sync_action:=enSyncActionChangeCode _
                         , p_count:=dctKnownChanged.Count
        Next vbcSource
    End If

    '~~ Collect New VB-Components
    If DueCollect("New") Then
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                     , p_sync_step:=enSyncStepCollecting _
                     , p_sync_action:=enSyncActionAddNew _
                     , p_count:=0
        For Each vbcSource In c_wbk_source.VBProject.VBComponents
            sId = SyncId(vbcSource, c_wbk_source)
            If Not KnownSource(sId) Then
                If Corresponding(c_this_vbc:=vbcSource _
                               , c_this_wbk:=c_wbk_source _
                               , c_that_wbk:=c_wbk_target) = "1:0" _
                Then
                    If vbcSource.Type <> vbext_ct_Document Then
                        mSync.DueSyncLet , enSyncObjectKindVBComponent, enSyncActionAddNew, , sId
                    Else
                        '~~ When the component type is Document-Module it may be a Workbook component which (still)
                        '~~ has a different name or a Worksheet which still has a different name or does not exist
                        If mComp.IsSheetDocMod(vbcSource, c_wbk_source, wshSource) Then
                            If Not mWsh.Exists(x_wbk:=c_wbk_target, x_wsh:=wshSource, x_wsh_result:=wshTarget) Then
                                '~~ The apparently new VBComponent which will be synchronized
                                '~~ along with the yet not done Worksheet synchronization where
                                '~~ the new Workbook will be cloned
                                KnownSkppdSource(sId) = True
                            End If
                        End If
                        KnownSkppdSource(sId) = True
                    End If
                End If
            End If
            mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                         , p_sync_step:=enSyncStepCollecting _
                         , p_sync_action:=enSyncActionAddNew _
                         , p_count:=dctKnownNew.Count
        Next vbcSource
    End If

    '~~ Collect obsolete VB-Components
    If DueCollect("Obsolete") Then
        mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                     , p_sync_step:=enSyncStepCollecting _
                     , p_sync_action:=enSyncActionRemoveObsolete _
                     , p_count:=0
        For Each vbcTarget In c_wbk_target.VBProject.VBComponents
            sId = SyncId(vbcTarget, c_wbk_target)
            If Not KnownTarget(sId) Then
                If Corresponding(c_this_vbc:=vbcTarget _
                               , c_this_wbk:=c_wbk_target _
                               , c_that_wbk:=c_wbk_source) = "0:1" _
                Then
                    If vbcTarget.Type <> vbext_ct_Document Then
                        mSync.DueSyncLet , enSyncObjectKindVBComponent, enSyncActionRemoveObsolete, , sId
                    Else
                        '~~ A not existing Document-Module indicates a new Worksheet yet not synchronized
                        KnownSkppdTarget(sId) = True
                    End If
                End If
            End If
            mSync.Progress p_kind:=enSyncObjectKindVBComponent _
                         , p_sync_step:=enSyncStepCollecting _
                         , p_sync_action:=enSyncActionRemoveObsolete _
                         , p_count:=dctKnownObsolete.Count
        Next vbcTarget
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Collected(ByVal c_action As enSyncAction) As Long
    Select Case True
        Case mSync.SyncActionIsChange(c_action):    Collected = dctKnownChanged.Count
        Case c_action = enSyncActionRemoveObsolete: Collected = dctKnownObsolete.Count
        Case c_action = enSyncActionAddNew:         Collected = dctKnownNew.Count
    End Select
End Function

Private Function CollectInSync(ByVal c_wbk_source As Workbook, _
                               ByVal c_wbk_target As Workbook) As Long
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "CollectInSync"
    
    On Error GoTo eh
    Dim vbcSource   As VBComponent
    Dim vbcTarget   As VBComponent
    Dim sIdSource   As String
    
    mBasic.BoP ErrSrc(PROC)
    If dctKnownInSync Is Nothing Then Set dctKnownInSync = New Dictionary
    For Each vbcSource In c_wbk_source.VBProject.VBComponents
        sIdSource = SyncId(vbcSource, c_wbk_source)
        If Not KnownInSync(sIdSource) Then
            If Corresponding(c_this_vbc:=vbcSource _
                           , c_this_wbk:=c_wbk_source _
                           , c_that_wbk:=c_wbk_target _
                           , c_vbc_result:=vbcTarget) = "1:1" _
                Then
                If Not CodeDiffers(vbcSource, vbcTarget) And vbcSource.Name = vbcTarget.Name Then
                    KnownInSync(sIdSource) = True
                End If
            End If
        End If
    Next vbcSource
                          
xt: CollectInSync = dctKnownInSync.Count
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Corresponding(ByVal c_this_vbc As VBComponent, _
                               ByVal c_this_wbk As Workbook, _
                               ByVal c_that_wbk As Workbook, _
                      Optional ByRef c_vbc_result As VBComponent) As String
' ------------------------------------------------------------------------------
' Returns:
'   "1:0" when the VBComponent exists in the Sync-Source-Workbook but not in the
'         Sync-Target-Workbook
'   "0:1" when the VBComponent exists in the Sync-Target-Workbook but not in the
'         Sync-Source-Workbook
'   "1:1" when the VBComponent exists in both Workbooks
' ------------------------------------------------------------------------------
    Const PROC = "Corresponding"
    
    On Error GoTo eh
    Dim sCorrespondingSheet     As String
    Dim sCorrespondingSource    As String:  sCorrespondingSource = "0"
    Dim sCorrespondingTarget    As String:  sCorrespondingTarget = "0"
    Dim vbc                     As VBComponent
    Dim wbkOpposite             As Workbook
    Dim wbkSource               As Workbook
    Dim wbkTarget               As Workbook
    Dim wshThis                 As Worksheet
    Dim wshOpposite             As Worksheet
    Dim bThisIsSource           As Boolean
    Dim bThisIsTarget           As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.Source

'    Debug.Print ErrSrc(PROC) & ": " &  wbkTarget.CodeName
'    Debug.Print ErrSrc(PROC) & ": " &  wbkSource.CodeName
    
    If c_this_wbk Is wbkTarget Then
        bThisIsTarget = True
        sCorrespondingTarget = "1"
        Set wbkOpposite = c_that_wbk
    ElseIf c_this_wbk Is wbkSource Then
        bThisIsSource = True
        sCorrespondingSource = "1"
        Set wbkOpposite = c_that_wbk
    End If
    
    If mComp.IsSheetDocMod(c_this_vbc, c_this_wbk, wshThis) Then
        '~~ Obtain the VBComponent's corresponding Worksheet in the opposite Workbook
        sCorrespondingSheet = mSyncSheets.Corresponding(c_wsh:=wshThis _
                                                      , c_wbk:=wbkOpposite _
                                                      , c_quality:=enOrNameCodeName _
                                                      , c_wsh_result:=wshOpposite)
        If sCorrespondingSheet = "1:1" Then
            For Each vbc In wbkOpposite.VBProject.VBComponents
                If vbc.Name = wshOpposite.CodeName Then
                    Set c_vbc_result = vbc
                    If bThisIsSource Then
                        sCorrespondingTarget = "1"
                    ElseIf bThisIsTarget Then
                        sCorrespondingSource = "1"
                    End If
                    Exit For
                End If
            Next vbc
        End If
    ElseIf IsWrkbkDocMod(c_this_vbc, c_this_wbk) Then
        '~~ Return the opposite Workbook's Workbook Document-Module
        For Each vbc In wbkOpposite.VBProject.VBComponents
            If IsWrkbkDocMod(vbc, wbkOpposite) Then
                Set c_vbc_result = vbc
                If bThisIsSource Then
                    sCorrespondingTarget = "1"
                ElseIf bThisIsTarget Then
                    sCorrespondingSource = "1"
                End If
            End If
        Next vbc
    
    Else
        For Each vbc In wbkOpposite.VBProject.VBComponents
            If vbc.Name = c_this_vbc.Name Then
                Set c_vbc_result = vbc
                If bThisIsSource Then
                    sCorrespondingTarget = "1"
                ElseIf bThisIsTarget Then
                    sCorrespondingSource = "1"
                End If
                GoTo xt
            End If
        Next vbc
    End If

xt: Corresponding = sCorrespondingSource & ":" & sCorrespondingTarget
    mBasic.EoP ErrSrc(PROC)
    Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function DueCollect(ByVal d_c As String) As Boolean
    Select Case d_c
        Case "New":         DueCollect = dctKnownNew Is Nothing
                            If DueCollect Then Set dctKnownNew = New Dictionary
        Case "Changed":     DueCollect = dctKnownChanged Is Nothing
                            If DueCollect Then Set dctKnownChanged = New Dictionary
        Case "Obsolete":    DueCollect = dctKnownObsolete Is Nothing
                            If DueCollect Then Set dctKnownObsolete = New Dictionary
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncComps." & s
End Function

Private Function GetCodeAsArray(ByVal gf_vbc As VBComponent) As Variant
' ------------------------------------------------------------------------------
' Returns the code of the VBComponent (gf_vbc) As File.
' ------------------------------------------------------------------------------
    Const PROC = "GetCodeAsArray"
    
    On Error GoTo eh
    Dim s       As String
    Dim sSplit  As String
    
    With gf_vbc.CodeModule
        If .CountOfLines > 0 Then
            s = .Lines(1, .CountOfLines)
            sSplit = SplitString(s)
            GetCodeAsArray = Split(s, sSplit)
        End If
    End With
    
xt: Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

              
Private Function GetComp(ByVal g_wbk As Workbook, _
                Optional ByVal g_vbc_id As String = vbNullString, _
                Optional ByVal g_vbc_name As String = vbNullString, _
                Optional ByRef g_vbc_result As VBComponent) As VBComponent
' ------------------------------------------------------------------------------
' Returns the VBComponent in the Workbook (g_wbk) identified either by its
' SyncId (g_vbc_id) or its Name (g_vbc_name).
' ------------------------------------------------------------------------------
    Const PROC = "GetComp"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    
    For Each vbc In g_wbk.VBProject.VBComponents
        If g_vbc_id <> vbNullString Then
            If SyncId(vbc, g_wbk) = g_vbc_id Then
                Set GetComp = vbc
                Set g_vbc_result = vbc
                Exit For
            End If
        Else
            If vbc.Name = g_vbc_name Then
                Set GetComp = vbc
                Set g_vbc_result = vbc
                Exit For
            End If
        End If
    Next vbc
    
    If GetComp Is Nothing _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The Workbook's (" & g_wbk.Name & "' has no VBComponent named '" & g_vbc_id & "'!"

xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Initialize()
    Set dctKnownInSync = Nothing
    Set dctKnownChanged = Nothing
    Set dctKnownNew = Nothing
    Set dctKnownObsolete = Nothing
    Set dctKnownSkppdSource = Nothing
    Set dctKnownSkppdTarget = Nothing
End Sub

Private Function IsWrkbkDocMod(ByVal i_vbc As VBComponent, _
                               ByVal i_wbk As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Retunrs True when the VBComponent represents the Workbook.
' ------------------------------------------------------------------------------
    IsWrkbkDocMod = i_vbc.Type = vbext_ct_Document And i_vbc.Name = i_wbk.CodeName
End Function

Private Function KnownSource(ByVal k_id As String) As Boolean
    KnownSource = KnownInSync(k_id) _
               Or KnownNew(k_id) _
               Or KnownChanged(k_id) _
               Or KnownSkppdSource(k_id)
End Function

Private Function KnownTarget(ByVal k_id As String) As Boolean
    KnownTarget = KnownInSync(k_id) _
               Or KnownObsolete(k_id) _
               Or KnownChanged(k_id) _
               Or KnownSkppdTarget(k_id)
End Function

Private Function SplitString(ByVal s_s As String) As String
    
    Select Case True
        Case InStr(s_s, vbCrLf) <> 0: SplitString = vbCrLf
        Case InStr(s_s, vbCr) <> 0:   SplitString = vbCr
        Case InStr(s_s, vbLf) <> 0:   SplitString = vbLf
    End Select
    If Len(SplitString) = 0 Then SplitString = vbCrLf
    
End Function

Private Function SyncId(ByVal s_vbc As VBComponent, _
                        ByVal s_wbk As Workbook) As String
' ------------------------------------------------------------------------------
' Returns a VBComponent-Name used as id throughout VBComponent synchronization.
' ------------------------------------------------------------------------------
    SyncId = s_vbc.Name & " (" & mComp.TypeString(s_vbc, s_wbk) & ")"
End Function

Public Sub SyncKind(ByVal s_wbk_source As Workbook, _
                    ByVal s_wbk_target As Workbook)
' ----------------------------------------------------------------------------
' Displays a dialog to perform all required action to synchronize the VB-
' Components in the Sync-Target-Workbook with those in the Sync-Source-
' Workbook.
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
    
    mCompManClient.Events ErrSrc(PROC), False
    mBasic.BoP ErrSrc(PROC)
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_COMPS)
    mSync.MonitorStep "Synchronizing VB-Components"
    mSyncComps.Collect s_wbk_source, s_wbk_target
    Bttn1 = "Perform all VBComponent synchronization actions" & vbLf & "listed above"
    
    sDueSyncs = mSync.DueSyncs(enSyncObjectKindVBComponent)
    With Msg.Section(1)
        .Text.MonoSpaced = True
        .Text.Text = sDueSyncs
    End With
    bDueObsolete = InStr(sDueSyncs, enSyncActionRemoveObsolete) <> 0
    bDueNew = InStr(sDueSyncs, SyncActionString(enSyncActionAddNew)) <> 0
    bDueChanged = InStr(sDueSyncs, SyncActionString(enSyncActionChangeCode)) <> 0
    lSection = 1
    
    If bDueObsolete Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_REMOVE_OBSOLETE & ":"
            End With
            .Text.Text = "Considered because it has no corresponding source VBComponent."
        End With
    End If
    If bDueNew Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_ADD_NEW & ":"
            End With
            .Text.Text = "Considered new because no corresponding VBComponent exists in the Sync-Target-Workbook."
        End With
    End If
    If bDueChanged Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_CHANGE_CODE & ":"
            End With
            .Text.Text = "The code of the VBComponent in the Sync-Source-Workbook has been modified."
        End With
    End If
    
    mMsg.BttnAppRun AppRunArgs, Bttn1 _
                                , ThisWorkbook _
                                , "mSyncComps.AppRunSyncAll"
    Set cllButtons = mMsg.Buttons(Bttn1)
    
    If bDueObsolete Or bDueNew Or bDueChanged Then
        With Msg.Section(8).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization, section Name Synchronization:"
            .FontColor = rgbBlue
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_SYNC_CHAPTER_NAMES
        End With
        
        '~~ Display the mode-less dialog for the Names synchronization to run
        mMsg.Dsply dsply_title:=TITLE_SYNC_COMPS _
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

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

