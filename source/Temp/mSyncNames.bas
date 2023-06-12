Attribute VB_Name = "mSyncNames"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mSyncNames
'
'
' W. Rauschenberger, Berlin, June 2022
' ------------------------------------------------------------------------
Private Const TITLE_SYNC_NAMES = "VB-Project Synchronization: Names"         ' Identifies the synchronization dialog

Private dctKnownChanged     As Dictionary
Private dctKnownInSync      As Dictionary
Private dctKnownMultiple    As Dictionary
Private dctKnownNew         As Dictionary
Private dctKnownObsolete    As Dictionary
Private dctKnownSkppdSource As Dictionary
Private dctKnownSkppdTarget As Dictionary
Private dctMultiple         As Dictionary
Private dctMultipleSource   As Dictionary
Private dctMultipleTarget   As Dictionary
Private lMultiply           As Long

Private Enum enCorrespondingNamesQuality
    '~~ All possible qualities of Names corresponding of which only some may be used
    enAndNameRefScope   ' Equal Name, RefersTo, and Scope
    enOrNameRef         ' Either Name or RefersTo is equal
End Enum

Public Property Get KnownChanged(Optional ByVal k_nme_id As String, _
                                 Optional ByVal k_nme_item As Variant) As Boolean
    k_nme_item = k_nme_item
    If Not dctKnownChanged Is Nothing Then
        KnownChanged = dctKnownChanged.Exists(k_nme_id)
    End If
End Property

Public Property Let KnownChanged(Optional ByVal k_nme_id As String, _
                                 Optional ByVal k_nme_item As Variant, _
                                          ByVal b As Boolean)
    If b Then mSyncNames.CollectKnown dctKnownChanged, k_nme_id, k_nme_item
End Property

Private Property Get KnownInSync(Optional ByVal k_nme_id As String) As Boolean
    If Not dctKnownInSync Is Nothing _
    Then KnownInSync = dctKnownInSync.Exists(k_nme_id)
End Property

Private Property Let KnownInSync(Optional ByVal k_nme_id As String, _
                                          ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownInSync, k_nme_id
End Property

Private Property Get KnownMultiple(Optional ByVal k_nme_id As String) As Boolean
    If Not dctKnownMultiple Is Nothing _
    Then KnownMultiple = dctKnownMultiple.Exists(k_nme_id)
End Property

Private Property Let KnownMultiple(Optional ByVal k_nme_id As String, _
                                            ByVal b As Boolean)
    If b Then
        If Not dctKnownMultiple.Exists(k_nme_id) Then dctKnownMultiple.Add k_nme_id, k_nme_id
    End If
End Property

Public Property Get KnownNew(Optional ByVal k_nme_id As String, _
                             Optional ByVal k_item As Variant) As Boolean
    k_item = k_item
    If Not dctKnownNew Is Nothing _
    Then KnownNew = dctKnownNew.Exists(k_nme_id)
End Property

Public Property Let KnownNew(Optional ByVal k_nme_id As String, _
                             Optional ByVal k_item As Variant, _
                                      ByVal b As Boolean)
    If b Then mSyncNames.CollectKnown dctKnownNew, k_nme_id, k_item
End Property

Public Property Get KnownObsolete(Optional ByVal k_nme_id As String, _
                                  Optional ByVal k_item As Variant) As Boolean
    k_item = k_item
    If Not dctKnownObsolete Is Nothing Then _
    KnownObsolete = dctKnownObsolete.Exists(k_nme_id)
End Property

Public Property Let KnownObsolete(Optional ByVal k_nme_id As String, _
                                  Optional ByVal k_item As Variant, _
                                           ByVal b As Boolean)
    If b Then mSyncNames.CollectKnown dctKnownObsolete, k_nme_id, k_item
End Property

Private Property Get KnownSkppdSource(Optional ByVal k_nme_id As String) As Boolean
    If Not dctKnownSkppdSource Is Nothing _
    Then KnownSkppdSource = dctKnownSkppdSource.Exists(k_nme_id)
End Property

Private Property Let KnownSkppdSource(Optional ByVal k_nme_id As String, _
                                           ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownSkppdSource, k_nme_id
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
' ------------------------------------------------------------------------------
' Returns TRUE when the number of in-sync Names plus any sipped/2 are equal
' the Names in the Sync-Source- and Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "AllDone"
    
    On Error GoTo eh
    Dim lInSync         As Long
    Dim lSkippedSource  As Long
    Dim lSkippedTarget  As Long
    Dim lValidTarget    As Long
    Dim lValidSource    As Long
    
    mBasic.BoP ErrSrc(PROC)
    lValidTarget = ValidNames(d_wbk_target)
    lValidSource = ValidNames(d_wbk_source)
    lInSync = CollectInSync(d_wbk_source, d_wbk_target)
    
    If Not dctKnownSkppdSource Is Nothing Then lSkippedSource = dctKnownSkppdSource.Count
    If Not dctKnownSkppdTarget Is Nothing Then lSkippedTarget = dctKnownSkppdTarget.Count
    If lValidTarget = lInSync + lSkippedTarget _
    And lValidSource = lInSync + lSkippedSource Then
        AllDone = True
        mMsg.MsgInstance TITLE_SYNC_NAMES, True
        wsSyncLog.SummaryDone("Name") = True
        mSync.DueSyncKindOfObjects.DeQueue , enSyncObjectKindName
    Else
        If wsService.MaxLenNameSyncId = 0 Then mSyncNames.MaxLenNameId d_wbk_source, d_wbk_target
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function AllNames(ByVal a_in As Variant, _
                         ByRef a_dct As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Return as Dictionary with all Names in (a_in) which may be a Workbook or a
' Worksheet.
' ----------------------------------------------------------------------------
    
    Dim nme As Name
    Dim dct As New Dictionary
    
    For Each nme In a_in.Names
        dct.Add SyncId(nme), nme
    Next nme
    Set AllNames = dct
    Set a_dct = dct
    Set dct = Nothing
                         
End Function

Private Function AnyForBeingSynced() As Boolean
    If Not dctKnownObsolete Is Nothing Then
        If dctKnownObsolete.Count > 0 Then
            AnyForBeingSynced = True
            Exit Function
        End If
    End If
    If Not dctKnownNew Is Nothing Then
        If AnyForBeingSynced = dctKnownNew.Count > 0 Then
            AnyForBeingSynced = True
            Exit Function
        End If
    End If
    If Not dctKnownChanged Is Nothing Then
        If AnyForBeingSynced = dctKnownChanged.Count > 0 Then
            AnyForBeingSynced = True
            Exit Function
        End If
    End If
    If Not dctKnownMultiple Is Nothing Then
        If AnyForBeingSynced = dctKnownMultiple.Count <> 0 Then
            AnyForBeingSynced = True
            Exit Function
        End If
    End If
End Function

Private Function AppErr(ByVal app_err_no As Long) As Long
' ----------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Private Sub AppRunChanged()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Changes the Name property of the
' Name object in the Sync-Target-Workbook which refers to the Name's (rc_nme)
' Range to the Name property of the Name object (rc_nme).
' ------------------------------------------------------------------------------
    Const PROC = "AppRunChanged"
    
    On Error GoTo eh
    Dim nmeSource   As Name
    Dim nmeTarget   As Name
    Dim sOldName    As String
    Dim sOldScope   As String
    Dim sOldRange   As String
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim bName       As Boolean
    Dim bRange      As Boolean
    Dim bScope      As Boolean
    Dim bNone       As Boolean
    Dim vSource     As Variant
    Dim vTarget     As Variant
    Dim i           As Long
    Dim sIdsSource  As String
    Dim sIdsTarget  As String
    Dim sIdTarget   As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    sIdsSource = AppRunChangedIdsSource
    sIdsTarget = AppRunChangedIdsTarget
    vSource = Split(sIdsSource, ",")
    vTarget = Split(sIdsTarget, ",")
    Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionChanged, 0)
    AppRunInit
    
    For i = LBound(vSource) To UBound(vSource)
        sIdTarget = vTarget(i)
        GetName vSource(i), wbkSource, nmeSource
        If nmeSource Is Nothing _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Name identified '" & vSource(i) & "' does not exist in Workbook '" & wbkSource.Name & "'!"
        GetName sIdTarget, wbkTarget, nmeTarget
        If nmeTarget Is Nothing _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Name identified '" & sIdTarget & "' does not exist in Workbook '" & wbkTarget.Name & "'!"
        
        PropertiesDiffer nmeSource, nmeTarget, bName, bRange, bScope, bNone
        sOldName = mNme.MereName(nmeTarget)
        sOldRange = nmeTarget.RefersTo
        sOldScope = ScopeName(nmeTarget)
        
        mNme.ChangeProperties p_target_nme:=nmeTarget _
                            , p_final_name:=MereName(nmeSource) _
                            , p_final_rng:=CorrespondingRange(nmeSource, wbkSource, wbkTarget) _
                            , p_final_scope:=CorrespondingScope(nmeSource, wbkTarget) _
                            , p_name_changed:=bName _
                            , p_scope_changed:=bScope
        
        With wsSyncLog
            If bName Then .Done "change", "Name", sIdTarget, "changed", "Name changed from " & sOldName & " to " & nmeTarget.Name
            If bRange Then .Done "change", "Name", sIdTarget, "changed", "Range changed from " & sOldRange & " to " & nmeTarget.RefersTo
            If bScope Then .Done "change", "Name", sIdTarget, "changed", "Scope changed from " & sOldScope & " to " & mNme.ScopeName(nmeTarget)
        End With
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionChanged, i + 1)
    Next i
    dctKnownChanged.RemoveAll ' indicates done
    AppRunTerminate
    
xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        Services.MessageUnload TITLE_SYNC_NAMES
        mSync.RunSync
    End If
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppRunChangedIdsSource() As String
    AppRunChangedIdsSource = mSync.DueSyncIdsByAction(enSyncObjectKindName, enSyncActionChanged, "to")
End Function

Private Function AppRunChangedIdsTarget() As String
    AppRunChangedIdsTarget = DueSyncIdsByAction(enSyncObjectKindName, enSyncActionChanged, "from")
End Function

Private Sub AppRunMultiple()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes all target Names identi-
' fied by their SyncId (ra_nme_target_ids) and replace them by adding the source
' Names (ra_nme_source_ids).
' ------------------------------------------------------------------------------
    Const PROC = "AppRunMultiple"
    
    On Error GoTo eh
    Dim i           As Long
    Dim nme         As Name
    Dim Rng         As Range
    Dim sAddress    As String
    Dim sName       As String
    Dim sSheetName  As String
    Dim va          As Variant
    Dim vScope      As Variant
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wsh         As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    va = Split(AppRunMultipleIdsTarget, ",")
    Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionMultipleRemove, 0)
    mSync.AppRunInit
    
    For i = LBound(va) To UBound(va)
        If Exists(va(i), wbkTarget, nme) Then
            Services.ServicedItem = nme
            nme.Delete
            wsSyncLog.Done "obsolete", "Name", SyncId(nme), "removed", "Multiple Name removed from Sync-Target-Workbook (working copy)"
        
        End If
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionMultipleRemove, i + 1)
    Next i
    
    va = Split(AppRunMultipleIdsSource, ",")
    Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionMultipleAdd, 0)
    
    For i = LBound(va) To UBound(va)
        GetName va(i), wbkSource, nme
        If nme Is Nothing _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Name identified '" & va(i) & "' does not exist in Workbook '" & wbkSource.Name & "'!"

        sName = mNme.MereName(nme)
        sSheetName = Replace(Split(nme.RefersTo, "!")(0), "=", vbNullString)
        sAddress = Split(nme.RefersTo, "!")(1)
        Set wsh = wbkTarget.Worksheets(sSheetName)
        Set Rng = wsh.Range(sAddress)
        If mNme.ScopeIsWorkbook(nme) Then
            Set vScope = wbkTarget
        ElseIf mNme.ScopeIsWorkSheet(nme, sSheetName) Then
            Set vScope = wbkTarget.Worksheets(sSheetName)
        End If
        
        mNme.Create c_name:=sName _
                  , c_rng:=Rng _
                  , c_scope:=vScope _
                  , c_nme:=nme
            
        wsSyncLog.Done "multiple", "Name", va(i), "new ambiguity added", "New! Multiply added to Sync-Target-Workbook (working copy)"
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionMultipleAdd, i + 1)
    Next i
    dctKnownMultiple.RemoveAll ' indicates that all ambiguities had been done
    mSync.AppRunTerminate
    
xt: Services.MessageUnload TITLE_SYNC_NAMES
    mBasic.BoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppRunMultipleIdsSource() As String
    AppRunMultipleIdsSource = DueSyncIdsByAction(enSyncObjectKindName, enSyncActionMultipleSource)
End Function

Private Function AppRunMultipleIdsTarget() As String
    AppRunMultipleIdsTarget = DueSyncIdsByAction(enSyncObjectKindName, enSyncActionMultipleTarget)
End Function

Private Sub AppRunNew()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Adds a Name object to the
' Sync-Target-Workbook with the name (ra_name) referring to
' range (ra_ref).
' ------------------------------------------------------------------------------
    Const PROC = "AppRunNew"
    
    On Error GoTo eh
    Dim nme         As Name
    Dim nmeTarget   As Name
    Dim Rng         As Range
    Dim sAddress    As String
    Dim sName       As String
    Dim sSheetName  As String
    Dim vScope      As Variant
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wsh         As Worksheet
    Dim va          As Variant
    Dim i           As Long
    
    mBasic.BoP ErrSrc(PROC)
    Services.MessageUnload TITLE_SYNC_NAMES ' for the next display
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    va = Split(AppRunNewIds(enSyncObjectKindName), ",")
    Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionAddNew, 0)
    AppRunInit
    
    For i = LBound(va) To UBound(va)
        GetName va(i), wbkSource, nme
        If nme Is Nothing _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Name identified '" & va(i) & "' does not exist in Workbook '" & wbkSource.Name & "'!"
        Services.ServicedItem = nme
        
        sName = mNme.MereName(nme)
        sSheetName = Replace(Split(nme.RefersTo, "!")(0), "=", vbNullString)
        sAddress = Split(nme.RefersTo, "!")(1)
        mSyncSheets.Corresponding c_wsh:=wbkSource.Worksheets(sSheetName) _
                                , c_wbk:=wbkTarget _
                                , c_quality:=enCorrespondingSheetsQuality.enOrNameCodeName _
                                , c_wsh_result:=wsh
        
        Set Rng = wsh.Range(sAddress)
        If mNme.ScopeIsWorkbook(nme) Then
            Set vScope = wbkTarget
        ElseIf mNme.ScopeIsWorkSheet(nme, sSheetName) Then
            Set vScope = wsh
        End If
        
        mNme.Create c_name:=sName _
                  , c_rng:=Rng _
                  , c_scope:=vScope _
                  , c_nme:=nmeTarget
            
        wsSyncLog.Done "new", "Name", va(i), "added", "New! Added to Sync-Target-Workbook (working copy)"
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionAddNew, i + 1)
    Next i
    dctKnownNew.RemoveAll ' indicates that all new Names had been added
    AppRunTerminate
    
xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        Services.MessageUnload TITLE_SYNC_NAMES
        mSync.RunSync
    End If
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AppRunObsolete()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes the Name objects identi-
' fied via a comma separated string if ids (rr_nme_ids) from the Sync-Target-
' Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunObsolete"
    
    On Error GoTo eh
    Dim nme         As Name
    Dim i           As Long
    Dim va          As Variant
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Services.MessageUnload TITLE_SYNC_NAMES ' for the next display
    Set wbkTarget = mSync.TargetWorkingCopy
    va = Split(mSync.AppRunObsoleteIds(enSyncObjectKindName), ",")
    Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionRemoveObsolete, 0)
    mSync.AppRunInit
    
    For i = LBound(va) To UBound(va)
        GetName va(i), wbkTarget, nme
        If Not nme Is Nothing Then
            Services.ServicedItem = nme
            nme.Delete
            wsSyncLog.Done "obsolete", "Name", va(i), "removed", "Obsolete! Removed from Sync-Target-Workbook (working copy)"
        Else
            Debug.Print "The Name '" & va(i) & "' no longer exists!"
        End If
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepSyncing, enSyncActionRemoveObsolete, i + 1)
    Next i
    dctKnownObsolete.RemoveAll ' indicates that all removals had been done
    mSync.AppRunTerminate
    
xt: Services.MessageUnload TITLE_SYNC_NAMES
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AppRunSyncAll()
' ------------------------------------------------------------------------------
' Called when the Synchronization Mode is "Summarized", performs all due Name
' synchronizations
' ------------------------------------------------------------------------------
    Const PROC = "AppRunSyncAll"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)

    If dctKnownChanged.Count > 0 Then
        mSyncNames.AppRunChanged
    End If
    If dctKnownNew.Count > 0 Then
        mSyncNames.AppRunNew
    End If
    If dctKnownObsolete.Count > 0 Then
        mSyncNames.AppRunObsolete
    End If
    If dctKnownMultiple.Count <> 0 Then
        mSyncNames.AppRunMultiple
    End If

xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        Services.MessageUnload TITLE_SYNC_NAMES
        mSync.RunSync
    End If
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

Public Sub Collect(ByVal c_wbk_source As Workbook, _
                   ByVal c_wbk_target As Workbook, _
          Optional ByRef c_terminated As Boolean = False)
' ------------------------------------------------------------------------
' Collects all due Name synchronizations. Any change of a Names's RefersTo
' property (refers a different range) is handled right along with this
' collection. Such changes may be skipped (for being solved manually when
' the synchronization has ended) or decided to interrupt the collection in
' order to solve the issue before the synchronization is continued (the
' Collection ends with c_terminated = TRUE).
' ------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim bNameDiffers        As Boolean
    Dim bNoneDiffers        As Boolean
    Dim bRefToDiffers       As Boolean
    Dim bScopeDiffers       As Boolean
    Dim dctCorrespSource    As Dictionary
    Dim dctCorrespTarget    As Dictionary
    Dim nmeCorrespSource    As Name
    Dim nmeCorrespTarget    As Name
    Dim nmeSource           As Name
    Dim nmeTarget           As Name
    Dim sCorrespondingName As String
    Dim sCorrespondingSheet As String
    Dim sIdSource           As String
    Dim sIdTarget           As String
    Dim vTarget             As Variant
    Dim wsh                 As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    If mSyncNames.AllDone(c_wbk_source, c_wbk_target) Then GoTo xt
    If lSyncMode = SyncByKind Then mSync.InitDueSyncs
    If Not dctKnownSkppdSource Is Nothing Then dctKnownSkppdSource.RemoveAll
    If Not dctKnownSkppdTarget Is Nothing Then dctKnownSkppdTarget.RemoveAll
    
    '~~ Collect Changed Name and/or Scope
    If DueCollect("Changed") Then
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionChanged, 0)
        
        For Each nmeSource In c_wbk_source.Names
            sIdSource = mSyncNames.SyncId(nmeSource)
            If Not KnownSource(sIdSource) Then
                If mSyncNames.Corresponding(c_nme:=nmeSource _
                                           , c_wbk_source:=c_wbk_source _
                                           , c_wbk_target:=c_wbk_target _
                                           , c_name_differs:=bNameDiffers _
                                           , c_scope_differs:=bScopeDiffers _
                                           , c_refto_differs:=bRefToDiffers _
                                           , c_nme_corresponding_target:=nmeTarget _
                                           , c_quality:=enOrNameRef _
                                           ) = "Source-1:Target-1" Then
                    If (bNameDiffers Or bScopeDiffers) And Not bRefToDiffers Then
'                        If bNameDiffers And bScopeDiffers Then Stop
                        CollectChanged nmeSource, nmeTarget, bNameDiffers, bScopeDiffers
                    ElseIf bRefToDiffers Then
                        If Not RefToChangeConfirmed(nmeSource, nmeTarget) Then
                            c_terminated = True
                            GoTo xt
                        Else
                            KnownSkppdSource(sIdSource) = True
                            KnownSkppdTarget(SyncId(nmeTarget)) = True
                            wsSyncLog.Done "change", "Name", SyncId(nmeTarget), "skipped", "Referred Range change from " & nmeTarget.RefersTo & " to " & nmeSource.RefersTo & " skipped"
                       End If
                    End If
                End If
            End If
            Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionChanged, dctKnownChanged.Count / 2)
        Next nmeSource
    End If
    
    '~~ Collect obsolete target Names.
    If DueCollect("Obsolete") Then
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionRemoveObsolete, 0)
        
        For Each nmeTarget In c_wbk_target.Names
            sIdTarget = mSyncNames.SyncId(nmeTarget)
            If Not KnownTarget(sIdTarget) Then
                sCorrespondingName = Corresponding(c_nme:=nmeTarget _
                                               , c_wbk_source:=c_wbk_source _
                                               , c_wbk_target:=c_wbk_target _
                                               , c_dct_source:=dctCorrespSource _
                                               , c_dct_target:=dctCorrespTarget _
                                               , c_name_differs:=bNameDiffers _
                                               , c_refto_differs:=bRefToDiffers _
                                               , c_scope_differs:=bScopeDiffers _
                                               , c_nme_corresponding_target:=nmeCorrespTarget _
                                               , c_nme_corresponding_source:=nmeCorrespSource _
                                               , c_quality:=enAndNameRefScope)
                Select Case sCorrespondingName
                    Case "Source-0:Target-1"
                        mSync.DueSyncLet , enSyncObjectKindName, enSyncActionRemoveObsolete, , sIdTarget
                    Case "Source-1:Target-n" ' the unique corresponding source is about to replace all target Names
                                             ' provided there is no cae of 'design change'
                        For Each vTarget In dctCorrespTarget
                            Set nmeTarget = dctCorrespTarget(vTarget)
                            sIdTarget = mSyncNames.SyncId(nmeTarget)
                            PropertiesDiffer nmeCorrespSource, nmeTarget, bNameDiffers, bRefToDiffers, bScopeDiffers, bNoneDiffers
                            If KnownChanged(sIdTarget) And Not bRefToDiffers Then
                                mSync.DueSyncLet , enSyncObjectKindName, enSyncActionRemoveObsolete, , sIdTarget
                            End If
                        Next vTarget
                End Select
            End If
            Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionRemoveObsolete, dctKnownObsolete.Count)
        Next nmeTarget
    End If
    
    '~~ Collect New Names
    If DueCollect("New") Then
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionAddNew, 0)
        
        For Each nmeSource In c_wbk_source.Names
            sIdSource = mSyncNames.SyncId(nmeSource)
            If Not KnownSource(sIdSource) Then
                sCorrespondingName = mSyncNames.Corresponding(c_nme:=nmeSource _
                                                         , c_wbk_source:=c_wbk_source _
                                                         , c_wbk_target:=c_wbk_target _
                                                         , c_dct_source:=dctCorrespSource _
                                                         , c_dct_target:=dctCorrespTarget _
                                                         , c_nme_corresponding_target:=nmeCorrespTarget _
                                                         , c_nme_corresponding_source:=nmeCorrespSource _
                                                         , c_quality:=enAndNameRefScope)
                Select Case sCorrespondingName
                    Case "Source-1:Target-0"
                        Set wsh = RefersToSheet(nmeSource)
                        sCorrespondingSheet = mSyncSheets.Corresponding(c_wsh:=wsh _
                                                                      , c_wbk:=c_wbk_target _
                                                                      , c_quality:=enOrNameCodeName)
                        If sCorrespondingSheet = "1:0" Then
                            '~~ The new Name belongs to a new (yet not synchronized) Worksheet and thus will be skipped
                            KnownSkppdSource(sIdSource) = True
                        Else
                            mSync.DueSyncLet , enSyncObjectKindName, enSyncActionAddNew, , sIdSource
                        End If
                    Case "Source-1:Target-n" ' the unique corresponding source is about to replace all target Names
                              ' provided there is no cae of 'design change'
                        For Each vTarget In dctCorrespTarget
                            Set nmeTarget = dctCorrespTarget(vTarget)
                            sIdTarget = mSyncNames.SyncId(nmeTarget)
                            If KnownTarget(sIdTarget) Then
                                PropertiesDiffer nmeCorrespSource, nmeTarget, bNameDiffers, bRefToDiffers, bScopeDiffers, bNoneDiffers
                                If bRefToDiffers Then
                                    Exit For
                                End If
                            End If
                        Next vTarget
                        If Not bRefToDiffers Then
                            mSync.DueSyncLet , enSyncObjectKindName, enSyncActionAddNew, , sIdSource
                        End If
                End Select
            End If
            Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionAddNew, dctKnownNew.Count)
        Next nmeSource
    End If
        
    '~~ Collect Multiple (specified as different names refer to the same range)
    If DueCollect("Multiple") Then
        lMultiply = 0
        Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionMultiple, 0)
        
        For Each nmeSource In c_wbk_source.Names
            sIdSource = mSyncNames.SyncId(nmeSource)
            If Not KnownSource(sIdSource) _
            And Not KnownMultiple("source:" & sIdSource) Then
                If Corresponding(c_nme:=nmeSource _
                                , c_wbk_source:=c_wbk_source _
                                , c_wbk_target:=c_wbk_target _
                                , c_quality:=enOrNameRef _
                                , c_dct_source:=dctCorrespSource _
                                , c_dct_target:=dctCorrespTarget _
                                , c_name_differs:=bNameDiffers _
                                , c_refto_differs:=bRefToDiffers _
                                , c_scope_differs:=bScopeDiffers _
                                , c_nme_corresponding_target:=nmeCorrespTarget _
                                , c_nme_corresponding_source:=nmeCorrespSource) = "Source-n:Target-n" Then
                    '~~ There are more than one corresponding target and source Names. Though not constituting a serious
                    '~~ synchronization problem the fact is at least communicated. All target names will be removed and
                    '~~ replaced by the source names provided there is not a 'design change' involved
                    CollectMultiple dctCorrespSource, dctCorrespTarget
                End If
            End If ' name corresponds to more than one target Name object
            
            Services.DsplyStatus mSync.Progress(enSyncObjectKindName, enSyncStepCollecting, enSyncActionMultiple, dctKnownMultiple.Count)
        Next nmeSource
    End If
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CollectChanged(ByVal c_source As Name, _
                           ByVal c_target As Name, _
                           ByVal c_change_name As Boolean, _
                           ByVal c_change_scope As Boolean)
' ------------------------------------------------------------------------
' Returns a Dictionary with those Names in the Sync-Target-Workbook which
' do not have any corresponding Name in the Sync-Target-Workbook and the
' only single corresponding Name in the Sync-Source-Workbook has either
' the Name or the referred Range changed.
' ------------------------------------------------------------------------
    Const PROC = "CollectChanged"
    
    On Error GoTo eh
    Dim enAction    As enSyncAction
    Dim sIdSource   As String:  sIdSource = SyncId(c_source)
    Dim sIdTarget   As String:  sIdTarget = SyncId(c_target)
    
    Select Case True
        Case c_change_name And Not c_change_scope:   enAction = enSyncActionChangeName
        Case Not c_change_name And c_change_scope:   enAction = enSyncActionChangeScope
        Case c_change_name And c_change_scope:       enAction = enSyncActionChangeNameAndScope
    End Select
    
    mSync.DueSyncLet , enSyncObjectKindName, enAction, "from", sIdTarget ' from old = target
    mSync.DueSyncLet , enSyncObjectKindName, enAction, "to", sIdSource   ' to new = source
    KnownChanged(sIdSource) = True
    KnownChanged(sIdTarget) = True
    
xt: Exit Sub

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
        Case mSync.SyncActionIsMultiple(c_action):  Collected = dctKnownMultiple.Count
    End Select
End Function

Private Function CollectInSync(ByVal c_wbk_source As Workbook, _
                               ByVal c_wbk_target As Workbook) As Long
' ----------------------------------------------------------------------------
' Collects the already in sync Names (in dctItemsAlreadyInSync) and returns
' their number.
' ----------------------------------------------------------------------------
    Const PROC = "CollectInSync"
    
    On Error GoTo eh
    Dim bNoneDiffers    As Boolean
    Dim nmeSource       As Name
    Dim sIdSource       As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dctKnownInSync = Nothing
    Set dctKnownInSync = New Dictionary
    
    For Each nmeSource In c_wbk_source.Names
        sIdSource = mSyncNames.SyncId(nmeSource)
        If Corresponding(c_nme:=nmeSource _
                        , c_wbk_source:=c_wbk_source _
                        , c_wbk_target:=c_wbk_target _
                        , c_none_differs:=bNoneDiffers _
                        , c_quality:=enAndNameRefScope _
                        ) = "Source-1:Target-1" Then
            If Not KnownInSync(sIdSource) And bNoneDiffers Then
                KnownInSync(sIdSource) = True
            End If
        End If
    Next nmeSource

xt: CollectInSync = dctKnownInSync.Count
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub CollectKnown(ByRef c_dct As Dictionary, _
                         ByVal c_id As String, _
                Optional ByVal c_v As Variant = Nothing)
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    If c_dct Is Nothing Then Set c_dct = New Dictionary
    If Not c_dct.Exists(c_id) Then
        If VarType(c_v) = vbObject Then
            c_dct.Add c_id, TypeName(c_v)
        Else
            c_dct.Add c_id, c_v
        End If
    End If
End Sub

Private Sub CollectMultiple(ByVal dct_corresp_source As Dictionary, _
                            ByVal dct_corresp_target As Dictionary)
' ------------------------------------------------------------------------
' Collects in the Dictionary mSync.dct a Collection of all source and
' target Names which correspond with each other dct_corresp_source and
' dct_corresp_target).
'
' Note1: Multiple are Name objects which refer to the same range by
'        different names either in the Sync-Target- and/or Sync-Source-
'        Workbook.
'        corresponding Names (i.e. with equal Name and or RefersTo) in
'        the Sync-Source-Workbook and vice versa.
' Note2: Multiple Names are synchronized separated for providing a best
'        objects are exempted from any synchronization but
'        displayed by a red warning that synchronization needs to be
'        discontinued until the problem is solved.
' ------------------------------------------------------------------------
    Const PROC = "CollectMultiple"
    
    On Error GoTo eh
    Dim sId As String
    Dim v   As Variant
    Dim nme As Name
    
    lMultiply = lMultiply + 1
    For Each v In dct_corresp_target
        Set nme = dct_corresp_target(v)
        sId = SyncId(nme)
        If Not KnownMultiple("target:" & sId) Then
            mSync.DueSyncLet , enSyncObjectKindName, enSyncActionMultipleTarget, , sId, "will be removed"
            KnownMultiple("target:" & sId) = True
        End If
    Next v
    
    For Each v In dct_corresp_source
        Set nme = dct_corresp_source(v)
        sId = SyncId(nme)
        If Not KnownMultiple("source:" & sId) Then
            mSync.DueSyncLet , enSyncObjectKindName, enSyncActionMultipleSource, , sId, "will be added"
            KnownMultiple("source:" & sId) = True
        End If
    Next v
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Corresponding(ByVal c_nme As Name, _
                               ByVal c_quality As enCorrespondingNamesQuality, _
                      Optional ByVal c_wbk_source As Workbook = Nothing, _
                      Optional ByVal c_wbk_target As Workbook = Nothing, _
                      Optional ByRef c_dct_source As Dictionary, _
                      Optional ByRef c_dct_target As Dictionary, _
                      Optional ByRef c_name_differs As Boolean, _
                      Optional ByRef c_refto_differs As Boolean, _
                      Optional ByRef c_scope_differs As Boolean, _
                      Optional ByRef c_none_differs As Boolean, _
                      Optional ByRef c_nme_corresponding_target As Name, _
                      Optional ByRef c_nme_corresponding_source As Name) As String
' ------------------------------------------------------------------------
' Returns all Names in the Sync-Source-Workbook (c_wbk_source) - when
' provided - and in the Sync-Target-Workbook (c_wbk_target) - when
' provided - which correspond in accordance with the requested quality
' (c_quality).
' When the result is expressed as a string Source-n:Target-n. When the
' result is corresponding result is Source-1:Target-1 all property
' differencies are returned.
' When the mode (c_quality) is enWeak a different scope still corresponds.
' When the mode is enStrong correspondence is equal Name, RefersTo and
' Scope.
' The syntax of the returned string is tn:m, whereby n = target- and m =
' source correspondencies.
' Note: While it is possible to provide a range with different names it is
'       not pssible to have the same name referring different ranges. This
'       is impossible neither by using the NameManager nor via VBA because
'       a subsequently added Name with the same name ends up with the
'       previous added Name object.
' ------------------------------------------------------------------------
    Const PROC = "Corresponding"
    
    On Error GoTo eh
    Dim sSource As String
    Dim sTarget As String
    Dim sRelate As String
    Dim lTarget As Long
    Dim lSource As Long
    
    
    mBasic.BoP ErrSrc(PROC)
    If Not c_wbk_source Is Nothing Then
        sSource = "Source-"
        mSyncNames.CorrespondingNames c_nme:=c_nme _
                                    , c_in:=c_wbk_source _
                                    , c_quality:=c_quality _
                                    , c_dct_corresponding:=c_dct_source
        lSource = c_dct_source.Count
    End If
    If Not c_wbk_target Is Nothing Then
        sTarget = "Target-"
        mSyncNames.CorrespondingNames c_nme:=c_nme _
                                    , c_in:=c_wbk_target _
                                    , c_quality:=c_quality _
                                    , c_dct_corresponding:=c_dct_target
        lTarget = c_dct_target.Count
    End If
    If sSource <> vbNullString And sTarget <> vbNullString Then sRelate = ":"
    
    If Not c_wbk_source Is Nothing Then
        Select Case lSource
            Case 0, 1: sSource = sSource & lSource
            Case Else: sSource = sSource & "n"
        End Select
        If lSource = 1 Then Set c_nme_corresponding_source = c_dct_source.Items()(0)
    End If
    If Not c_wbk_target Is Nothing Then
        Select Case lTarget
            Case 0, 1: sTarget = sTarget & lTarget
            Case Else: sTarget = sTarget & "n"
        End Select
        If lTarget = 1 Then Set c_nme_corresponding_target = c_dct_target.Items()(0)
    End If
        
xt: Corresponding = sSource & sRelate & sTarget
    If lTarget = 1 And lSource = 1 Then
        PropertiesDiffer c_nme_corresponding_source, c_nme_corresponding_target, c_name_differs, c_refto_differs, c_scope_differs, c_none_differs
    End If
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CorrespondingNames(ByVal c_nme As Name, _
                                    ByVal c_in As Variant, _
                                    ByVal c_quality As enCorrespondingNamesQuality, _
                           Optional ByRef c_dct_corresponding As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all Name objects in the provided Workbook (c_wbk)
' With a name equal to the provided Name object's (c_nme) name. Name objects
' with an invalid Name or invalid RefersTo property are ignored.
' - In weak mode (c_quality) corresponding are all with the same name but with
'   different RefersTo and or different scope.
' - In strong mode only same name, same RefersTo and same Scope corresponds.
' Note: The correspondence has to consider yet not synchronized Worksheets
'       which means that the corresponding target sheet name may differ from
'       the source sheet name but when the CodeNames are equal are still
'       corresponding.
' ----------------------------------------------------------------------------
    Const PROC = "CorrespondingNames"
    
    On Error GoTo eh
    Dim nme             As Name
    Dim dct             As New Dictionary
    Dim sId             As String
    Dim sIdWbk          As String
    Dim bNameDiffers    As Boolean
    Dim bRefToDiffers   As Boolean
    Dim bScopeDiffers   As Boolean
    Dim bNoneDiffers    As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    sIdWbk = SyncId(c_nme)
    For Each nme In c_in.Names
        sId = SyncId(nme)
        
        If InStr(nme.Name, "#") = 0 And InStr(c_nme.Name, "#") = 0 Then
            PropertiesDiffer c_nme, nme, bNameDiffers, bRefToDiffers, bScopeDiffers, bNoneDiffers
            If c_quality = enOrNameRef Then
                If bNameDiffers And bRefToDiffers Then
                    '~~ when both are different this is not a corressponding Name
                Else
                    If Not dct.Exists(sId) Then
                        dct.Add sId, nme
                    End If
                End If
            ElseIf c_quality = enAndNameRefScope Then ' all properties are equal
                If Not bNameDiffers And Not bRefToDiffers And Not bScopeDiffers Then
                    If Not dct.Exists(sId) Then
                        dct.Add sId, nme
                    End If
                End If
            End If
        End If
    Next nme
    
    Set CorrespondingNames = dct
    Set c_dct_corresponding = dct
    Set dct = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CorrespondingRange(ByVal c_nme_source As Name, _
                                    ByVal c_wbk_source As Workbook, _
                                    ByVal c_wbk_target As Workbook) As Range
' ------------------------------------------------------------------------------
' Returns a target range which corresponds to the source range by considering
' that the soure and the target sheet may differ in the name - but will still
' have the same CodeName.
' Note: The function is essential for synchronizing the Names before synchro-
'       nizing the Worksheets.
' ------------------------------------------------------------------------------
    Const PROC = "CorrespondingRange"
    
    On Error GoTo eh
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    
    Set wshSource = c_wbk_source.Worksheets(RefersToSheetName(c_nme_source))
    mSyncSheets.Corresponding c_wsh:=wshSource _
                            , c_wbk:=c_wbk_target _
                            , c_quality:=enOrNameCodeName _
                            , c_wsh_result:=wshTarget
    Set CorrespondingRange = wshTarget.Range(RefersToAddress(c_nme_source))
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CorrespondingScope(ByVal c_nme_source As Name, _
                                    ByVal c_wbk_target As Workbook) As Variant
' ------------------------------------------------------------------------------
' Returns a target range which corresponds to the source range by considering
' that the soure and the target sheet may differ in the name - but will still
' have the same CodeName.
' Note: The function is essential for synchronizing the Names before synchro-
'       nizing the Worksheets.
' ------------------------------------------------------------------------------
    Const PROC = "CorrespondingSope"
    
    On Error GoTo eh
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    
    If ScopesToSheet(c_nme_source, wshSource) Then
        mSyncSheets.Corresponding c_wsh:=wshSource _
                                , c_wbk:=c_wbk_target _
                                , c_quality:=enOrNameCodeName _
                                , c_wsh_result:=wshTarget
        Set CorrespondingScope = wshTarget
    Else
        Set CorrespondingScope = c_wbk_target
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
        Case "Multiple":    DueCollect = dctKnownMultiple Is Nothing
                            If DueCollect Then Set dctKnownMultiple = New Dictionary
    End Select
End Function

Private Function DueNameSync(ByVal d_index As Long, _
                             ByVal d_max_action As Long, _
                             ByVal d_max_direction As Long, _
                             ByVal d_max_kind As Long, _
                             ByVal d_max_name As Long, _
                             ByVal d_max_ref As Long) As String
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    Const PROC = "DueNameSync"
    
    On Error GoTo eh
    Dim sNbsp       As String:  sNbsp = Services.NonBreakingSpace
    Dim sKind       As String:  sKind = cllKind(d_index)
    Dim sAction     As String:  sAction = cllAction(d_index)
    Dim sDirection  As String:  sDirection = cllDirection(d_index)
    Dim sId         As String:  sId = cllId(d_index)
    Dim sComment    As String:  sComment = cllComment(d_index)
    Dim sName       As String:  sName = Split(sId, SYNC_ID_SEPARATOR)(0)
    Dim sRef        As String:  sRef = Split(sId, SYNC_ID_SEPARATOR)(1)
    Dim sScope      As String:  sScope = Split(sId, SYNC_ID_SEPARATOR)(2)
        
    DueNameSync = mBasic.Align(sAction, d_max_action, AlignLeft, , sNbsp)
    DueNameSync = DueNameSync & mBasic.Align(sDirection, d_max_direction, AlignLeft, , sNbsp)
    DueNameSync = DueNameSync & mBasic.Align(sKind, d_max_kind, AlignLeft)
    DueNameSync = DueNameSync & mBasic.Align(sName, d_max_name) & "  " & mBasic.Align(sRef, d_max_ref) & "  " & sScope
    DueNameSync = DueNameSync & sComment
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncNames." & s
End Function

Private Function Exists(ByVal x_nme As Name, _
                        ByVal x_in As Variant, _
               Optional ByRef x_nme_dct_result As Dictionary, _
               Optional ByRef x_nme_result As Name, _
               Optional ByRef x_nme_id As String) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Names (x_nme) name exists in (x_in) - which may be
' a Workbook or a Worksheet - and returns the existing Name as object.
' ------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim sName   As String:  sName = MereName(x_nme)
    Dim nme     As Name
    
    On Error Resume Next
    Set nme = x_in.Names(sName)
    Exists = Err.Number = 0
    
    Select Case mSyncNames.CorrespondingNames(c_nme:=x_nme _
                                            , c_in:=x_in _
                                            , c_quality:=enOrNameRef).Count
        Case 0
            '~~ No Name in Workkbook (x_in) corresponds with the provided Name (x_nme)
            Exit Function
        Case 1
            '~~ A single Name in Workkbook (x_in) corresponds with the provided Name (x_nme)
            Set x_nme_result = x_nme_dct_result.Items()(0)
            x_nme_id = x_nme_dct_result.Keys()(0)
            Exists = True
        Case Else
            '~~ More than one Name in Workkbook (x_in) corresponds with the provided Name (x_nme)
            Exists = True
    End Select
                            
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function GetName(ByVal gn_sync_id As String, _
                         ByVal gn_from_wbk As Workbook, _
                Optional ByRef gn_nme As Name = Nothing) As Name
' ------------------------------------------------------------------------------
' Retuns the Name object - identified by its SyncId (gn_sync_id) - from the
' provided Workbook (gn_from_wbk). When the Name doesn't exist Nothing is
' returned.
' ------------------------------------------------------------------------------
    Const PROC = "GetName"
    
    On Error GoTo eh
    Dim nme As Name
    
    mBasic.BoP ErrSrc(PROC)
    Set gn_nme = Nothing
    Set GetName = Nothing
    For Each nme In gn_from_wbk.Names
        If mSyncNames.SyncId(nme) = gn_sync_id Then
            Set GetName = nme
            Set gn_nme = nme
            Exit For
        End If
    Next nme

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Initialize()
    Set dctMultiple = Nothing
    Set dctMultipleSource = Nothing
    Set dctMultipleTarget = Nothing
    Set dctKnownMultiple = Nothing
    Set dctKnownChanged = Nothing
    Set dctKnownNew = Nothing
    Set dctKnownObsolete = Nothing
    Set dctKnownSkppdSource = Nothing
    Set dctKnownSkppdTarget = Nothing
End Sub

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

Public Sub MaxLenNameId(ByVal ml_wbk_source As Workbook, _
                         ByVal ml_wbk_target As Workbook)
    Const PROC = "MaxLenNameId"
    
    On Error GoTo eh
    Dim nme     As Name
    Dim lMax    As Long
    
    For Each nme In ml_wbk_source.Names
        lMax = Max(lMax, Len(SyncId(nme)))
    Next nme
    For Each nme In ml_wbk_target.Names
        lMax = Max(lMax, Len(SyncId(nme)))
    Next nme
    
xt: wsService.MaxLenNameSyncId = lMax
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ParentWbk(ByVal p_nme As Name) As String
    Dim fso As New FileSystemObject
    If TypeOf p_nme.Parent Is Worksheet _
    Then ParentWbk = fso.GetBaseName(p_nme.Parent.Parent.Name) _
    Else ParentWbk = fso.GetBaseName(p_nme.Parent.Name)
    Set fso = Nothing
End Function

Private Sub PropertiesDiffer(ByVal p_nme_source As Name, _
                             ByVal p_nme_target As Name, _
                    Optional ByRef p_name_differs As Boolean, _
                    Optional ByRef p_range_differs As Boolean, _
                    Optional ByRef p_scope_differs As Boolean, _
                    Optional ByRef p_none_differs As Boolean)
' ----------------------------------------------------------------------------
' Returns TRUE for those Properties which differ (p_name_difffers,
' p_range_differs, p_scope_differs) between the provided Name objects
' (p_nme_source, p_nme_target).
' Note: The service considers target Worksheets yet not synchronized with
'       their corresponding source sheet which means they might still have a
'       different (but fortunately the equal CodeName will uniquely identify
'       the corresponding sheet.
' ----------------------------------------------------------------------------
'    If SyncId(p_nme_source) Like "*NameAndScope*" _
'    And SyncId(p_nme_target) Like "*NameAndScope*" Then Stop
    
    p_name_differs = mNme.MereName(p_nme_source) <> mNme.MereName(p_nme_target)
    p_range_differs = ReferredRangeDiffers(p_nme_source, p_nme_target)
    p_scope_differs = ScopeDiffers(p_nme_source, p_nme_target)
    p_none_differs = Not p_name_differs And Not p_range_differs And Not p_scope_differs
End Sub

Private Function ReferredRangeDiffers(ByVal r_nme_source As Name, _
                                      ByVal r_nme_target As Name, _
                             Optional ByRef r_wsh_source As Worksheet, _
                             Optional ByRef r_wsh_target As Worksheet) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the target Name (r_nme_target) Name RefersTo a different
' address than the source Name (r_nme_source).
' ----------------------------------------------------------------------------
    Set r_wsh_source = RefersToSheet(r_nme_source)
    Set r_wsh_target = RefersToSheet(r_nme_target)
    If r_wsh_source.CodeName = r_wsh_target.CodeName Then
        '~~ Equal CodeNames means that the target and the source Worksheet correspond even when
        '~~ the might have different Names
        ReferredRangeDiffers = RefersToAddress(r_nme_target) <> RefersToAddress(r_nme_source)
    Else
        ReferredRangeDiffers = True
    End If
    
End Function

Private Function RefersToAddress(ByVal r_nme As Name) As String
    RefersToAddress = Split(r_nme.RefersTo, "!")(1)
End Function

Private Function RefersToSheet(ByVal r_nme As Name) As Worksheet
    Set RefersToSheet = RefersToWrkbk(r_nme).Worksheets(RefersToSheetName(r_nme))
End Function

Private Function RefersToSheetName(ByVal r_nme As Name) As String
    RefersToSheetName = Split(Replace(r_nme.RefersTo, "=", vbNullString), "!")(0)
End Function

Private Function RefersToWrkbk(ByVal r_nme As Name) As Workbook
    If TypeOf r_nme.Parent Is Workbook Then Set RefersToWrkbk = r_nme.Parent Else Set RefersToWrkbk = r_nme.Parent.Parent
End Function

Private Function RefToChangeConfirmed(ByVal r_nme_source As Name, _
                                      ByVal r_nme_target As Name) As Boolean
    Dim Msg As mMsg.TypeMsg
    Dim Bttn1           As String
    Dim Bttn2           As String
    Dim sReferredSheet  As String
    
    sReferredSheet = RefersToSheetName(r_nme_source)
    With Msg
        With .Section(1).Text
            .MonoSpaced = True
            .Text = "The referred range of" & vbLf & _
                    "Name: '" & mNme.MereName(r_nme_source) & "' has changed" & vbLf & _
                    "from: '" & r_nme_target.RefersTo & "'" & vbLf & _
                    "to  : '" & r_nme_source.RefersTo & "'."
        End With
        With .Section(2).Text
            .Text = "The above may have two reasons which require either this issue will be skipped or the" & _
                    "synchronization is interrupted to solve this issue before it is continued. " & vbLf & vbLf & _
                    "First: The change/modification of the VB-Project included a design change in the " & _
                    "Worksheet '" & sReferredSheet & "', i.e one or more cells, columns, or rows had been added " & _
                    "or removed. Since Worksheet changes regarding the layout cannot be synchronized this is a " & _
                    "serious issue. When skipped the issue should be addressed manually later." & vbLf & _
                    "Second: While a VB-Project maintenance is underway, in the origin Workbook range (usually a " & _
                    "row) has been added. In this case skip and continue is appropriate."
        End With
        With .Section(3)
            .Label.Text = "Skip and continue:"
            .Label.FontBold = True
            .Text.Text = "Issue will be solved once the synchronization has ended."
        End With
        With .Section(4)
            .Label.Text = "Terminate/interrupt the snchronization:"
            .Label.FontBold = True
            .Text.Text = "Issue will be solved immediately before the synchronization is continued."
        End With
    End With
    With Msg.Section(8).Label
        .Text = "See in README chapter: CompMan's VB-Project-Synchronization" & vbLf & _
                "              section: Name Synchronization"
        .MonoSpaced = True
        .FontColor = rgbBlue
        .OpenWhenClicked = mCompMan.README_URL & mCompMan.README_SYNC_CHAPTER_NAMES
    End With
    
    Bttn1 = "Skip and continue"
    Bttn2 = "Terminate/interrupt the snchronization"
    Select Case mMsg.Dsply(dsply_title:="The referred range of a Name has changed!" _
                         , dsply_msg:=Msg _
                         , dsply_buttons:=mMsg.Buttons(Bttn1, Bttn2))
        Case Bttn1: RefToChangeConfirmed = True
        Case Bttn2: RefToChangeConfirmed = False
    End Select
    
End Function

Public Sub RemoveAllOfSheet(ByVal rn_wsh As Worksheet)
' ----------------------------------------------------------------------------
' Removes all Name objects which either refer to a range of the Worksheet or
' are scoped to the Worksheet.
' Calles when a Worksheet is about to be deleted
' ----------------------------------------------------------------------------
    Const PROC = "RemoveAllOfSheet"
    
    On Error GoTo eh
    Dim wbk As Workbook:    Set wbk = rn_wsh.Parent
    Dim nme As Name
        
    For Each nme In wbk.Names
        If InStr(nme.RefersTo, "=" & rn_wsh.Name & "!") <> 0 _
        Or InStr(nme.Name, rn_wsh.Name & "!") <> 0 Then
            nme.Delete
        End If
    Next nme

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RemoveDuplicatesCreatedBySheetClone(ByVal r_in_dct As Dictionary, _
                                               ByVal r_wbk_source As Workbook, _
                                               ByVal r_wbk_target As Workbook)
' ------------------------------------------------------------------------------
' Delete of all duplicate Names caused by the new cloned/added Worksheet (r_wsh)
' when one has Workbook scope one with the same Name and RefersTo has Worksheet
' scope, the one with worksheet scope.
' ------------------------------------------------------------------------------
    Const PROC = "RemoveDuplicatesCreatedBySheetClone"
    
    On Error GoTo eh
    Dim dctSource       As Dictionary
    Dim dctTarget       As Dictionary
    Dim nme             As Name
    Dim sCorresponding  As String
    Dim nmeTarget       As Name
    Dim nmeSource       As Name
    Dim sId             As String
    
    mBasic.BoP ErrSrc(PROC)
    For Each nme In r_wbk_target.Names
        If Not r_in_dct.Exists(SyncId(nme)) Then
            '~~ This is a Name which resulted from cloning the new Worksheet
            sCorresponding = mSyncNames.Corresponding(c_nme:=nme _
                                                    , c_quality:=enAndNameRefScope _
                                                    , c_wbk_source:=r_wbk_source _
                                                    , c_wbk_target:=r_wbk_target _
                                                    , c_dct_source:=dctSource _
                                                    , c_dct_target:=dctTarget _
                                                    , c_nme_corresponding_source:=nmeSource _
                                                    , c_nme_corresponding_target:=nmeTarget)
            Select Case sCorresponding
                Case "Source-1:Target-1"
                    '~~ This is a Name which exactly matches with the corresponding
                    '~~ source Worksheet's Name
                    sId = SyncId(nmeSource)
                    If KnownSkppdSource(sId) = True Then dctKnownSkppdSource.Remove SyncId(nmeSource)
                    
                Case "Source-0:Target-1"
                    '~~ This is a Name created by a Worksheet clone which does not correspond
                    '~~ with the source Worksheet's Name
                    nmeTarget.Delete
            End Select
        End If
    Next nme

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ScopeDiffers(ByVal s_nme_source As Name, _
                              ByVal s_nme_target As Name, _
                     Optional ByRef s_wsh_source As Worksheet, _
                     Optional ByRef s_wsh_target As Worksheet) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the scope of the target Name (s_nme_target) differs from
' the source Name (s_nme_source) by considering corresponding Worksheets (i.e
' Worksheets havin a different Name but only an equal CodeName because they
' might yet not be synchronized.
' ----------------------------------------------------------------------------
    Select Case True
        Case ScopesToSheet(s_nme_source) And Not ScopesToSheet(s_nme_target)
            ScopeDiffers = True
        Case Not ScopesToSheet(s_nme_source) And ScopesToSheet(s_nme_target)
            ScopeDiffers = True
        Case ScopesToSheet(s_nme_source, s_wsh_source) And ScopesToSheet(s_nme_target, s_wsh_target)
            ScopeDiffers = s_wsh_source.CodeName <> s_wsh_target.CodeName
    End Select
End Function

Private Function ScopeHasChanged(ByVal shc_nme_source As Name, _
                                 ByVal shc_nme_target As Name) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Scope of the source Name object (shc_nme_source)
' differs from the Scope of the traget Name object (shc_nme_target) either in
' the Type or when both scopes are Worksheet in the scoped sheet's Name.
' ----------------------------------------------------------------------------
    Select Case True
        Case TypeName(shc_nme_source.Parent) = "Workbook" _
         And TypeName(shc_nme_target.Parent) = "Workbook"
        
        Case TypeName(shc_nme_source.Parent) = "Worksheet" _
         And TypeName(shc_nme_target.Parent) = "Worksheet" _
         And shc_nme_source.Parent.Name = shc_nme_target.Parent.Name
        
        Case Else: ScopeHasChanged = True
    End Select
End Function

Private Function ScopeName(ByVal s_nme As Name) As String
' ----------------------------------------------------------------------------
' Returns the scope of the provided Name (s_nme) in the form:
' scope: Workbook <workbook-base-name> or
' scope: Worksheet <workbook-base-name>.<worksheet-name>
' ----------------------------------------------------------------------------

    Dim fso As New FileSystemObject
    Dim wbk As Workbook
    Dim wsh As Worksheet
 
    ScopeName = "scope: "
    If TypeOf s_nme.Parent Is Worksheet Then
        Set wsh = s_nme.Parent
        Set wbk = wsh.Parent
        ScopeName = ScopeName & "Worksheet " & fso.GetBaseName(wbk.Name) & "." & wsh.Name
    Else
        Set wbk = s_nme.Parent
        ScopeName = ScopeName & "Workbook " & fso.GetBaseName(wbk.Name)
    End If
    Set fso = Nothing
    
End Function

Private Function ScopesToSheet(ByVal s_nme As Name, _
                      Optional ByRef s_wsh As Worksheet) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the scope of the Name (s_nme) is Worksheet and the scoped
' Worksheet (s_wsh) else FALSE and s_wsh Is Nothing.
' ----------------------------------------------------------------------------
    If TypeOf s_nme.Parent Is Worksheet Then
        ScopesToSheet = True
        Set s_wsh = s_nme.Parent
    End If
End Function

Public Function SyncId(ByVal s_nme As Name) As String
' ----------------------------------------------------------------------------
' Returns a unified id for a Name object (s_nme) in the form:
' <name>-<refersto> when the scope is Workbook or
' <name>-<refersto>-<scopedsheetname> when the scope is Worksheet.
'
' Note: A Worksheet scoped Name object (s_nme) may not conform with the sheet
'       the Name the ReferTo property specifies.
' ----------------------------------------------------------------------------
    Dim fso         As New FileSystemObject
    Dim sMereName   As String
    Dim sScopeName  As String
    Dim v           As Variant
    Dim wsh         As Worksheet
    Dim wbk         As Workbook
    
    sScopeName = "scope: "
    v = Split(s_nme.Name, "!")
    sMereName = v(UBound(v))
    If TypeOf s_nme.Parent Is Worksheet Then
        Set wsh = s_nme.Parent
        Set wbk = wsh.Parent
        sScopeName = sScopeName & "Worksheet " & fso.GetBaseName(wbk.Name) & "." & wsh.Name
    Else
        Set wbk = s_nme.Parent
        sScopeName = sScopeName & "Workbook " & fso.GetBaseName(wbk.Name)
    End If
    
    SyncId = sMereName & SYNC_ID_SEPARATOR & s_nme.RefersTo & SYNC_ID_SEPARATOR & sScopeName
    Set fso = Nothing
    
End Function

Public Sub SyncKind(ByVal s_wbk_source As Workbook, _
                    ByVal s_wbk_target As Workbook)
' ----------------------------------------------------------------------------
' Displays a dialog to perform all required action to synchronize the Names
' of the Sync-Target-Workbook with thos in the Sync-source-Workbook
' ----------------------------------------------------------------------------
    Const PROC                      As String = "SyncKind"
    
    On Error GoTo eh
    Dim AppRunArgs      As New Dictionary
    Dim bDueMultiple    As Boolean
    Dim bDueChange      As Boolean
    Dim bDueNew         As Boolean
    Dim bDueObsolete    As Boolean
    Dim Bttn1           As String
    Dim cllButtons      As New Collection
    Dim fSync           As fMsg
    Dim lSection        As Long
    Dim Msg             As TypeMsg
    Dim sDueSyncs       As String
    Dim bTerminated     As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_NAMES)
    mSync.MonitorStep "Synchronizing Names"
    mSyncNames.Collect s_wbk_source, s_wbk_target, bTerminated
    If bTerminated Then GoTo xt
    
    sDueSyncs = DueSyncs(enSyncObjectKindName)
    With Msg.Section(1)
        .Text.MonoSpaced = True
        .Text.Text = sDueSyncs
    End With
     
    bDueObsolete = InStr(sDueSyncs, enSyncActionRemoveObsolete) <> 0
    bDueNew = InStr(sDueSyncs, SyncActionString(enSyncActionAddNew)) <> 0
    bDueChange = InStr(sDueSyncs, SyncActionString(enSyncActionChangeName)) <> 0 _
              Or InStr(sDueSyncs, SyncActionString(enSyncActionChangeScope)) <> 0
    bDueMultiple = InStr(sDueSyncs, SyncActionString(enSyncActionMultipleTarget)) <> 0 _
              Or InStr(sDueSyncs, SyncActionString(enSyncActionMultipleSource)) <> 0
    
    lSection = 1
    If bDueObsolete Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_REMOVE_OBSOLETE & ":"
            End With
            .Text.Text = "Considered because it has no corresponding source Name, or the target Name will be " & _
                         "replaced by a new source Name."
        End With
    End If
    If bDueNew Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_ADD_NEW & ":"
            End With
            .Text.Text = "Considered new because no corresponding Name exists in the Sync-Target-Workbook, " & _
                         "or because the Name is about to replace more than one target Name."
        End With
    End If
    If bDueChange Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_CHANGE_NAME & ", " & SYNC_ACTION_CHANGE_SCOPE & ":"
            End With
            .Text.Text = "Either the Name or the Scope changed."
        End With
    End If
        
    If bDueMultiple Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_MULTIPLE_SOURCE & ", " & SYNC_ACTION_MULTIPLE_TARGET & ":"
            End With
            .Text.Text = "More than one Name in the Sync-Target- and the Sync-Source-Workbook refer to the same range. " & _
                         "While this does not hinder Name synchronization it should be worth having a look at it since " & _
                         "multiple Names are error prone."
        End With
    End If
        
    Bttn1 = "Perform all synchronization actions" & vbLf & "listed above"
    Set cllButtons = mMsg.Buttons(Bttn1)
    mMsg.BttnAppRun AppRunArgs, Bttn1 _
                                , ThisWorkbook _
                                , "mSyncNames.AppRunSyncAll"
    
    If bDueObsolete Or bDueNew Or bDueChange Then
        lSection = lSection + 1
        With Msg.Section(lSection).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization, section Name Synchronization:"
            .FontColor = rgbBlue
            .OpenWhenClicked = mCompMan.README_URL & mCompMan.README_SYNC_CHAPTER_NAMES
        End With
        
        Application.EnableEvents = True
        '~~ Display the mode-less dialog for the Names synchronization to run
        mMsg.Dsply dsply_title:=TITLE_SYNC_NAMES _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=Services.DialogTop & ";" & Services.DialogLeft
        DoEvents
    End If
     
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SyncProperties(ByVal p_nme_source As Name, _
                           ByVal p_nme_target As Name)
' ------------------------------------------------------------------------
' Synchronizes the properties of the target Name object (p_nme_target)
' with the source Name object (p_nme_source): Name, RefersTo, and Scope.
'
' Note-1: Properties synchronization is only applicable for a target Name
'         object with one single corresponding source Name.
'
' Note-2: The Worksheet Names in the Sync-Target-Workbook are identical
'         with those in the Sync-Source-Workbook, i.e. the Worksheets are
'         in sync.
' ------------------------------------------------------------------------
    Const PROC = "SyncProperties"

    On Error GoTo eh
    mNme.ChangeProperties p_nme_target, p_nme_source

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ValidNames(ByVal vn_wbk As Workbook) As Long
    
    Dim nme As Name
    
    For Each nme In vn_wbk.Names
        If mNme.IsValidUserRangeName(nme) Then ValidNames = ValidNames + 1
    Next nme

End Function

