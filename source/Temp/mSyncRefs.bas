Attribute VB_Name = "mSyncRefs"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mSyncRefs: Synchronization of all new or obsolete References.
' ------------------------------------------------------------------------------
Private Const TITLE_SYNC_REFS = "VB-Project Synchronization: References"

Private dctKnownInSync      As Dictionary
Private dctKnownNew         As Dictionary
Private dctKnownObsolete    As Dictionary
                     
Public Property Get KnownNew(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownNew Is Nothing _
    Then KnownNew = dctKnownNew.Exists(nk_id)
End Property

Public Property Let KnownNew(Optional ByVal nk_id As String, _
                                      ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownNew, nk_id
End Property

Public Property Get KnownObsolete(Optional ByVal nk_id As String) As Boolean
    KnownObsolete = dctKnownObsolete.Exists(nk_id)
End Property

Public Property Let KnownObsolete(Optional ByVal nk_id As String, _
                                           ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownObsolete, nk_id
End Property
                     
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRefs." & s
End Function

Public Function MaxLenRefId(ByVal ml_wbk_source As Workbook, _
                            ByVal ml_wbk_target As Workbook) As Long
    Const PROC = "MaxLenRefId"
    
    On Error GoTo eh
    Dim ref     As Reference
    
    For Each ref In mSync.source.VBProject.References
        MaxLenRefId = Max(MaxLenRefId, Len(SyncId(ref)))
    Next ref
    For Each ref In mSync.Target.VBProject.References
        MaxLenRefId = Max(MaxLenRefId, Len(SyncId(ref)))
    Next ref
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function




Public Sub SyncKind(ByVal c_wbk_source As Workbook, _
                    ByVal c_wbk_target As Workbook)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "SyncKind"
    
    On Error GoTo eh
    Dim AppRunArgs      As New Dictionary
    Dim bDueNew         As Boolean
    Dim bDueObsolete    As Boolean
    Dim Bttn1           As String
    Dim cllButtons      As Collection
    Dim lSection        As Long
    Dim Msg             As mMsg.TypeMsg
    Dim sDueSyncs       As String
    
    mBasic.BoP ErrSrc(PROC)
    mSyncRefs.Collect c_wbk_source, c_wbk_target
    Bttn1 = "Perform all Reference synchronization actions" & vbLf & "listed above"
    
    sDueSyncs = mSync.DueSyncs(enSyncObjectKindReference)
    With Msg.Section(1)
        .Text.MonoSpaced = True
        .Text.Text = sDueSyncs
    End With
    bDueObsolete = InStr(sDueSyncs, enSyncActionRemoveObsolete) <> 0
    bDueNew = InStr(sDueSyncs, SyncActionString(enSyncActionAddNew)) <> 0

    lSection = 1
    If bDueObsolete Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_REMOVE_OBSOLETE & ":"
            End With
            .Text.Text = "Considered because it has no corresponding source Reference."
        End With
    End If
    If bDueNew Then
        lSection = lSection + 1
        With Msg.Section(lSection)
            With .Label
                .FontBold = True
                .Text = SYNC_ACTION_ADD_NEW & ":"
            End With
            .Text.Text = "Considered new because no corresponding Reference exists in the Sync-Target-Workbook's VBProject."
        End With
    End If
    mMsg.BttnAppRun AppRunArgs, Bttn1 _
                                , ThisWorkbook _
                                , "mSyncRefs.AppRunSyncAll"
    Set cllButtons = mMsg.Buttons(Bttn1)
    
    If bDueObsolete Or bDueNew Then
        With Msg.Section(8).Label
            .Text = "See in README chapter CompMan's VB-Project-Synchronization, section Name Synchronization:"
            .FontColor = rgbBlue
            .OpenWhenClicked = mCompMan.README_URL & mCompMan.README_SYNC_CHAPTER_NAMES
        End With
        
        Application.EnableEvents = True
        '~~ Display the mode-less dialog for the Names synchronization to run
        mMsg.Dsply dsply_title:=TITLE_SYNC_REFS _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=DialogTop & ";" & DialogLeft
        DoEvents
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function SyncId(ByVal s_ref As Reference) As String
    SyncId = s_ref.Description
End Function

Private Function DueCollect(ByVal d_c As String) As Boolean
    Select Case d_c
        Case "New":         DueCollect = dctKnownNew Is Nothing:        If DueCollect Then Set dctKnownNew = New Dictionary
        Case "Obsolete":    DueCollect = dctKnownObsolete Is Nothing:   If DueCollect Then Set dctKnownObsolete = New Dictionary
    End Select
End Function

Public Sub Collect(ByVal c_wbk_source As Workbook, _
                   ByVal c_wbk_target As Workbook)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim ref As Reference
    Dim sId As String
    
    mBasic.BoP ErrSrc(PROC)
    If lSyncMode = SyncByKind Then mSync.InitDueSyncs
    
    '~~ Collect obsolete
    If DueCollect("Obsolete") Then
        mService.DsplyStatus mSync.Progress(enSyncObjectKindReference, enSyncStepCollecting, enSyncActionRemoveObsolete, 0)
        For Each ref In c_wbk_target.VBProject.References
            sId = SyncId(ref)
            If Not KnownInSync(sId) Then
                If Not mRef.Exists(ref, c_wbk_source) Then
                    mSync.DueSyncLet , enSyncObjectKindReference, enSyncActionRemoveObsolete, , sId
                End If
            End If
            mService.DsplyStatus mSync.Progress(enSyncObjectKindReference, enSyncStepCollecting, enSyncActionRemoveObsolete, dctKnownObsolete.Count)
        Next ref
    End If
    
    '~~ Collect new
    If DueCollect("New") Then
        mService.DsplyStatus mSync.Progress(enSyncObjectKindReference, enSyncStepCollecting, enSyncActionAddNew, 0)
        For Each ref In c_wbk_source.VBProject.References
            sId = SyncId(ref)
            If Not KnownInSync(sId) Then
                If Not mRef.Exists(ref, c_wbk_target) Then
                    mSync.DueSyncLet , enSyncObjectKindReference, enSyncActionAddNew, , sId
                End If
            End If
            mService.DsplyStatus mSync.Progress(enSyncObjectKindReference, enSyncStepCollecting, enSyncActionAddNew, dctKnownNew.Count)
        Next ref
    End If
                   
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AppRunSyncAll()
' ------------------------------------------------------------------------------
' Silent References synchronization.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunSyncAll"
    
    On Error GoTo eh
    Dim ref As Reference
    Dim RefDesc     As String
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source

    '~~ Remove obsolete References
    With mSync.TargetWorkingCopy.VBProject
        For Each ref In .References
            If Not mRef.Exists(ref, wbkSource) Then
                mService.ServicedItem = ref
                RefDesc = ref.Description
                On Error Resume Next
                .References.Remove ref
                If Err.Number = 0 Then
                    wsSyncLog.Done "obsolete", "Reference", RefDesc, "removed", "Removed!"
                Else
                    wsSyncLog.Done "obsolete", "Reference", RefDesc, "remove failed!", "Remove failed! (" & Err.Description & ")"
                End If
            End If
        Next ref
    End With

    '~~ Add new References
    With mSync.source.VBProject
        For Each ref In .References
            If Not mRef.Exists(mSync.TargetWorkingCopy, ref) Then
                On Error Resume Next
                mSync.TargetWorkingCopy.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
                If Err.Number = 0 Then
                    wsSyncLog.Done "new", "Reference", ref.Description, "added", "Added"
                Else
                    wsSyncLog.Done "new", "Reference", ref.Description, "add failed!", "Add failed! (" & Err.Description & ")"
                End If
            End If
        Next ref
    End With
    
    wsSyncLog.RefsDone = True
    
xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        mService.MessageUnload TITLE_SYNC_REFS
        mSync.RunSync
    End If
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


Private Property Get KnownInSync(Optional ByVal nis_id As String) As Boolean
    If Not dctKnownInSync Is Nothing _
    Then KnownInSync = dctKnownInSync.Exists(nis_id)
End Property

Private Property Let KnownInSync(Optional ByVal nis_id As String, _
                                         ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownInSync, nis_id
End Property

Private Function Exists(ByVal x_ref As Variant, _
                        ByVal x_wbk As Workbook, _
               Optional ByRef x_result As Reference = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Reference (x_ref) - which might be a Reference object
' or a string - exists in the VB-Project of the Workbook (x_wbk). When a string
' is provided the reference exists when the string is equal to the Name argument
' or it is LIKE the Description argument of any Reference. The existing
' Refwerence is returned as object (x_ref_result).
' ------------------------------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In x_wbk.VBProject.References
        If TypeName(x_ref) = "Reference" Then
            Exists = ref.Name = x_ref.Name
        ElseIf TypeName(x_ref) = "String" Then
            Exists = SyncId(ref) = x_ref Or ref.Description Like x_ref
        End If
        If Exists Then
            Set x_result = ref
            Exit Function
        End If
    Next ref

End Function

Private Function CollectInSync(ByVal c_wbk_source As Workbook, _
                               ByVal c_wbk_target As Workbook) As Long
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "CollectInSync"
    
    On Error GoTo eh
    Dim refSource   As Reference
    Dim refTarget   As Reference
    Dim sId         As String

    mBasic.BoP ErrSrc(PROC)
    If dctKnownInSync Is Nothing Then Set dctKnownInSync = New Dictionary
    For Each refSource In c_wbk_source.VBProject.References
        sId = SyncId(refSource)
        If Not KnownInSync(sId) Then
            If Exists(refSource, c_wbk_target, refTarget) Then
                KnownInSync(sId) = True
            End If
        End If
    Next refSource
                          
xt: CollectInSync = dctKnownInSync.Count
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Collected(ByVal c_action As enSyncAction) As Long
    Select Case True
        Case c_action = enSyncActionRemoveObsolete: Collected = dctKnownObsolete.Count
        Case c_action = enSyncActionAddNew:         Collected = dctKnownNew.Count
    End Select
End Function

Public Function AllDone(ByVal d_wbk_source As Workbook, _
                        ByVal d_wbk_target As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when there are no more Reference synchronizations outstanding.
' ----------------------------------------------------------------------------
    Const PROC = "AllDone"
    
    Dim lInSync     As Long
    Dim lSourceRefs As Long
    Dim lTargetRefs As Long
    
    mBasic.BoP ErrSrc(PROC)
    lInSync = CollectInSync(d_wbk_source, d_wbk_target)
    lSourceRefs = d_wbk_source.VBProject.References.Count
    lTargetRefs = d_wbk_target.VBProject.References.Count
    lInSync = CollectInSync(d_wbk_source, d_wbk_target)

    If lInSync = lSourceRefs And lInSync = lTargetRefs Then
        AllDone = True
        mMsg.MsgInstance TITLE_SYNC_REFS, True ' unload the mode-less displayed message window
        wsSyncLog.SummaryDone("Reference") = True
        mSync.DueSyncKindOfObjects.DeQueue , enSyncObjectKindReference
    End If
    mBasic.EoP ErrSrc(PROC)

End Function
