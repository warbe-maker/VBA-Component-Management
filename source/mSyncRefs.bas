Attribute VB_Name = "mSyncRefs"
Option Explicit
' ------------------------------------------------------------------------------
' Standard module mSyncRefs
'
' Synchronization of References new or obsolete. The to be synchronized
' Regerences (new and/or obsolete) are presented in a mode-less dialog. When
' more the one is to be synchronized an 'Synchronize All' button is displayed
' which synchronized all the displayed References (by omitting those which may
' additionally need to by synchronized but are not displayed due to the
' limit of the mMsg.Dsply service which allow only a maximum of 7 button rows.
'
' W. Rauschenberger, Berlin June 2022
' ------------------------------------------------------------------------------

Public Sub CollectAllItems()
' ------------------------------------------------------------------------------
' Writes the References potentially synchronizedto the wsSynch sheet.
' ------------------------------------------------------------------------------
    Const PROC = "CollectAllItems"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim ref         As Reference
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    
    For Each ref In mSync.Source.VBProject.References
        If Not dct.Exists(ref.Description) _
        Then mDct.DctAdd dct, ref.Description, ref, order_bykey, seq_ascending, sense_casesensitive
    Next ref
    For Each ref In mSync.TargetCopy.VBProject.References
        If Not dct.Exists(ref.Description) _
        Then mDct.DctAdd dct, ref.Description, ref, order_bykey, seq_ascending, sense_casesensitive
    Next ref

    For Each v In dct
        wsSync.RefItemAll(v) = True
    Next v
    Set dct = Nothing

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function CollectNew() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Collection of new References. Each new Reference is represented in
' the Collection as Collection with the following structure:
' 1. The buttons caption
' 2. The servicing Workbook ( sync_wbk_servicing)
' 3. The service to run (mSyncRefs.RunAdd)
' 4. The Reference objec t to be removed
' ------------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim ref         As Reference
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim v           As Variant
    
    For Each ref In mSync.Source.VBProject.References
        If Not mRef.Exists(mSync.TargetCopy, ref) Then
            Log.ServicedItem = ref
            Set cll = New Collection
            cll.Add ref.Description & vbLf & vbLf & "Add"  ' 1. The button caption
            cll.Add ThisWorkbook                           ' 2. The servicing Workbook
            cll.Add "mSyncRefs.RunAdd"                     ' 3. The service to run
            cll.Add ref                                    ' 4. The Reference to add
            mDct.DctAdd dct, ref.Description, cll, order_bykey, seq_ascending, sense_casesensitive
            Set cll = Nothing
        End If
    Next ref
        
    If wsSync.RefNumberNew = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.RefItemNew(v) = True
        Next v
    End If

xt: Set CollectNew = dct
    Set dct = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function CollectObsolete() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Collection of obsolete References. Each obsolete Reference is
' represented in the Collection as Collection with the following structure:
' 1. The buttons caption
' 2. The servicing Workbook (sync_wbk_servicing)
' 3. The service to run (mSyncRefs.RunRemove)
' 4. The Reference objec t to be removed
' ------------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim ref         As Reference
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim v           As Variant
    
    For Each ref In mSync.TargetCopy.VBProject.References
        If Not mRef.Exists(mSync.Source, ref) Then
            Log.ServicedItem = ref
            Set cll = New Collection
            cll.Add ref.Description & vbLf & vbLf & "Remove"   ' 1. The button's caption
            cll.Add ThisWorkbook                               ' 2. The servicing Workbook
            cll.Add "mSyncRefs.RunRemove"                      ' 3. The service to run
            cll.Add ref                                        ' 4. The Reference to remove
            mDct.DctAdd dct, ref.Description, cll, order_bykey, seq_ascending, sense_casesensitive
            Set cll = Nothing
        End If
    Next ref

    If wsSync.RefNumberObsolete = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.RefItemObsolete(v) = True
        Next v
    End If

xt: Set CollectObsolete = dct
    Set dct = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Done(ByRef sync_new As Dictionary, _
                     ByRef sync_obsolete As Dictionary) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when there are no more sheet synchronizations outstanding plus
' the collections of the outstanding items.
' ----------------------------------------------------------------------------
    Set sync_new = mSyncRefs.CollectNew
    Set sync_obsolete = mSyncRefs.CollectObsolete
    
    Done = sync_new.Count _
         + sync_obsolete.Count = 0
    If Done Then mMsg.MsgInstance TITLE_SYNC_REFS, True
    
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRefs." & s
End Function

Public Sub RunAdd(ByVal sync_ref As Reference)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Adds the Reference (sync_ref) to
' the provided Workbook (sync_wbk_target).
' ------------------------------------------------------------------------------
    
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    Log.ServicedItem = sync_ref
    
    On Error Resume Next
    mSync.TargetCopy.VBProject.References.AddFromGuid sync_ref.GUID, sync_ref.Major, sync_ref.Minor
    If Err.Number = 0 Then
        Log.Entry = "Added!"
        wsSync.RefItemNewDone(sync_ref.Description) = True
    Else
        Log.Entry = "Add failed! (" & Err.Description & ")"
        wsSync.RefItemNewFailed(sync_ref.Description) = True
    End If
    
    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    mSync.RunSync

End Sub

Public Sub SyncAllReferences()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all References by
' removing obsolete and adding new.
' ------------------------------------------------------------------------------
    Const PROC = "SyncAllReferences"
    
    On Error GoTo eh
    Dim ref As Reference
    Dim RefDesc     As String
    
    mBasic.BoP ErrSrc(PROC)
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    
    '~~ Remove obsolete References
    With mSync.TargetCopy.VBProject
        For Each ref In .References
            If Not mRef.Exists(mSync.Source, ref) Then
                Log.ServicedItem = ref
                RefDesc = ref.Description
                On Error Resume Next
                .References.Remove ref
                If Err.Number = 0 Then
                    Log.Entry = "Removed!"
                    wsSync.RefItemObsoleteDone(RefDesc) = True
                Else
                    Log.Entry = "Remove failed! (" & Err.Description & ")"
                    wsSync.RefItemObsoleteFailed(RefDesc) = True
                End If
            End If
        Next ref
    End With

    '~~ Add new References
    With mSync.Source.VBProject
        For Each ref In .References
            If Not mRef.Exists(mSync.TargetCopy, ref) Then
                Log.ServicedItem = ref
                On Error Resume Next
                mSync.TargetCopy.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
                If Err.Number = 0 Then
                    Log.Entry = "Added"
                    wsSync.RefItemNewDone(ref.Description) = True
                Else
                    Log.Entry = "Add failed! (" & Err.Description & ")"
                    wsSync.RefItemNewFailed(ref.Description) = True
                End If
            End If
        Next ref
    End With
    wsSync.RefSyncDone = True
    
    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    UnloadSyncMessage TITLE_SYNC_REFS
    mSync.RunSync

xt: mBasic.BoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RunRemove(ByVal sync_ref As Reference)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes the Reference (rr_ref)
' from the provided Workbook (rr_wbk_target).
' ------------------------------------------------------------------------------
    Dim ref As Reference
    Dim RefDesc     As String
    
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    
    With mSync.TargetCopy.VBProject
        For Each ref In .References
            If ref.Name = sync_ref.Name Then
                Log.ServicedItem = ref
                On Error Resume Next
                RefDesc = ref.Description
                .References.Remove ref
                If Err.Number = 0 Then
                    Log.Entry = "Removed!"
                    wsSync.RefItemObsoleteDone(RefDesc) = True
                Else
                    Log.Entry = "Remove failed!"
                    wsSync.RefItemObsoleteFailed(RefDesc) = True
                End If
                Exit For
            End If
        Next ref
    End With
    
    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    mSync.RunSync

End Sub

Public Sub Sync(ByRef sync_new As Dictionary, _
                ByRef sync_obsolete As Dictionary)
' ------------------------------------------------------------------------------
' Called by mSync.RunSync when there are still outstanding References to be
' synchronized.
' ------------------------------------------------------------------------------
    Const PROC          As String = "Sync"
    
    On Error GoTo eh
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim cllButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    If sync_obsolete.Count + sync_new.Count = 0 Then GoTo xt

    wsService.SyncDialogTitle = TITLE_SYNC_REFS
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_REFS)
    With Msg.Section(1)
        .Label.Text = "Reference(s) obsolete:"
        .Label.FontColor = rgbDarkGreen
        .Text.MonoSpaced = True
        For Each v In sync_obsolete
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(2)
        .Label.Text = "Reference(s) new:"
        .Label.FontColor = rgbDarkGreen
        .Text.MonoSpaced = True
        For Each v In sync_new
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(3)
        .Label.Text = "About:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Synchronizing References is one of the tasks of the CompMan's Synchronization service. " & _
                     "This dialog is displayed because there is at least one Reference which requires synchronization."
    End With
    With Msg.Section(5)
        With .Label
            .Text = "See: Using the Synchronization Service"
            .FontColor = rgbBlue
            .OpenWhenClicked = mCompMan.README_URL & mSync.README_SYNC_CHAPTER
        End With
        .Text.Text = "The chapter 'Using the Synchronization Service' will provide additional information"
    End With
        
    '~~ Prepare a Command-Buttonn with an Application.Run action for the synchronization of all References
    Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN)
    mMsg.ButtonAppRun AppRunArgs, SYNC_ALL_BTTN _
                                , ThisWorkbook _
                                , "mSyncRefs.SyncAllReferences"
        
    '~~ Display the mode-less dialog for the synchronization to run
    mMsg.Dsply dsply_title:=TITLE_SYNC_REFS _
             , dsply_msg:=Msg _
             , dsply_buttons:=cllButtons _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=AppRunArgs _
             , dsply_width_min:=45 _
             , dsply_pos:=wsService.SyncDialogTop & ";" & wsService.SyncDialogLeft
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

