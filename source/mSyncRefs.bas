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
Private Const SERVICE_SYNCHRONIZE_REFERENCES    As String = "Synchronize VB-Project References"

Private dctRefsNew                              As New Dictionary
Private dctRefsObsolete                         As New Dictionary

Private Function CollectNew(ByVal cn_wb_target As Workbook, _
                            ByVal cn_wb_source As Workbook) As Collection
' ------------------------------------------------------------------------------
' Returns a Collection of new References. Each new Reference is represented in
' the Collection as Collection with the following structure:
' 1. The buttons caption
' 2. The servicing Workbook (so_wb_servicing)
' 3. The service to run (mSyncRefs.RunRefAdd)
' 4. The Reference objec t to be removed
' 5. The 'Synchronization-Target-Workbook' (so_wb_target)
' 6. The 'Synchronization-Source-workbook' (so_wb_source)
' ------------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim ref     As Reference
    Dim cll1    As Collection
    Dim cll2    As New Collection
        
    For Each ref In cn_wb_source.VBProject.References
        If Not RefExists(cn_wb_target, ref) Then
            Set cll1 = New Collection
            cll1.Add ref.Description & vbLf & vbLf & "Add"   ' 1. The button caption
            cll1.Add ThisWorkbook                            ' 2. The servicing Workbook
            cll1.Add "mSyncRefs.RunRefAdd"                   ' 3. The service to run
            cll1.Add ref                                     ' 4. The Reference to add
            cll1.Add cn_wb_target                            ' 5. The 'Synchronization-Target-Workbook'
            cll1.Add cn_wb_source                            ' 6. The 'Synchronization-Source-Workbook'
            cll2.Add cll1
        End If
    Next ref
        
xt: Set CollectNew = cll2
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CollectObsolete(ByVal so_wb_target As Workbook, _
                                 ByVal so_wb_source As Workbook) As Collection
' ------------------------------------------------------------------------------
' Returns a Collection of obsolete References. Each obsolete Reference is
' represented in the Collection as Collection with the following structure:
' 1. The buttons caption
' 2. The servicing Workbook (so_wb_servicing)
' 3. The service to run (mSyncRefs.RunRefRemove)
' 4. The Reference objec t to be removed
' 5. The 'Synchronization-Target-Workbook' (so_wb_target)
' 6. The 'Synchronization-Source-workbook' (so_wb_source)
' ------------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim ref         As Reference
    Dim cll1        As Collection
    Dim cll2        As New Collection
    
    For Each ref In so_wb_target.VBProject.References
        If Not RefExists(so_wb_source, ref) Then
            Set cll1 = New Collection
            cll1.Add ref.Description & vbLf & vbLf & "Remove"    ' 1. The button's caption
            cll1.Add ThisWorkbook                                ' 2. The servicing Workbook
            cll1.Add "mSyncRefs.RunRefRemove"                    ' 3. The service to run
            cll1.Add ref                                         ' 4. The Reference to remove
            cll1.Add so_wb_target                                ' 5. The 'Synchronization-Target-Workbook'
            cll1.Add so_wb_source                                ' 6. The 'Synchronization-Source-workbook'
            cll2.Add cll1
        End If
    Next ref

xt: Set CollectObsolete = cll2
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRefs." & s
End Function

Public Function RefExists(ByRef re_wb As Workbook, _
                          ByVal re_ref As Variant, _
                 Optional ByRef re_result As Reference = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Reference (re_ref) - which might be a Reference object
' or a string - exists in the VB-Project of the Workbook (re_wb). When a string
' is provided the reference exists when the string is equal to the Name argument
' or it is LIKE the Description argument of any Reference. The existing
' Refwerence is returned as object (re_result).
' ------------------------------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In re_wb.VBProject.References
        If TypeName(re_ref) = "Reference" Then
            RefExists = ref.Name = re_ref.Name
        ElseIf TypeName(re_ref) = "String" Then
            RefExists = ref.Name = re_ref Or ref.Description Like re_ref
        End If
        If RefExists Then Exit Function
    Next ref

End Function

Public Sub RunAll(ByVal sync_wb_target As Workbook, _
               ByVal sync_wb_source As Workbook)
' ------------------------------------------------------------------------------
' Synchronizes all References by removing obsolete and adding new.
' ------------------------------------------------------------------------------
    Const PROC = "RunAll"
    
    On Error GoTo eh
    Dim ref As Reference
    
    If Stats Is Nothing Then Set Stats = New clsStats
    If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = SERVICE_SYNCHRONIZE_REFERENCES
        EstablishTraceLogFile sync_wb_target
    End If
        
    '~~ Remove obsolete References
    With sync_wb_target.VBProject
        For Each ref In .References
            If Not RefExists(sync_wb_source, ref) Then
                Log.ServicedItem = ref
                Stats.Count sic_refs_obsolete
                If dctRefsObsolete.Exists(ref) _
                Then .References.Remove ref
                Log.Entry = "Removed!"
            End If
        Next ref
    End With

    '~~ Ass new References
    With sync_wb_source.VBProject
        For Each ref In .References
            If Not RefExists(sync_wb_target, ref) Then
                Log.ServicedItem = ref
                Stats.Count sic_refs_new
                If dctRefsNew.Exists(ref) _
                Then sync_wb_target.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
                Log.Entry = "Added"
            End If
        Next ref
    End With
        
    '~~ Re-display the synchronization dialog for References still to be synchronized
    mSyncRefs.SyncRefs sync_wb_target, sync_wb_source

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RunRefAdd(ByVal ra_ref As Reference, _
                     ByVal ra_wb_target As Workbook, _
                     ByVal ra_wb_source As Workbook)
' ------------------------------------------------------------------------------
' Adds the Reference (ra_ref) to the Workbook (ra_wb_target).
' The service is exclusively called by through Application.Run invoked by a
' Command-button in a mode-less displayed dialog/message.
' ------------------------------------------------------------------------------
    
    If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = SERVICE_SYNCHRONIZE_REFERENCES
        EstablishTraceLogFile ra_wb_target
    End If
    
    Log.ServicedItem = ra_ref
'    ra_wb_target.VBProject.References.AddFromGuid ra_ref.GUID, ra_ref.Major, ra_ref.Minor
    On Error Resume Next
    ra_wb_target.VBProject.References.AddFromGuid ra_ref.GUID, ra_ref.Major, ra_ref.Minor
'    ra_wb_target.VBProject.References.AddFromFile
    
    If Err.LastDllError = 0 Then Log.Entry = "Added!" Else Log.Entry = "Add failed! (" & Err.Description & ")"
    
    '~~ Re-display the synchronization dialog for References still to be synchronized
    mSyncRefs.SyncRefs ra_wb_target, ra_wb_source

End Sub

Public Sub RunRefRemove(ByVal rr_ref As Reference, _
                     ByVal rr_wb_target As Workbook, _
                     ByVal rr_wb_source As Workbook)
' ------------------------------------------------------------------------------
' Removes the Reference (rr_ref) from Workbook (rr_wb_target).
' The service is exclusively called by through Application.Run invoked by a
' Command-button in a mode-less displayed dialog/message.
' ------------------------------------------------------------------------------
    Dim ref As Reference
    
    If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = SERVICE_SYNCHRONIZE_REFERENCES
        EstablishTraceLogFile rr_wb_target
    End If
    
    With rr_wb_target.VBProject
        For Each ref In .References
            If ref.Name = rr_ref.Name Then
                Log.ServicedItem = ref
                .References.Remove ref
                Log.Entry = "Removed!"
                Exit For
            End If
        Next ref
    End With
    
    '~~ Re-display the synchronization dialog for References still to be synchronized
    mSyncRefs.SyncRefs rr_wb_target, rr_wb_source

End Sub

Public Sub SyncRefs(ByVal sync_wb_target As Workbook, _
                    ByVal sync_wb_source As Workbook)
' ------------------------------------------------------------------------------
' Collects to be synchronized References and displays them in a mode-less dialog
' for being confirmed.
' Called by mSync.RunSync
' ------------------------------------------------------------------------------
    Const PROC          As String = "SyncRefs"
    Const SYNC_ALL_BTTN As String = "Synchronize All"
    
    On Error GoTo eh
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim cllButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    Dim cll         As Collection
    Dim i           As Long
    Dim cllObsolete As New Collection
    Dim cllNew      As New Collection
    
    mMsg.MsgInstance TITLE_SYNC_REFS, True  ' unload any previous displayed dialog
    
    Set cllObsolete = CollectObsolete(sync_wb_target, sync_wb_source)
    Set cllNew = CollectNew(sync_wb_target, sync_wb_source)
    dctRefsObsolete.RemoveAll
    dctRefsNew.RemoveAll
    
    If cllObsolete.Count + cllNew.Count <> 0 Then
        '~~ There's at least one Reference still to be synchronized
        Set fSync = mMsg.MsgInstance(TITLE_SYNC_REFS)
        With Msg.Section(1)
            .Label.Text = "Remove:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Removes the obsolete Reference from the 'Synchronization-Target-Workbook's' VB-Project"
        End With
        With Msg.Section(2)
            .Label.Text = "Add:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Adds the new Reference to the 'Synchronization-Target-Workbook's' VB-Project"
        End With
        With Msg.Section(4)
            .Label.Text = "About:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Synchronizing References is one of the tasks of the CompMan's Synchronization service. " & _
                         "This dialog will be re-displayed until there are no more References to be synchronized. " & _
                         "The one-by-one approach of the service provides full control over what is synchronized."
        End With
        
        If cllObsolete.Count + cllNew.Count = 0 Then GoTo xt
        
        If cllObsolete.Count + cllNew.Count > 1 Then
            '~~ Add an additional 'Synchronize All' button when there is more than one item to be synched
            With Msg.Section(3)
                .Label.Text = SYNC_ALL_BTTN & ":"
                .Label.FontColor = rgbRed
                .Text.Text = "The below named obsolete/new References are synchronized without any " & _
                             "further user interaction. References to be synchronized but not listed " & _
                             "(because of the limited reply button rows of the used mMsg.Dsply service) " & _
                             "will be displayed in a subsequent dialog."
            End With
            Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN, vbLf)
            mMsg.ButtonAppRun AppRunArgs, SYNC_ALL_BTTN, ThisWorkbook, "mSyncRefs.RunAll", sync_wb_target, sync_wb_source
        End If
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for obsolete References for being removed
        For i = 1 To Min(7 - AppRunArgs.Count, cllObsolete.Count)
            Set cll = cllObsolete(i)
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(1), vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(1), cll(2), cll(3), cll(4), cll(5), cll(6)
            dctRefsObsolete.Add cll(4), cll(5)  ' keeps a record of about what 'Synchronize All' means in the specific dialog
        Next i
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for new References for being added
        For i = 1 To Min(7 - AppRunArgs.Count, cllNew.Count)
            Set cll = cllNew(i)
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(1), vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(1), cll(2), cll(3), cll(4), cll(5), cll(6)
            dctRefsNew.Add cll(4), cll(5) ' keeps a record of about what 'Synchronize All' means in the specific dialog
        Next i
        
        '~~ Display the mode-less dialog for the confirmation which Reference
        '~~ to be synchronized which may be one by one or all at once
        If cllButtons(cllButtons.Count) = vbLf Then cllButtons.Remove cllButtons.Count
        
        Application.EnableEvents = True
        sync_wb_source.Activate
        Range("A1").Select
        DoEvents
        
        mMsg.Dsply dsply_title:=TITLE_SYNC_REFS _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs
        
        sync_wb_source.Activate
        Range("A1").Select
        DoEvents
    
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

