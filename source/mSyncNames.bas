Attribute VB_Name = "mSyncNames"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mSyncNames
'
'
' W. Rauschenberger, Berlin, June 2022
' ------------------------------------------------------------------------

Public Sub CollectAllItems()
' ------------------------------------------------------------------------------
' Writes the Range Names potentially synchronizedto the wsSynch sheet.
' ------------------------------------------------------------------------------
    Const PROC = "CollectAllItems"
    
    Dim dct         As New Dictionary
    Dim nme          As Name
    Dim v           As Variant
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    Dim sProgress   As String
    Dim lMaxLenName As Long
    Dim lMaxLenRef  As Long
    
    mBasic.BoP ErrSrc(PROC)
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    For Each nme In wbkSource.Names
        lMaxLenName = Max(lMaxLenName, Len(nme.Name))
        lMaxLenRef = Max(lMaxLenRef, Len(nme.RefersTo))
    Next nme
    For Each nme In wbkTarget.Names
        lMaxLenName = Max(lMaxLenName, Len(nme.Name))
        lMaxLenRef = Max(lMaxLenRef, Len(nme.RefersTo))
    Next nme
    wsService.MaxLenName = lMaxLenName
    wsService.MaxLenRefTo = lMaxLenRef
    
    
    For Each nme In wbkSource.Names
        mSync.MonitorStep "Collecting Range Names all" & sProgress
        If IsRelevantRangeName(nme) Then
            If Not dct.Exists(ItemSyncName(nme)) Then
                mDct.DctAdd dct, ItemSyncName(nme), nme, order_bykey, seq_ascending, sense_casesensitive
            End If
        End If
        sProgress = sProgress & "."
    Next nme

    For Each nme In wbkTarget.Names
        mSync.MonitorStep "Collecting Range Names all" & sProgress
        lMaxLenName = Max(lMaxLenName, Len(nme.Name))
        lMaxLenRef = Max(lMaxLenRef, Len(nme.RefersTo))
        If IsRelevantRangeName(nme) Then
            If Not dct.Exists(ItemSyncName(nme)) Then
                mDct.DctAdd dct, ItemSyncName(nme), nme, order_bykey, seq_ascending, sense_casesensitive
            End If
        End If
        sProgress = sProgress & "."
    Next nme
    
    For Each v In dct
        wsSync.NmeItemAll(v) = True
    Next v
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Function CollectChanged() As Dictionary
' -----------------------------------------------------------------------------
' Returns a collection of those Names which refer to the same range but with a
' different Name.
' Precondition: The Worksheets had already been synchronized!
' -----------------------------------------------------------------------------
    Const PROC = "CollectChanged"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim nmeSource    As Name
    Dim nmeTarget    As Name
    Dim v           As Variant
    Dim wbkSource    As Workbook
    Dim wbkTarget    As Workbook
    Dim SyncIssue   As String
    Dim sProgress       As String
    
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each nmeSource In wbkSource.Names
        mSync.MonitorStep "Collecting Range Names changed" & sProgress
        If IsRelevantRangeName(nmeSource) Then
            If CollectChangedStillRelevant(nmeSource) Then
                If HasChanged(nmeSource, wbkTarget, nmeTarget) Then
                    SyncIssue = "The Name referring to '" & nmeTarget.RefersTo & "' changed to '" & nmeSource.Name & "'"
                    mDct.DctAdd dct, ItemSyncName(nmeSource), SyncIssue, order_bykey, seq_ascending, sense_casesensitive
                End If
            End If
        End If
        sProgress = sProgress & "."
    Next nmeSource
       
    If wsSync.RngNumberChanged = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.NmeItemChanged(v) = True
        Next v
    End If
    
xt: Set CollectChanged = dct
    Set dct = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CollectChangedStillRelevant(ByVal nme As Name) As Boolean
    With wsSync
        CollectChangedStillRelevant = IsSyncItem(nme) _
                         And Not .NmeItemChangedDone(ItemSyncName(nme)) _
                         And Not .NmeItemChangedFailed(ItemSyncName(nme))
    End With
End Function

Public Function CollectNew() As Dictionary
' ------------------------------------------------------------------------
' Returns a Collection of all Range-Names in the Sync-Source-Workbook
' (cn_wbk_source) which either do not exist in the Sync-Target-Workbook
' (cn_wbk_target) or refer to a different range.
' ------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim dct             As New Dictionary
    Dim nmeSource        As Name
    Dim wbkTarget        As Workbook
    Dim wbkSource        As Workbook
    Dim v               As Variant
    Dim sProgress       As String
    
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each nmeSource In wbkSource.Names
        mSync.MonitorStep "Collecting Range Names new" & sProgress
        If IsRelevantRangeName(nmeSource) Then
            If IsNew(nmeSource, wbkTarget) Then
                Log.ServicedItem = nmeSource
                mDct.DctAdd dct, ItemSyncName(nmeSource), nmeSource.Name, order_bykey, seq_ascending, sense_casesensitive
            End If
        End If
        sProgress = sProgress & "."
    Next nmeSource

    If wsSync.NmeNumberNew = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.NmeItemNew(v) = True
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
' ------------------------------------------------------------------------
' Returns a Collection of all Range-Names in the Sync-Source-Workbook
' (cn_wbk_source) which either do not exist in the Sync-Target-Workbook
' (cn_wbk_target) or refer to a different range.
' ------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim nmeTarget    As Name
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    Dim v           As Variant
    Dim sProgress       As String
    
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each nmeTarget In wbkTarget.Names
        mSync.MonitorStep "Collecting Range obsolete" & sProgress
        If IsRelevantRangeName(nmeTarget) Then
            If IsObsolete(nmeTarget, wbkSource) Then
                Log.ServicedItem = nmeTarget
                mDct.DctAdd dct, ItemSyncName(nmeTarget), nmeTarget.Name, order_bykey, seq_ascending, sense_casesensitive
            End If
        End If
        sProgress = sProgress & "."
    Next nmeTarget

    If wsSync.NmeNumberObsolete = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.NmeItemObsolete(v) = True
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

Public Function CorrespName(ByVal cn_nme_name As String, _
                            ByVal cn_wbk As Workbook, _
                   Optional ByRef cn_nme_result As Name) As Name
' -----------------------------------------------------------------------------
' Returns the Name object of the corresponding Name in the Workbook (cn_wbk).
' When there is no corresponding Name in the Workbook (cn-wb) the function
' returns Nothing.
' -----------------------------------------------------------------------------
    Dim nme As Name
    
    For Each nme In cn_wbk.Names
        If nme.Name = cn_nme_name Then
            Set CorrespName = nme
            Set cn_nme_result = nme
            Exit For
        End If
    Next nme
    
End Function


Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncNames." & s
End Function

Private Function Exists(ByVal ex_nme As Name, _
                        ByVal ex_wbk As Workbook, _
               Optional ByRef ex_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE and the resulting name object (ex_nme_result) when the Range
' Name (ex_nme) exists in the Workbook (ex_wbk) - disregarding any
' difference in the RefersTo argument.
' ------------------------------------------------------------------------
    Dim nme As Name
    
    For Each nme In ex_wbk.Names
        If nme.Name = ex_nme.Name Then
            Exists = True
            Set ex_nme_result = nme
            Exit Function
        End If
    Next nme
                        
End Function

Private Function HasChanged(ByVal crtn_nme As Name, _
                            ByVal crtn_wbk As Workbook, _
                   Optional ByRef crtn_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' - When any Name object in the Workbook (crtn_wbk) is referring to the
'   same Range and this object has an identical Name-Property the function
'   returns the identified Name object (crtn_nme) And FALSE
' - When only one Name object in the Workbook (crtn_wbk) is referring to
'   the same Range but with a different Name property the function returns
'   the identified Name object (crtn_nme) And FALSE.
' - When more then one Name object in the Workbook (crtn_wbk) is referring
'   to the same Range and none has an identical Name property the function
'   returns FALSE and None in (crtn_nme_result) indicating that the Name
'   object (crtn_nme) is a new when the Workbook (crtn_wbk) is the Sync-
'   Target-Workbook or is an obsolete one, when the Workbook (crtn_wbk) is
'   the Sync-Source-Workbook.
' ------------------------------------------------------------------------
    Const PROC          As String = "HasChanged"
    
    On Error GoTo eh
    Dim nme              As Name
    Dim dct             As New Dictionary
    Dim v               As Variant
    
    '~~ Collect all Names referring to the same Range as the provided Name (crt_nme)
    For Each nme In crtn_wbk.Names
        If IsRelevantRangeName(nme) Then
            If nme.RefersTo = crtn_nme.RefersTo Then
                dct.Add nme, nme.RefersTo
            End If
        End If
    Next nme
    
    '~~ Analyse the collected Name objects
    Select Case True
        Case dct.Count = 0
        Case dct.Count = 1 And dct.Keys()(0).Name <> crtn_nme.Name
            HasChanged = True
            Set crtn_nme_result = dct.Keys()(0)
        Case dct.Count = 1 And dct.Keys()(0).Name = crtn_nme.Name
            HasChanged = False
            Set crtn_nme_result = dct.Keys()(0)
        Case dct.Count > 1
            '~~ There is more than one Name object referring to the same range.
            '~~ When none of them has the Name property equal to the provided
            '~~ Name object (crt_nme) the returned Name object (crtn_nme_result)
            '~~ is set to Nothing, indicating that the Name is either new or obsolete.
            Set crtn_nme_result = Nothing
            For Each v In dct
                Set nme = v
                If nme.Name = crtn_nme.Name Then
                    Set crtn_nme_result = nme
                    Exit For
                End If
            Next v
    End Select
    
    Set dct = Nothing
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function
                        
Private Function IsNew(ByVal in_nme_source As Name, _
                       ByVal in_wbk_target As Workbook) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Name property of the Name object (in_nme_source)
' is a name relevant for sybchronization and when none of the Name
' objects in Workbook (in-wbk_target) is referring to the same Range as
' the Name (in_nme_source).
' ------------------------------------------------------------------------
    Dim nmeTarget As Name
    
    If IsRelevantRangeName(in_nme_source) Then
        If Not HasChanged(in_nme_source, in_wbk_target, nmeTarget) Then
            IsNew = nmeTarget Is Nothing
        End If
    End If
    
End Function

Private Function IsObsolete(ByVal io_nme_target As Name, _
                            ByVal io_wbk_source As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Name property of the Name object (in_nme_source) is a
' name relevant for synchronization and when none of the Name objects in the
' Workbook (io_wbk) is referring the Range the Name (io_nme) does.
' ----------------------------------------------------------------------------
    Dim nmeSource As Name
    
    If IsRelevantRangeName(io_nme_target) Then
        If Not HasChanged(io_nme_target, io_wbk_source, nmeSource) Then
            IsObsolete = nmeSource Is Nothing
        End If
    End If
    
End Function

Private Function IsSyncItem(ByVal nme As Name) As Boolean
    With wsSync
        IsSyncItem = .NmeNumberAll > 0 _
              And .NmeItemAll(ItemSyncName(nme))
    End With
End Function

Public Function ItemSyncName(ByVal nme As Name) As String
' ----------------------------------------------------------------------------
' Returns a unified Range Name Id.
' ----------------------------------------------------------------------------
    ItemSyncName = Align(nme.Name, wsService.MaxLenName) & " " & nme.RefersTo
End Function

Private Function IsRelevantRangeName(ByVal nme As Name) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the properties of the Name object indicate a name relevant
' for synchronization.
' ----------------------------------------------------------------------------
    IsRelevantRangeName = nme.RefersTo <> vbNullString And nme.Name <> vbNullString
    If IsRelevantRangeName Then _
       IsRelevantRangeName = InStr(nme.RefersTo, "=") <> 0 _
                    And InStr(nme.RefersTo, "!") <> 0 _
                    And InStr(nme.RefersTo, "$") <> 0 _
                    And nme.RefersTo <> "=#NAME?"
End Function

Public Sub SyncAllRangeNames()
' ----------------------------------------------------------------------------
' Called via Application.Run, synchronizes all Range Names.
' ----------------------------------------------------------------------------
    Const PROC = "SyncAllRangeNames"
    
    On Error GoTo eh
    Dim nmeTarget    As Name
    Dim nmeSource    As Name
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    Dim sRangeName  As String
    Dim sProgress   As String
    
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    '~~ 1. Remove obsolete Range Names
    For Each nmeTarget In wbkTarget.Names
        mSync.MonitorStep "Synchronizing Range Names obsolete" & sProgress
        If IsObsolete(nmeTarget, wbkSource) Then
            Log.ServicedItem = nmeTarget
            sRangeName = ItemSyncName(nmeTarget)
            nmeTarget.Delete
            Log.Entry = "Obsolete! Removed from Sync-Target-Workbook"
            wsSync.NmeItemObsoleteDone(sRangeName) = True
        End If
        sProgress = sProgress & "."
    Next nmeTarget
    
    '~~ 2. Add new Range Names
    sProgress = " "
    For Each nmeSource In wbkSource.Names
        mSync.MonitorStep "Synchronizing Range Names new" & sProgress
        If IsNew(nmeSource, wbkTarget) Then
            Log.ServicedItem = nmeSource
            sRangeName = ItemSyncName(nmeSource)
            wbkTarget.Names.Add nmeSource.Name, nmeSource.RefersTo
            Log.Entry = "New! Added to Sync-Target-Workbook"
            wsSync.NmeItemNewDone(sRangeName) = True
        End If
        sProgress = sProgress & "."
    Next nmeSource
    
    '~~ 3. Synchronize changed RefersTo argument
    sProgress = " "
    For Each nmeSource In wbkSource.Names
        mSync.MonitorStep "Synchronizing Range Names changed" & sProgress
        If HasChanged(nmeSource, wbkTarget, nmeTarget) Then
            Log.ServicedItem = nmeSource
            sRangeName = ItemSyncName(nmeSource)
            nmeTarget.Name = nmeSource.Name
            Log.Entry = "New! Added to Sync-Target-Workbook"
            wsSync.NmeItemChangedDone(sRangeName) = True
        End If
        sProgress = sProgress & "."
    Next nmeSource
    
    wsSync.NmeSyncDone = True
    
    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
   UnloadSyncMessage TITLE_SYNC_NAMES
   mSync.RunSync

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Sync(ByRef sync_new As Dictionary, _
                ByRef sync_obsolete As Dictionary, _
                ByVal sync_changed As Dictionary)
' ------------------------------------------------------------------------------
' Collects to be synchronized Names and displays them in a mode-less dialog
' for being confirmed.
' Called by mSync.RunSync
' ------------------------------------------------------------------------------
    Const PROC          As String = "Sync"
    
    On Error GoTo eh
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim cllButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    Dim v           As Variant
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    
    If sync_obsolete.Count + sync_new.Count = 0 Then GoTo xt
    '~~ There's at least one Range Name in need of synchronization
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_NAMES)
    
    With Msg.Section(1)
        .Label.Text = "Remove the obsolete Name(s):"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_obsolete
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(2)
        .Label.Text = "Add the new Name(s):"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_new
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(3)
        .Label.Text = "Ranges refered to by a different Name:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_changed
            .Text.Text = .Text.Text & vbLf & v & sync_changed(v)
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(4)
        .Label.Text = "About synchronizing Range Names:"
        .Label.FontColor = rgbBlue
        .Text.FontColor = rgbRed
        .Text.Text = "Attention! Synchronizing Names is a pretty delicate synchronization tasks. It is " & _
                     "an absolute must that any additional new columns and/or rows had already been " & _
                     "added to the Sync-Target-Workbook - and provided with a specified Range Name! -  " & _
                     "b e f o r e  any VB-Project modification is started (i.e. before the productive " & _
                     "Workbook is copied into a dedicated folder in CompMan's 'Serviced-Folder'. In case " & _
                     "of any doubts the synchronization should be terminated by closing this dialog and " & _
                     "started again once the required preparation (new columns and/or rows added) has " & _
                     "been done - in both, the Sync-Target- and the Sync-Source-Workbook!"
    End With
        
    '~~ Prepare a Command-Buttonn with an Application.Run action for the synchronization of all Worksheets
    Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN)
    mMsg.ButtonAppRun AppRunArgs, SYNC_ALL_BTTN _
                                , ThisWorkbook _
                                , "mSyncNames.SyncAllRangeNames"
    
    '~~ Display the mode-less dialog for the Names synchronization to run
    mMsg.Dsply dsply_title:=TITLE_SYNC_NAMES _
             , dsply_msg:=Msg _
             , dsply_buttons:=cllButtons _
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

