Attribute VB_Name = "mSyncSheets"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mSyncSheets
'
' Synchronizations provided: New, Obsolete, Name/CodeName Change. When
' all synchronizations are done back-links are eliminated in formulas and
' range names.
'
' Public services:
' - ClearLinksToSource
'   1. Retrieves all links in the provided Workbook (sync_wbk_target) and
'      breaks them, i.e. in all used formulas.
'   2. Retrieves all Names referring to a range in a provided Sync-Source-
'      Workbook and removes it from the ReferredTo property.
' - CollectChangedCodeName
'   Returns a collection of those Worksheets in the Sync-Source-Workbook
'   which do not exist in the Sync-Target-Workbook under the Name but
'   with a different CodeName
' - CollectChanged
'   Returns a collection of those Worksheets in the Sync-Source-Workbook
'   which do not exist in the Sync-Target-Workbook under the CodeName but
'   with a different Name.
' - CollectNew
'   Returns a collection of those Worksheets in the Sync-Source-Workbook
'   which do not exist in the Sync-Target-Workbook under the Name and
'   CodeName.
' - CollectObsolete
'   Returns a collection of those Worksheets in the Sync-Target-Workbook
'   which are regarded obsolete because they do not exist in the Sync-
'   Source-Workbook under its Name and CodeName.
' - Done
'   Returns TRUE when there are no more sheet synchronizations
'   outstanding, i.e. all Collections have 0 items.
' - RunSyncCodeName
'   Called via Application.Run by CommonButton: Changes the CodeName of
'   the provided Sheet in the provided Sync-Target-Workbook to the
'   CodeName of the sheet in the provided Sync-Source-Workbook.
' - RunChange
'   Called via Application.Run by CommonButton: Changes the name of the
'   provided Sheet in the provided Sync-Target-Workbook to the Name of
'   the sheet in the provided Sync-Source-Workbook.
' - RunAdd
'   Called via Application.Run by CommonButton: Copies a provided sheet
'   from a provided Sync-Source-Workbook to a provided Sync-Target-
'   Workbook.
' - RunRemove
'   Called via Application.Run by CommonButton: Removes a provided Sheet
'   from a provided Workbook
' - Sync
'   Called by mSync.RunSync when there are (still) any outstanding Sheet
'   synchronizations to be done. Displays them in a mode-less dialog for
'   being confirmed one by one.
' - SyncOrder
'   Called by mSync.RunSync: Syncronizes the sheet's order in the Sync-
'   Target-Workbook to appear in the same order as in the Sync-Source-
'   Workbook.
'
' W. Rauschenberger Berlin, June 2022
' ------------------------------------------------------------------------

Public Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub ClearLinksToSource(ByVal cls_wsh_source As Worksheet, _
                              ByVal cls_wsh_target As Worksheet)
' -----------------------------------------------------------------------------
' Required when a new Worksheet is copied from the Sync-Source-Workbook to the
' Sync-Target-Workbook.
' 1. Retrieves all links in the provided Workbook (sync_wbk_target) and breaks
'    them, i.e. in all used formulas
' 2. Retrieves all Names referring to a range in the Sync-Source-Workbook
'    (sync_wbk_source) and removes it from the ReferneceTo property
' 3. Clear back-links to the Sync-Source-Workbook (sync_wbk_source) in Shape's
'    OnAction property
' -----------------------------------------------------------------------------
    Const PROC = "ClearLinksToSource"
    
    On Error GoTo eh
    Dim aLinks      As Variant
    Dim v           As Variant
    Dim nme         As Name
    Dim wsh         As Worksheet
    Dim shp         As Shape
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim sRefersTo   As String
    Dim sOnAction   As String
    
    Set wbkTarget = cls_wsh_target.Parent
    Set wbkSource = cls_wsh_source.Parent
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    aLinks = wbkTarget.LinkSources(xlExcelLinks)
    If Not IsEmpty(aLinks) Then
        For Each v In aLinks
            Log.ServicedItem = cls_wsh_target
            wbkTarget.BreakLink v, xlLinkTypeExcelLinks
            DoEvents
            Log.Entry = "(Back)Link to '" & Split(v, "\")(UBound(Split(v, "\"))) & "' cleared"
        Next v
    End If
        
    '~~ Clear Link to source Workbook in Range-Names
    For Each nme In wbkTarget.Names
        If InStr(nme.RefersTo, wbkSource.Name) <> 0 Then
            Log.ServicedItem = cls_wsh_target
            sRefersTo = nme.RefersTo
            nme.RefersTo = Replace(nme.RefersTo, "[" & wbkSource.Name & "]", vbNullString)
            Log.Entry = "Referring back to source removed (RefersTo '" & sRefersTo & "' changed to '" & nme.RefersTo & "')"
        End If
    Next nme
'    wsService.ClearLinksDone = True
    
    '~~ Clear Link to source in any shapes OnAction property
    For Each wsh In wbkTarget.Sheets
        Log.ServicedItem = wsh
        For Each shp In wsh.Shapes
            On Error Resume Next
            sOnAction = shp.OnAction
            If Err.Number = 0 Then ' shape has an OnAction property
                If InStr(sOnAction, wbkSource.Name) <> 0 Then
                    Log.ServicedItem = wsh
                    shp.OnAction = Replace(sOnAction, wbkSource.Name, wbkTarget.Name)
                    Log.Entry = "Referring basck to source component removed (OnAction changed from '" & sOnAction & "' to '" & shp.OnAction & "'"
                End If
            End If
        Next shp
    Next wsh

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CollectAllItems()
' ------------------------------------------------------------------------------
' Writes the potentially synched Worsheets Name to the wsSync sheet.
' ------------------------------------------------------------------------------
    Const PROC = "CollectAllItems"
    
    Dim dct         As New Dictionary
    Dim v           As Variant
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wsh         As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    For Each wsh In wbkSource.Worksheets
        If Not dct.Exists(wsh.Name & wsh.CodeName) _
        Then mDct.DctAdd dct, ItemSyncName(wsh), vbNullString, order_bykey, seq_ascending, sense_casesensitive
    Next wsh
    For Each wsh In wbkTarget.Worksheets
        If Not dct.Exists(wsh.Name & wsh.CodeName) _
        Then mDct.DctAdd dct, ItemSyncName(wsh), vbNullString, order_bykey, seq_ascending, sense_casesensitive
    Next wsh
    
    For Each v In dct
        wsSync.WshItemAll = v
    Next v
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Function ItemSyncName(ByVal wsh As Worksheet) As String
' -----------------------------------------------------------------------------
' Returns a kind of unique sheet id throughout this synchronization.
' -----------------------------------------------------------------------------
    ItemSyncName = wsh.Name & " (" & wsh.CodeName & ") in " & wsh.Parent.Name
End Function

Public Function CollectChanged() As Dictionary
' -----------------------------------------------------------------------------
' Returns a collection of those Worksheets in the Sync-Source-Workbook which do
' not exist in the Sync-Target-Workbook neither under its Name nor its CodeName.
' -----------------------------------------------------------------------------
    Const PROC = "CollectChanged"
    
    On Error GoTo eh
    Dim wshSource   As Worksheet
    Dim dct         As New Dictionary
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim v           As Variant
    Dim wshTarget   As Worksheet
    Dim sChanged    As String
    
    mSync.MonitorStep "Collecting Worksheets the Name or CodeName has changed"
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each wshSource In wbkSource.Sheets
        If HasChanged(wshSource, wbkTarget, wshTarget) Then
            sChanged = "  changed to  " & ItemSyncName(wshSource)
            mDct.DctAdd dct, ItemSyncName(wshTarget), sChanged, order_bykey, seq_ascending, sense_casesensitive
        End If
    Next wshSource
       
    If wsSync.WshNumberChanged = 0 Then
        '~~ Write not yet registered items to wsSync sheet
        For Each v In dct
            wsSync.WshItemChanged(v) = True
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

Public Function CollectNew() As Dictionary
' -----------------------------------------------------------------------------
' Returns a collection of those Worksheets in the Sync-Source-Workbook
' (cn_wbk_source) which do not exist in theSync-Target-Workbook (cn_wbk_target)
' neither under the Name and CodeName.
' -----------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim wshSource   As Worksheet
    Dim dct         As New Dictionary
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim v           As Variant
    
    mSync.MonitorStep "Collecting Worksheets which are new"
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
        
    For Each wshSource In wbkSource.Sheets
        If IsNew(wshSource, wbkTarget) Then
            Log.ServicedItem = wshSource
            mDct.DctAdd dct, ItemSyncName(wshSource), vbNullString, order_bykey, seq_ascending, sense_casesensitive
        End If
    Next wshSource
       
    If wsSync.WshNumberNew = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.WshItemNew(v) = True
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
' Returns a collection of those Worksheets in Sync-Target-Workbook
' (sync_target-wb) which are regarded obsolete because they do not exist in the
' Sync-Source-Workbook (sync_source-wb) under its Name and CodeName.
' ------------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wshTarget   As Worksheet
    Dim v           As Variant
    
    mSync.MonitorStep "Collecting Worksheets which are obsolete"
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each wshTarget In wbkTarget.Sheets
        If IsObsolete(wshTarget, wbkSource) Then
            '~~ The sheet exists in the Sync-Target-Workbook
            '~~ neither under its Name nor under its CodeName
            Log.ServicedItem = wshTarget
            mDct.DctAdd dct, ItemSyncName(wshTarget), vbNullString, order_bykey, seq_ascending, sense_casesensitive
        End If
    Next wshTarget
       
    If wsSync.WshNumberNew = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.WshItemObsolete(v) = True
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
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Neither a Sheet's Name nor a Sheet's CodeName is provided!"
    
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

Public Sub SyncAllSheets()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all Sheets by
' removing obsolete, adding new, and changing the Name/CodeName.
' ------------------------------------------------------------------------------
    Const PROC = "SyncAllSheets"
    
    On Error GoTo eh
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    Dim sSheetName  As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE

    '~~ Synchronize obsolete Worksheets
    For Each wshTarget In wbkTarget.Worksheets
        If IsObsolete(wshTarget, wbkSource) Then
            Log.ServicedItem = wshTarget
            Application.DisplayAlerts = False
            sSheetName = ItemSyncName(wshTarget)
            wshTarget.Delete
            Application.DisplayAlerts = True
            Log.Entry = "Obsolete (deleted)"
            wsSync.WshItemObsoleteDone(sSheetName) = True
        End If
    Next wshTarget

    '~~ Synchronize new and changed Worksheets
    For Each wshSource In wbkSource.Worksheets
        Select Case True
            Case IsNew(wshSource, wbkTarget)
                CopySourceSheetToTargetWorkbook wshSource, wbkTarget, wshTarget
                ClearLinksToSource wshSource, wshTarget
            Case HasChanged(wshSource, wbkTarget, wshTarget)
                SyncSheetsNames wshSource, wshTarget
        End Select
    Next wshSource
            
    mSyncSheets.SyncOrder
    wsSync.WshSyncDone = True
    
    '~~ Re-display the synchronization dialog for still to be synchronized items
    UnloadSyncMessage TITLE_SYNC_SHEETS
    mSync.RunSync

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CopySourceSheetToTargetWorkbook(ByVal sync_wsh_source As Worksheet, _
                                           ByVal sync_wbk_target As Workbook, _
                                           ByRef sync_wsh_target As Worksheet)
' -----------------------------------------------------------------------------
' Copies the sheet (sync_wsh_source) from the Sync-Source-Workbook to the
' Sync-Target-Workbook and returns the target sheet as object.
' -----------------------------------------------------------------------------
    Const PROC = "CopySourceSheetToTargetWorkbook"
    
    On Error GoTo eh
    Dim wbkSource As Workbook
    
    Set wbkSource = sync_wsh_source.Parent
    mService.EstablishServiceLog sync_wbk_target, mCompManClient.SRVC_SYNCHRONIZE
    
    Log.ServicedItem = sync_wsh_source
    sync_wsh_source.Copy After:=sync_wbk_target.Sheets(sync_wbk_target.Worksheets.Count)
    Set sync_wsh_target = sync_wbk_target.Sheets(sync_wbk_target.Worksheets.Count)
    Log.Entry = "Copied to Sync-Target-Workbook"
    wsSync.WshItemNewDone(ItemSyncName(sync_wsh_source)) = True
       
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


















Public Sub SyncSheetsNames(ByVal sn_wsh_source As Worksheet, _
                          ByVal sn_wsh_target As Worksheet)
' -----------------------------------------------------------------------------
' Synchronizes the Name and CodeName of the Sync-TargetWorksheet (sn_wsh_target)
' with the CodeName of the Sync-Source-Worksheet (sn_wsh_source).
' -----------------------------------------------------------------------------
    Const PROC = "SyncSheetsNames"

    On Error GoTo eh
    Dim wbkTarget       As Workbook
    Dim vbcTarget       As VBComponent
    Dim TargetCodeName  As String
    Dim TargetName      As String
    
    Set wbkTarget = sn_wsh_target.Parent
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    If sn_wsh_target.Name <> sn_wsh_source.Name Then
        '~~ Synchronize Worsheet Names
        TargetName = sn_wsh_target.Name
        Log.ServicedItem = sn_wsh_target
        sn_wsh_target.Name = sn_wsh_source.Name
        Log.Entry = "Worksheet Name changed from '" & TargetName & "' to '" & sn_wsh_target.Name & "'"
        wsSync.WshItemChangedDone(ItemSyncName(sn_wsh_source)) = True
    ElseIf sn_wsh_target.CodeName <> sn_wsh_source.CodeName Then
        '~~ Synchronize Worsheet CodeNames
        TargetCodeName = sn_wsh_target.CodeName
        For Each vbcTarget In wbkTarget.VBProject.VBComponents
            If vbcTarget.Name = TargetCodeName Then
                Log.ServicedItem = sn_wsh_target
                vbcTarget.Name = sn_wsh_source.CodeName
                Log.Entry = "CodeName changed from '" & TargetCodeName & "' to '" & vbcTarget.Name & "'"
                wsSync.WshItemChangedDone(ItemSyncName(sn_wsh_source)) = True
                Exit For
            End If
        Next vbcTarget
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function IsNew(ByVal wsh_source As Worksheet, _
                       ByVal wbk_target As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Worksheet (wsh_source) exists in the Workbook (wbk_target)
' neither under the sheet's Name nor its CodeName.
' ------------------------------------------------------------------------------
    Dim wshTarget As Worksheet
    
    IsNew = Not HasChanged(wsh_source, wbk_target, wshTarget) _
            And wshTarget Is Nothing
            
End Function

Private Function IsObsolete(ByVal wsh_target As Worksheet, _
                            ByVal wbk_source As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Worksheet (wsh_target) exists in the Workbook
' (wbk_source) neither under the sheet's Name nor its CodeName.
' ------------------------------------------------------------------------------
    Dim wshSource As Worksheet
    
    IsObsolete = Not HasChanged(wsh_target, wbk_source, wshSource) _
                 And wshSource Is Nothing
End Function

Private Function HasChanged(ByVal hc_wsh As Worksheet, _
                            ByVal hc_wbk As Workbook, _
                   Optional ByRef hc_wsh_result As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' - When the Worksheet (hc_wsh) exists in the Workbook (hc_wbk) either under the
'   Name or the CodeName of the Worksheet (hc_wsh) and neither of the two differs
'   the function returns FALSE and the identified Worksheet (hc_wsh_result)
' - When the Worksheet (hc_wsh) exists in the Workbook (hc_wbk) either under the
'   Name or the CodeName but either of the two is different the function returns
'   TRUE and the identified Worksheet (hc_wsh_result)
' - When no Worksheet had been identified, neither with the same Name nor the
'   same CodeName the function returns FALSE and Nothing in hc_wsh_result.
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet
    
    For Each wsh In hc_wbk.Worksheets
        With wsh
            If .Name = hc_wsh.Name Or .CodeName = hc_wsh.CodeName Then
                '~~ Return the Sheet which may have changed or not
                Set hc_wsh_result = wsh
            End If
            If (.Name = hc_wsh.Name And .CodeName <> hc_wsh.CodeName) _
            Or (.CodeName = hc_wsh.CodeName And .Name <> hc_wsh.Name) Then
                HasChanged = True
                Set hc_wsh_result = wsh
                Exit For
            End If
        End With
    Next wsh

End Function

Public Function CorrespSheet(ByVal sync_wsh As Worksheet, _
                             ByVal sync_wbk As Workbook, _
                    Optional ByRef sync_wsh_result As Worksheet) As Worksheet
' ------------------------------------------------------------------------------
' Returns the Worksheet in the Workbook (sync_wbk) which corresponds to the
' Worksheet (sync_wsh) either by its Name or its CodeName.
' Note: By considering both, the corresponding sheet is identified even when the
'       sheet's names had yet not been synchronized.
' ------------------------------------------------------------------------------
    Const PROC = "CorrespSheet"
    
    On Error GoTo eh
    Dim wsh As Worksheet
    
    For Each wsh In sync_wbk.Worksheets
        If wsh.Name = sync_wsh.Name _
        Or wsh.CodeName = sync_wsh.CodeName _
        Then
            Set sync_wsh_result = wsh
            Set CorrespSheet = wsh
            Exit For
        End If
    Next wsh

xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Sync(ByRef sync_new As Dictionary, _
                ByRef sync_obsolete As Dictionary, _
                ByRef sync_changed As Dictionary)
' ------------------------------------------------------------------------------
' Called by mSync.RunSync provided there are (still) any outstanding Sheet
' synchronizations to be done. Displays them in a mode-less dialog for being
' confirmed either one by one or all together at a time.
' ------------------------------------------------------------------------------
    Const PROC = "Sync"
    
    On Error GoTo eh
    Dim AppRunArgs  As New Dictionary
    Dim cllButtons  As New Collection
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    If sync_obsolete.Count + sync_new.Count + sync_changed.Count = 0 Then GoTo xt
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    '~~ There's at least one Worksheet in need of synchronization
    wsService.SyncDialogTitle = TITLE_SYNC_SHEETS
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_SHEETS)
    With Msg.Section(1)
        .Label.Text = "Obsolete Worksheet(s):"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_obsolete
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(2)
        .Label.Text = "New Worksheet(s):"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_new
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(3)
        .Label.Text = "Worksheet(s) the Name or CodeName has changed:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_changed
            .Text.Text = .Text.Text & vbLf & v & sync_changed(v)
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(4)
        .Label.Text = "About Worksheets synchronization:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Synchronizing Worksheets is one of the tasks of the CompMan's Synchronization service. " & _
                     "This dialog is displayed because there are Worksheets which require synchronization."
    End With
        
    '~~ Prepare a Command-Buttonn with an Application.Run action for the synchronization of all Worksheets
    Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN)
    mMsg.ButtonAppRun AppRunArgs, SYNC_ALL_BTTN _
                                , ThisWorkbook _
                                , "mSyncSheets.SyncAllSheets"
    
    '~~ Display the mode-less dialog for the confirmation which Sheet synchronization to run
    mMsg.Dsply dsply_title:=TITLE_SYNC_SHEETS _
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

Public Sub SyncOrder()
' ----------------------------------------------------------------------------
' Syncronizes the sheet's order in the Sync-Target-Workbook (sync_wbk_target)
' to appear in the same order as in the Sync-Source-Workbook (sync_wbk_source).
' Precondition: All sheet synchronizations had been done.
' ----------------------------------------------------------------------------
    Const PROC = "SyncOrder"
    
    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
   
    For i = 1 To wbkSource.Worksheets.Count
        Set wshSource = wbkSource.Worksheets(i)
        j = i
        With wbkTarget
            Set wshTarget = .Worksheets(j)
            While wshTarget.Name <> wshSource.Name
                .Worksheets(j).Move After:=.Worksheets(.Worksheets.Count)
                DoEvents
                Set wshTarget = .Worksheets(j)
            Wend
        End With
        Log.ServicedItem = wshTarget
        Log.Entry = "Order in sync!"
    Next i
        
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

