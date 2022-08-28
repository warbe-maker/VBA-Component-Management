Attribute VB_Name = "mSyncComps"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSyncComps
'                 Provides the services to synchronize VBComponents between a
'                 source and a target Workbook/VB-Project.
' Public services:
' - ByCodeLines
' - CollectChanged  Returns a Collection with all changed VBComponents
' - CollectNew      Returns a Collection with all new VBComponents
' - CollectObsolete Returns a Collection with all obsolete VBComponents.
' - Done            Returns TRUE when all the above Collections total items
'                   is 0.
' - CollectAllItems Writes all relevant items of a kind to the wsSyc Worksheet
' - SyncNew         Called via Application.Run by CommonButton: Adds a
'                   VBComponent to the provided Workbook's VBProject.
' - SyncChanged     Called via Application.Run by CommonButton: Synchronizes a
'                   code change in a Sync-Source-Workbook's VBComponent with
'                   the corresponding Sync-Target-Workbook's VBComponent.
' - RunDsplyDiff    Called via Application.Run by CommonButton: Displays the
'                   code changes in a Sync-Source-Workbook's VBComponent
'                   compared with the corresponding Sync-Target-Workbook's
'                   VBComponent.
' - SyncObsolete    Called via Application.Run by CommonButton: Removes
'                   VBComponent from the Sync-Target-Workbook's VBProject.
' - Sync            Called by mSync.RunSync when there are still outstanding
'                   VBComponents to be synchronized.
'
' W. Rauschenberger Berlin June 2022
' ----------------------------------------------------------------------------

Public Function AppErr(ByVal app_err_no As Long) As Long
' ----------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub ByCodeLines(ByVal sync_vbc As VBComponent)
' ----------------------------------------------------------------------------
' Synchronizes the code of a Target-VBComponent (sync_target_vbc_name) in a
' Target-Workbook/VB-Project (Synch.TargetWb) with the code in the Export-File
' of the corresponding Sync-Source-Workbook/VB-Project's
' component line by line.
' ----------------------------------------------------------------------------
    Const PROC = "ByCodeLines"

    On Error GoTo eh
    Dim i           As Long
    Dim SourceComp  As clsComp
    Dim SourceCode  As Dictionary
    Dim v           As Variant
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    Debug.Print ErrSrc(PROC) & " Target: " & wbkTarget.Name
    Debug.Print ErrSrc(PROC) & " Source: " & wbkSource.Name
    
    '~~ Obtain non provided code lines for the line by line syncronization
    Set SourceComp = New clsComp
    With SourceComp
        Set .Wrkbk = wbkSource
        .CompName = sync_vbc.Name
        Set SourceCode = .CodeLines
    End With
    
    With wbkTarget.VBProject.VBComponents(sync_vbc.Name).CodeModule
        If .CountOfLines > 0 _
        Then .DeleteLines 1, .CountOfLines   ' Remove all lines from the cloned raw component
        
        For Each v In SourceCode   ' Insert the raw component's code lines
            i = i + 1
            .InsertLines i, SourceCode(v)
        Next v
    End With
                
xt: Set SourceComp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CollectAllItems()
' ------------------------------------------------------------------------------
' Returns the number of VBComponents involved in synchronization.
' ------------------------------------------------------------------------------
    Const PROC = "CollectAllItems"
    
    Dim dct         As New Dictionary
    Dim vbc         As VBComponent
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    Dim v           As Variant
    Dim lMaxLen     As Long
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    For Each vbc In wbkTarget.VBProject.VBComponents
        lMaxLen = mBasic.Max(lMaxLen, Len(vbc.Name))
    Next vbc
    For Each vbc In wbkSource.VBProject.VBComponents
        lMaxLen = mBasic.Max(lMaxLen, Len(vbc.Name))
    Next vbc
    wsService.MaxLenVbcName = lMaxLen
    
    For Each vbc In wbkTarget.VBProject.VBComponents
        If Not dct.Exists(vbc.Name) Then mDct.DctAdd dct, ItemSyncName(vbc), vbc, , seq_ascending
    Next vbc
    For Each vbc In wbkSource.VBProject.VBComponents
        If Not dct.Exists(vbc.Name) Then mDct.DctAdd dct, ItemSyncName(vbc), vbc, , seq_ascending
    Next vbc
    
    For Each v In dct
        wsSync.VbcItemAll(v) = True
    Next v
       
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Function CollectChanged() As Dictionary
' ------------------------------------------------------------------------------
' Returns a collection of those VBComponents in the Sync-Target-Workbook
' (sync_target_workbook) whith a different code in the corresponding VBComponent
' in the Sync-Source-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "CollectChanged"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim SourceComp  As clsComp
    Dim TargetComp  As clsComp
    Dim v           As Variant
    Dim vbc         As VBComponent
    Dim wbkSource    As Workbook
    Dim wbkTarget    As Workbook
    
    mSync.MonitorStep "Collecting VB-Components the code has changed"
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
       
    If wsSync.VbcNumberChanged = 0 _
    Or (wsSync.VbcNumberChanged > 0 And wsSync.VbcNumberChangedDone + wsSync.VbcNumberChangedFailed < wsSync.VbcNumberChanged) Then
        For Each vbc In wbkSource.VBProject.VBComponents
            Set SourceComp = New clsComp
            With SourceComp
                Set .Wrkbk = wbkSource
                .CompName = vbc.Name
            End With
            If Exists(vbc, wbkTarget) Then
                Set TargetComp = New clsComp
                With TargetComp
                    Set .Wrkbk = wbkTarget
                    .CompName = vbc.Name
                End With
                If CompChanged(SourceComp.CodeLines, TargetComp.CodeLines) Then
                    Log.ServicedItem = vbc
                    Set cll = New Collection
                    cll.Add mComp.TypeString(vbc) & _
                            " " & vbc.Name & _
                            vbLf & vbLf & _
                            "Synchronize code change"   ' 1. The button's caption
                    cll.Add ThisWorkbook                ' 2. The servicing Workbook
                    cll.Add "mSyncComps.SyncChanged"    ' 3. The service to run
                    cll.Add vbc.Name                    ' 4. The VBComponent to update/renew
                    mDct.DctAdd dct, ItemSyncName(vbc), cll, order_bykey, seq_ascending, sense_casesensitive
                    Set cll = Nothing
                End If
                Set TargetComp = Nothing
                Set SourceComp = Nothing
            End If
        Next vbc
    End If
    
    If wsSync.VbcNumberChanged = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.VbcItemChanged(v) = True
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

Private Function ItemSyncName(ByVal vbc As VBComponent) As String
' ------------------------------------------------------------------------------
' Returns a VBComponent-Name used as id throughout VBComponent synchronization.
' ------------------------------------------------------------------------------
    ItemSyncName = mBasic.Align(vbc.Name, wsService.MaxLenVbcName) & " " & mComp.TypeString(vbc)
End Function

Public Function CollectNew() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Collection of all new VBComponents, i.e. VBComponentsd which exist
' in the Sync-Source-Workbook  but not in the Sync-Targe-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim vbc         As VBComponent
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim v           As Variant
    
    mSync.MonitorStep "Collecting VB-Components new"
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    If wsSync.VbcNumberNew = 0 _
    Or (wsSync.VbcNumberNew > 0 And wsSync.VbcNumberNewDone + wsSync.VbcNumberNewFailed < wsSync.VbcNumberNew) Then
        For Each vbc In wbkSource.VBProject.VBComponents
            If vbc.Type <> vbext_ct_Document And vbc.Type <> vbext_ct_ActiveXDesigner Then
                If Not Exists(vbc, wbkTarget) Then
                    Log.ServicedItem = vbc
                    Set cll = New Collection
                    cll.Add mComp.TypeString(vbc) & _
                            " " & vbc.Name & _
                            vbLf & vbLf & _
                            "Add"                   ' 1. The button's caption
                    cll.Add ThisWorkbook            ' 2. The servicing Workbook
                    cll.Add "mSyncComps.SyncNew"    ' 3. The service to run
                    cll.Add vbc.Name                ' 4. The VBComponent to remove
                    mDct.DctAdd dct, ItemSyncName(vbc), cll, order_bykey, seq_ascending, sense_casesensitive
                    Set cll = Nothing
                End If
            End If
        Next vbc
    End If
    
    If wsSync.VbcNumberNew = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.VbcItemNew(v) = True
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
' Collect all obsolete VBComponents, i.e. those wich exist in the Sync-Target-
' Workbook (sync_wbk_target) but not in the Sync-Source-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim CompTarget  As clsComp
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    Dim v           As Variant
    
    mSync.MonitorStep "Collecting VB-Components obsolete"
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    If wsSync.VbcNumberObsolete = 0 _
    Or (wsSync.VbcNumberObsolete > 0 And wsSync.VbcNumberObsoleteDone + wsSync.VbcNumberObsoleteFailed < wsSync.VbcNumberObsolete) Then
        For Each vbc In wbkTarget.VBProject.VBComponents
            If vbc.Type <> vbext_ct_Document Then
                '~~ Datamodules cannot be obsolete but only Worksheets (not the subject here!)
                If Not Exists(vbc, wbkSource) Then
                    Log.ServicedItem = vbc
                    Set cll = New Collection
                    cll.Add mComp.TypeString(vbc) & _
                            " " & vbc.Name & _
                            vbLf & vbLf & _
                            "Remove"                    ' 1. The button's caption
                    cll.Add ThisWorkbook                ' 2. The servicing Workbook
                    cll.Add "mSyncComps.SyncObsolete"   ' 3. The service to run
                    cll.Add vbc.Name                    ' 4. The VBComponent to remove
                    mDct.DctAdd dct, ItemSyncName(vbc), cll, order_bykey, seq_ascending, sense_casesensitive
                    Set cll = Nothing
                End If
                Set CompTarget = Nothing
            End If
        Next vbc
    End If
    
    If wsSync.VbcNumberObsolete = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.VbcItemObsolete(v) = True
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

Private Function CompChanged(ByVal cc_source As Dictionary, _
                             ByVal cc_target As Dictionary) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Dictionaries (source_code, target_code) are different.
' A difference only constituted by a case difference is ignored.
' ----------------------------------------------------------------------------
    CompChanged = mDct.DctDiffers(dd_dct1:=cc_source _
                                , dd_dct2:=cc_target _
                                , dd_diff_items:=True _
                                , dd_diff_keys:=False _
                                , dd_ignore_items_empty:=False _
                                , dd_ignore_case:=True)
End Function

Public Function Done(ByRef sync_new As Dictionary, _
                     ByRef sync_obsolete As Dictionary, _
                     ByRef sync_changed As Dictionary) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when there are no more VBComponent synchronizations outstanding
' Returns collections of the outstanding items.
' ----------------------------------------------------------------------------
    
    Set sync_obsolete = mSyncComps.CollectObsolete
    Set sync_new = mSyncComps.CollectNew
    Set sync_changed = mSyncComps.CollectChanged
    
    Done = (sync_new.Count _
          + sync_obsolete.Count _
          + sync_changed.Count) = 0

    If Done Then mMsg.MsgInstance TITLE_SYNC_VBCOMPS, True ' unload the mode-less displayed message window

End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncComps." & s
End Function

Private Function Exists(ByVal ex_vbc As VBComponent, _
                        ByVal ex_wbk As Workbook, _
               Optional ByRef ex_vbc_result As VBComponent) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE and the VBComponent (ex_vbc_result) when the VBComponent (ex_vbc)
' exists in the Workbook's (ex_wbk) VBProject.
' ------------------------------------------------------------------------------
    Dim vbc As VBComponent
    
    For Each vbc In ex_wbk.VBProject.VBComponents
        If vbc.Name = ex_vbc.Name Then
            Exists = True
            Set ex_vbc_result = vbc
            Exit For
        End If
    Next vbc
    
End Function

                    
Private Function GetVBComp(ByVal get_vbc_name As String, _
                           ByVal get_wbk As Workbook) As VBComponent
' ------------------------------------------------------------------------------
' Returns the VBComponent named (get_vbc_name) frme the Workbook's (get_wbk)
' VBProject.
' ------------------------------------------------------------------------------
    Const PROC = "GetVBComp"
    
    On Error Resume Next
    Set GetVBComp = get_wbk.VBProject.VBComponents(get_vbc_name)
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The VBProject in the Workbook '" & get_wbk.Name & "' has no VBComponent named '" & get_vbc_name & "'!"

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SyncNew(ByVal sync_vbc_name As String)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Adds the VBComponent (sync_vbc)
' to the provided Workbook's (sync_wbk_target) VBProject.
' ------------------------------------------------------------------------------
    Const PROC = "SyncNew"
    
    On Error GoTo eh
    Dim SourceComp  As clsComp
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    Set SourceComp = New clsComp
    With SourceComp
        Set .Wrkbk = wbkSource
        .CompName = sync_vbc_name
        If Not .Exists(wbkTarget) Then
            Log.ServicedItem = .VBComp
            wbkTarget.VBProject.VBComponents.Import .ExpFileFullName
            Log.Entry = "Added (by import of the ExportFile from the corresponding Sync-Source-Workbook's VBComponent"
            wsSync.VbcItemNewDone(ItemSyncName(.VBComp)) = True
        End If
    End With
    Set SourceComp = Nothing
    
    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    UnloadSyncMessage TITLE_SYNC_VBCOMPS
    mSync.RunSync

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncChanged(ByVal sync_vbc_name As String)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes a code change in the
' VBComponent (sync_vbc) in the Sync-Source-Workbook with the
' corresponding VBComponent in the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "SyncChanged"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim SourceComp  As clsComp
    Dim TargetComp  As clsComp
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each vbc In wbkSource.VBProject.VBComponents
        If vbc.Name = sync_vbc_name Then
            Log.ServicedItem = vbc
            Set SourceComp = New clsComp
            With SourceComp
                Set .Wrkbk = wbkSource
                .CompName = vbc.Name
            End With
            Set TargetComp = New clsComp
            With TargetComp
                Set .Wrkbk = wbkTarget
                .CompName = vbc.Name
            End With
                
            Select Case SourceComp.VBComp.Type
                Case vbext_ct_Document
                    mSyncComps.ByCodeLines vbc
                    Log.Entry = "Code lines in CodeModule replaced based on ExportFile"
                    wsSync.VbcItemChangedDone(ItemSyncName(vbc)) = True
                    TargetComp.ExpFile.Delete
                Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                    mRenew.ByImport bi_wbk_serviced:=wbkTarget _
                                  , bi_vbc_name:=vbc.Name _
                                  , bi_exp_file:=SourceComp.ExpFileFullName
                    wsSync.VbcItemChangedDone(ItemSyncName(vbc)) = True
                    TargetComp.ExpFile.Delete
                Case Else
                    Err.Raise AppErr(1), ErrSrc(PROC), "Type '" & SourceComp.TypeString & "' is not supported!"
            End Select
            Set TargetComp = Nothing
            Set SourceComp = Nothing
            Exit For
        End If
    Next vbc

    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    UnloadSyncMessage TITLE_SYNC_VBCOMPS
    mSync.RunSync
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RunDsplyDiff(ByVal sync_vbc_name As String)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Displays the code changes of the
' VBComponent (sync_vbc_name) in the Sync-Source-Workbook compared with the
' corresponding component in the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "RunDsplyDiff"
    
    On Error GoTo eh
    Dim SourceComp  As clsComp
    Dim TargetComp  As clsComp
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    Set SourceComp = New clsComp
    With SourceComp
        Set .Wrkbk = wbkSource
        .CompName = sync_vbc_name
    End With
    Set TargetComp = New clsComp
    With TargetComp
        Set .Wrkbk = wbkTarget
        .CompName = sync_vbc_name
    End With
    mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=TargetComp.ExpFileFullName _
                               , fd_exp_file_right_full_name:=SourceComp.ExpFileFullName _
                               , fd_exp_file_left_title:="Outdated (in-use) Component: '" & TargetComp.ExpFileFullName & "'" _
                               , fd_exp_file_right_title:="Current (modified, up-to-date) Component: '" & SourceComp.ExpFileFullName & "'"
                       
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

                           
Public Sub SyncObsolete(ByVal sync_vbc_name As String)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes the VBComponent (sync_vbc)
' from the provided Sync-Target-Workbook's VBProject.
' ------------------------------------------------------------------------------
    Const PROC = "SyncObsolete"
    
    On Error GoTo eh
    Dim vbc             As VBComponent
    Dim wbkTarget       As Workbook
    Dim wbkSource       As Workbook
    Dim sItemSyncName   As String
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    Set vbc = GetVBComp(sync_vbc_name, wbkTarget)
    sItemSyncName = ItemSyncName(vbc)
    Log.ServicedItem = vbc
    wbkTarget.VBProject.VBComponents.Remove vbc
    Log.Entry = "Removed!"
    wsSync.VbcItemObsoleteDone(sItemSyncName) = True

    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    UnloadSyncMessage TITLE_SYNC_VBCOMPS
    mSync.RunSync

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Sync(ByRef sync_new As Dictionary, _
                ByRef sync_obsolete As Dictionary, _
                ByRef sync_changed As Dictionary)
' ------------------------------------------------------------------------------
' Called by mSync.RunSync when there are still outstanding VBComponents to be
' synchronized.
' ------------------------------------------------------------------------------
    Const PROC          As String = "Sync"
    
    On Error GoTo eh
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim cllButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    Dim cll         As Collection
    Dim i           As Long
    Dim BttnDiff    As String
    Dim v           As Variant
    Dim wbkTarget    As Workbook
    Dim wbkSource    As Workbook
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    wsService.SyncDialogTitle = TITLE_SYNC_REFS
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_REFS)
    With Msg.Section(1)
        .Label.Text = "Obsolete components:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_obsolete
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(2)
        .Label.Text = "New components:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_new
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(3)
        .Label.Text = "Modified components:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_changed
            .Text.Text = .Text.Text & vbLf & v
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(4)
        .Label.Text = "About synchronizing VB-Components:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Code modification is done in two ways:" & vbLf & _
                     "For Non-Datamodules the code is updated by removing it and re-importing the ExportFile " & _
                     "of the Sync-Source-Workbook's corresponding VB-Component." & vbLf & _
                     "For Datamodules the code is replaced line-by-line based on the ExportFile of the " & _
                     "Sync-Source-Workbook's corresponding VB-Component." & vbLf & vbLf & _
                     "Updating VB-Components one by one has been proved the most stable way though it might " & _
                     "take a bit more effort and time. After each synchronization action the dialog is re-displayed " & _
                     "with the still remaining synchronization tasks."
    End With
    If sync_obsolete.Count + sync_new.Count + sync_changed.Count > 0 Then
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for obsolete VBComponents for being removed
        For i = 1 To Min(7 - AppRunArgs.Count, sync_obsolete.Count)
            Set cll = sync_obsolete.Items()(i - 1)
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(APP_RUN_ARG_BUTTON_CAPTION), vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(APP_RUN_ARG_BUTTON_CAPTION) _
                                        , cll(APP_RUN_ARG_SERVICING_WORKBOOK) _
                                        , cll(APP_RUN_ARG_SERVICE) _
                                        , cll(APP_RUN_ARG_SERVICE_ARG1)
        Next i
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for new VBComponents for being added
        For i = 1 To Min(7 - AppRunArgs.Count, sync_new.Count)
            Set cll = sync_new.Items()(i - 1)
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(APP_RUN_ARG_BUTTON_CAPTION), vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(APP_RUN_ARG_BUTTON_CAPTION) _
                                        , cll(APP_RUN_ARG_SERVICING_WORKBOOK) _
                                        , cll(APP_RUN_ARG_SERVICE) _
                                        , cll(APP_RUN_ARG_SERVICE_ARG1)
        Next i
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for changed VBComponents for being updated
        For i = 1 To Min(7 - AppRunArgs.Count, sync_changed.Count)
            Set cll = sync_changed.Items()(i - 1)
            BttnDiff = cll(4) & vbLf & vbLf & "Display Code Changes"
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(APP_RUN_ARG_BUTTON_CAPTION), BttnDiff, vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(APP_RUN_ARG_BUTTON_CAPTION) _
                                        , cll(APP_RUN_ARG_SERVICING_WORKBOOK) _
                                        , cll(APP_RUN_ARG_SERVICE) _
                                        , cll(APP_RUN_ARG_SERVICE_ARG1)
            mMsg.ButtonAppRun AppRunArgs, BttnDiff _
                                        , cll(APP_RUN_ARG_SERVICING_WORKBOOK) _
                                        , "mSyncComps.RunDsplyDiff" _
                                        , cll(APP_RUN_ARG_SERVICE_ARG1)
        Next i
        
        '~~ Display the mode-less dialog for the confirmation which Reference
        '~~ to be synchronized which may be one by one or all at once
        If cllButtons(cllButtons.Count) = vbLf Then cllButtons.Remove cllButtons.Count
               
        mMsg.Dsply dsply_title:=TITLE_SYNC_VBCOMPS _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=wsService.SyncDialogTop & ";" & wsService.SyncDialogLeft
        
    End If
        
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

