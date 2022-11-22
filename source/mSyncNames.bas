Attribute VB_Name = "mSyncNames"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mSyncNames
'
'
' W. Rauschenberger, Berlin, June 2022
' ------------------------------------------------------------------------

Public Function Scopes(ByVal sc_nme As Name, _
              Optional ByVal sc_wbk As Workbook = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns the scope of the Name (sc_nme) as an object in a Collection which is
' either a Workbook or a Worksheet. When no Workbook (sc_wbk) is provided it
' is the scope of the Name, else all scopes, i.e. all corresponding Names
' scope wheras the Name and the referred range are the criterion.
' Note: Because a Name with a certain referred range may exist with more than
'       one scope the scope(s) are returned as Collection.
' Usage examples:
' Set cll = Scope(MyName, MyWbk)
' For v in cll
'    For Each nme In v.Names
'        Debug.Print Split(nme.Name, "!")(UBound(Split(nme.Name, "!")) _
'                    & "|" & nme.RefersTo _
'                    & "|" & v.Name,
'    Next nme
' Next v
' ----------------------------------------------------------------------------
    Dim cll         As New Collection
    Dim wsh         As Worksheet
    Dim wbk         As Workbook
    Dim ScopeWsh    As Worksheet
    Dim ScopeWbk    As Workbook
    Dim nme As Name
    
    If sc_wbk Is Nothing Then
        cll.Add ScopeName(sc_nme)
    Else
        For Each nme In sc_wbk.Names
            If nme.RefersTo = sc_nme.RefersTo Then
                cll.Add ScopeName(sc_nme)
            End If
        Next nme
        For Each wsh In sc_wbk.Worksheets
            For Each nme In wsh.Names
                If nme.RefersTo = sc_nme.RefersTo Then
                    cll.Add ScopeName(sc_nme)
                End If
            Next nme
        Next wsh
    End If
    
    Set Scopes = cll
    Set cll = Nothing
    
End Function

Public Sub CollectAllItems()
' ------------------------------------------------------------------------------
' Writes the Range Names potentially relevant for bein synchronized to the
' wsSynch sheet.
' Note: Sync_Source-Workbook Names not used in any code line are excempted from
'       being collected and thus will be regarded obsolete when still existing
'       in the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "CollectAllItems"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim nme         As Name
    Dim v1          As Variant
    Dim sProgress   As String
    Dim lMaxLenName As Long
    Dim lMaxLenRef  As Long
    Dim dctSource   As Dictionary
    Dim dctTarget   As Dictionary
    Dim sName       As String
    
    mBasic.BoP ErrSrc(PROC)
    mSyncNames.All mSync.Source, dctSource
    mSyncNames.All mSync.TargetCopy, dctTarget
    lMaxLenName = 0
    lMaxLenRef = 0
    sProgress = String((dctSource.Count + dctTarget.Count) * 2, ".")
    
    '~~ Get max Name length
    For Each v1 In dctSource
        Set nme = dctSource(v1)
        If mName.IsValidUserRangeName(nme) Then
            mSync.MonitorStep "Collecting Names" & sProgress
            sName = ItemSyncName(nme)
            lMaxLenName = Max(lMaxLenName, Len(sName))
            lMaxLenRef = Max(lMaxLenRef, Len(nme.RefersTo))
            sProgress = Left(sProgress, Len(sProgress) - 1)
        End If
    Next v1
    For Each v1 In dctTarget
        Set nme = dctTarget(v1)
        If mName.IsValidUserRangeName(nme) Then
            mSync.MonitorStep "Collecting Names" & sProgress
            sName = ItemSyncName(nme)
            lMaxLenName = Max(lMaxLenName, Len(sName))
            lMaxLenRef = Max(lMaxLenRef, Len(nme.RefersTo))
            sProgress = Left(sProgress, Len(sProgress) - 1)
        End If
    Next v1
    
    wsService.MaxLenName = lMaxLenName
    wsService.MaxLenRefTo = lMaxLenRef
       
    '~~ Again with known max lengths
    For Each v1 In dctSource
        mSync.MonitorStep "Collecting Names" & sProgress
        Set nme = dctSource(v1)
        If mName.IsValidUserRangeName(nme) Then
            If mName.IsInUse(nme, mSync.Source) Then
                If Not dct.Exists(ItemSyncFullName(nme)) Then
                    mDct.DctAdd dct, ItemSyncFullName(nme), nme, order_bykey, seq_ascending, sense_casesensitive
                End If
            End If
        End If
        sProgress = Left(sProgress, Len(sProgress) - 1)
    Next v1

    For Each v1 In dctTarget
        mSync.MonitorStep "Collecting Names" & sProgress
        Set nme = dctTarget(v1)
        If mName.IsValidUserRangeName(nme) Then
            If Not dct.Exists(ItemSyncFullName(nme)) Then
                mDct.DctAdd dct, ItemSyncFullName(nme), nme, order_bykey, seq_ascending, sense_casesensitive
            End If
        End If
        sProgress = Left(sProgress, Len(sProgress) - 1)
    Next v1
    
    mSync.MonitorStep "Collecting Names" & sProgress
    For Each v1 In dct
        wsSync.NmeItemAll(v1) = True
    Next v1
    
xt: Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub NamesInScope(ByVal sn_wsh As Worksheet)
    
    Dim wbk As Workbook
    Dim nme As Name
    
    Set wbk = sn_wsh.Parent
    For Each nme In wbk.Names
        If nme.Parent.Name = sn_wsh.Name Then
            Debug.Print "Name: " & nme.Name & " Refers-To: " & nme.RefersTo
        End If
    Next nme
    
End Sub
                    
Public Function CollectNew() As Dictionary
' ------------------------------------------------------------------------
' Returns a Collection of all Range-Names in the Sync-Source-Workbook
' (cn_wbk_source) which do not exist in the Sync-Target-Workbook. This is
' only the case when range a name in the Sync-Source-Workbook refers-to
' a range which has no name in the Sync-Target-Workbook.
' ------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim dctSource   As Dictionary
    Dim dctTarget   As Dictionary
    Dim nmeSource   As Name
    Dim sProgress   As String
    Dim v           As Variant
    Dim cllNames    As Collection
    
    sProgress = " "
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    mSyncNames.All mSync.Source, dctSource, , True, True ' exclude unused or invalid names
    mSyncNames.All mSync.TargetCopy, dctTarget
    
    For Each v In dctSource
        Set nmeSource = dctSource(v)
        Debug.Print PROC & " nmeSource.Name = " & nmeSource.Name & " (" & v & ")"
        mSync.MonitorStep "Collecting Range Names new" & sProgress
        If IsValidUserRangeName(nmeSource) Then
            If Not dctTarget.Exists(v) Then
                Set cll = New Collection
                cll.Add "Add New" & vbLf & vbLf & _
                        v & vbLf & vbLf & _
                        nmeSource.RefersTo          ' 1. The button's caption
                cll.Add ThisWorkbook                ' 2. The servicing Workbook
                cll.Add "mSyncNames.RunAdd"         ' 3. The service to run
                cll.Add nmeSource                   ' 4. The new name
                mDct.DctAdd dct, ItemSyncFullName(nmeSource), cll, order_bykey, seq_ascending, sense_casesensitive
                Set cll = Nothing
            End If
        End If
        sProgress = sProgress & "."
    Next v

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

Public Sub Properties(ByVal p_nme_source As Name, _
                      ByVal p_nme_target As Name)
' ------------------------------------------------------------------------
' Synchronizes the properties RefersTo and the Scope of the target Name
' (p_nme_target) with the source Name (p_nme_source).
' Precondition: The Worksheet Names in the Sync-Target-Workbook are
'               identical with those in the Sync-Source-Workbook.
' ------------------------------------------------------------------------
    Const PROC = "Properties"
    
    Dim sScopeSheetName As String
    Dim wbkTarget       As Workbook
    Dim wshTarget       As Worksheet
    
    p_nme_target.RefersTo = p_nme_source.RefersTo
    
    If mName.ScopeIsWorkbook(p_nme_source) And Not mName.ScopeIsWorkbook(p_nme_target) Then
        Debug.Print PROC & ": Sync Workbook-Scope for: " & p_nme_target.Name & " based on: " & p_nme_source
        '~~ Delete and re-create the target Name with Scope Workbook
        Set wbkTarget = p_nme_target.Parent
        p_nme_target.Delete
        With p_nme_source
            wbkTarget.Names.Add Name:=mName.Mere(p_nme_source) _
                              , RefersTo:=.RefersTo _
                              , Visible:=.Visible
        End With
    ElseIf mName.ScopeIsWorkSheet(p_nme_source, sScopeSheetName) And Not mName.ScopeIsWorkSheet(p_nme_target) Then
        Debug.Print PROC & ": Sync Worksheet-Scope for: " & p_nme_target.Name & " based on: " & p_nme_source
        '~~ Delete and re-create the target Name with Scope Worksheet
        Set wbkTarget = p_nme_target.Parent.Parent
        p_nme_target.Delete
        Set wshTarget = wbkTarget.Worksheets(sScopeSheetName)
        With p_nme_source
            wshTarget.Names.Add Name:=mName.Mere(p_nme_source) _
                              , RefersTo:=.RefersTo _
                              , Visible:=.Visible
        End With
        
    End If
End Sub

Public Function CollectObsolete() As Dictionary
' ------------------------------------------------------------------------
' Returns a Dictionary of all (mere!) Range-Names in the Sync-Target-
' Workbook which do not exist in the Sync-Source-Workbook or where
' the Name is not/no longer in-use (in any code line or formula) in the
' Sync-Source-Workbook. Obsolete Names are additionally saved to the
' dctObsoleteNames Dictionary.
' ------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim cllNames    As Collection
    Dim dct         As New Dictionary
    Dim nmeSource   As Name
    Dim nmeTarget   As Name
    Dim v           As Variant
    Dim sProgress   As String
    Dim dctTarget   As Dictionary
    Dim dctSource   As Dictionary
    
    sProgress = " "
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    
    mSyncNames.All mSync.Source, dctSource
    mSyncNames.All mSync.TargetCopy, dctTarget
    For Each v In dctTarget
        mSync.MonitorStep "Collecting obsolete Names" & sProgress
        Set nmeTarget = dctTarget(v)
        If IsObsolete(nmeTarget, mSync.Source, mSync.TargetCopy) Then
            Set cll = New Collection
            cll.Add "Remove Obsolete" & vbLf & vbLf & _
                    v & vbLf & vbLf & _
                    nmeTarget.RefersTo          ' 1. The button's caption
            cll.Add ThisWorkbook                ' 2. The servicing Workbook
            cll.Add "mSyncNames.RunRemove"      ' 3. The service to run
            cll.Add nmeTarget                   ' 4. The obsolete name
            mDct.DctAdd dct, ItemSyncFullName(nmeTarget), cll, order_bykey, seq_ascending, sense_casesensitive
            Set cll = Nothing
        End If
        sProgress = sProgress & "."
    Next v
    
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
    Const PROC = "CorrespName"
    Dim nme As Name
    
    Debug.Print "cn_nma_name = " & cn_nme_name
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

Private Function IsObsolete(ByVal io_nme_target As Name, _
                            ByVal io_wbk_source As Workbook, _
                            ByVal io_wbk_target As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Name object's (in_nme_target) Name property is no
' longer relevant which is the case when:
' - no Name in the Sync-Target-Workbook (io_wbk_target) has an identical
'   Name property, i.e. is doesn't exist
' - the Name object's (in_nme_target) Name property is not in use by any of
'   the Sync-Source-Wokbook's (io_wbk_source) code lines
' ----------------------------------------------------------------------------
    Dim nmeSource As Name
    
    If IsValidUserRangeName(io_nme_target) Then
        IsObsolete = Not mName.Exists(io_nme_target, io_wbk_source) And Not mName.Exists(io_nme_target, io_wbk_target)
        If Not IsObsolete Then
            IsObsolete = Not mName.IsInUse(io_nme_target, io_wbk_source)
        End If
    End If
    
End Function

Private Function IsSyncItem(ByVal nme As Name) As Boolean
    With wsSync
        IsSyncItem = .NmeNumberAll > 0 _
              And .NmeItemAll(ItemSyncFullName(nme))
    End With
End Function

Public Function ItemSyncFullName(ByVal isfn_nme As Name, _
                        Optional ByVal isfn_max_name_length As Long = 0) As String
' ----------------------------------------------------------------------------
' Returns a unified Range Name Id including its referred range.
' ----------------------------------------------------------------------------
    If isfn_max_name_length = 0 _
    Then ItemSyncFullName = Align(ItemSyncName(isfn_nme), wsService.MaxLenName) & " " & isfn_nme.RefersTo _
    Else ItemSyncFullName = Align(ItemSyncName(isfn_nme), isfn_max_name_length) & " " & isfn_nme.RefersTo
End Function

Public Function ItemSyncName(ByVal nme As Name) As String
' ----------------------------------------------------------------------------
' Returns a unified Range Name Id.
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    Dim sName   As String
    
    sName = mName.Mere(nme)
    On Error Resume Next
    ItemSyncName = sName & "(" & fso.GetBaseName(nme.Parent.Parent.FullName) & ")"
    If Err.Number <> 0 _
    Then ItemSyncName = sName & "(" & fso.GetBaseName(nme.Parent.FullName) & ")"
    Set fso = Nothing

End Function

Public Sub RunAdd(ByVal ra_nme As Name)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Adds a Name object to the
' Sync-Target-Workbook with the name (ra_name) referring to
' range (ra_ref).
' ------------------------------------------------------------------------------
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE

    Log.ServicedItem = ra_nme
    mSync.TargetCopy.Names.Add ra_nme.Name, ra_nme.RefersTo
    Log.Entry = "New! Added to Sync-Target-Workbook"
    wsSync.NmeItemNewDone(ItemSyncFullName(ra_nme)) = True

End Sub

Public Sub RunAddAdditionla(ByVal raa_nme As Name)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Adds a Name object to the
' Sync-Target-Workbook with the Name object's (raa_name) Name referring to
' the Name object's (raa_name) RefersTo range.
' ------------------------------------------------------------------------------
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE

    Log.ServicedItem = raa_nme
    mSync.TargetCopy.Names.Add raa_nme.Name, raa_nme.RefersTo
    Log.Entry = "A d d i t i o n a l ! name added to Sync-Target-Workbook"
    wsSync.NmeItemNewDone(ItemSyncFullName(raa_nme)) = True

End Sub

Public Sub RunChange(ByVal rc_nme As Name)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Changes the Name property of the
' Name object in the Sync-Target-Workbook which refers to the Name's (rc_nme)
' Range to the Name property of the Name object (rc_nme).
' ------------------------------------------------------------------------------
    Dim nme         As Name
    
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
    
    For Each nme In mSync.TargetCopy.Names
        If nme.RefersTo = rc_nme.RefersTo Then
            Log.ServicedItem = rc_nme
            nme.Name = rc_nme.Name
            Log.Entry = "Name referring to '" & Replace(nme.RefersTo, "=", vbNullString) & "' changed to '" & nme.Name & "'"
            wsSync.NmeItemNewDone(ItemSyncFullName(rc_nme)) = True
            Exit For
        End If
    Next nme
    
End Sub

'Public Sub SyncAllRangeNames()
'' ----------------------------------------------------------------------------
'' Called via Application.Run, synchronizes all Range Names.
'' ----------------------------------------------------------------------------
'    Const PROC = "SyncAllRangeNames"
'
'    On Error GoTo eh
'    Dim nmeTarget   As Name
'    Dim nmeSource   As Name
'    Dim sRangeName  As String
'    Dim sProgress   As String
'
'    sProgress = " "
'    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE
'
'    '~~ 1. Remove obsolete Range Names
'    For Each nmeTarget In mSync.TargetCopy.Names
'        mSync.MonitorStep "Synchronizing Range Names obsolete" & sProgress
'        If IsObsolete(nmeTarget, mSync.Source) Then
'            Log.ServicedItem = nmeTarget
'            sRangeName = ItemSyncFullName(nmeTarget)
'            nmeTarget.Delete
'            Log.Entry = "Obsolete! Removed from Sync-Target-Workbook"
'            wsSync.NmeItemObsoleteDone(sRangeName) = True
'        End If
'        sProgress = sProgress & "."
'    Next nmeTarget
'
'    '~~ 2. Add new Range Names
'    sProgress = " "
'    For Each nmeSource In mSync.Source.Names
'        mSync.MonitorStep "Synchronizing Range Names new" & sProgress
'        If IsNew(nmeSource, mSync.TargetCopy) Then
'            Log.ServicedItem = nmeSource
'            sRangeName = ItemSyncFullName(nmeSource)
'            mSync.TargetCopy.Names.Add nmeSource.Name, nmeSource.RefersTo
'            Log.Entry = "New! Added to Sync-Target-Workbook"
'            wsSync.NmeItemNewDone(sRangeName) = True
'        End If
'        sProgress = sProgress & "."
'    Next nmeSource
'
'    '~~ 3. Synchronize changed RefersTo argument
'    sProgress = " "
'    For Each nmeSource In mSync.Source.Names
'        mSync.MonitorStep "Synchronizing Range Names changed" & sProgress
'        If HasChanged(nmeSource, mSync.TargetCopy, nmeTarget) Then
'            Log.ServicedItem = nmeSource
'            sRangeName = ItemSyncFullName(nmeSource)
'            nmeTarget.Name = nmeSource.Name
'            Log.Entry = "New! Added to Sync-Target-Workbook"
'            wsSync.NmeItemChangedDone(sRangeName) = True
'        End If
'        sProgress = sProgress & "."
'    Next nmeSource
'
'    wsSync.NmeSyncDone = True
'
'    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
'   UnloadSyncMessage TITLE_SYNC_NAMES
'   mSync.RunSync
'
'xt: Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Public Sub RunRemove(ByVal rr_nme As Name)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes the Name object (rr_nme)
' from the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Dim s           As String
    
    mService.EstablishServiceLog mSync.TargetCopy, mCompManClient.SRVC_SYNCHRONIZE

    Log.ServicedItem = rr_nme
    s = ItemSyncFullName(rr_nme)
    rr_nme.Delete
    Log.Entry = "Obsolete! Removed from Sync-Target-Workbook"
    wsSync.NmeItemObsoleteDone(s) = True

End Sub

Public Sub All(ByVal all_wbk As Workbook, _
      Optional ByRef all_by_name As Dictionary, _
      Optional ByRef all_by_ref As Dictionary, _
      Optional ByVal all_in_use_only As Boolean = False, _
      Optional ByVal all_valid_only As Boolean = True)
' ------------------------------------------------------------------------
' Returns a Dictionary with all (mere) Names in Workbook (all_wbk) ordered
' ascending by a unique key.
' ------------------------------------------------------------------------
    Const PROC = "All"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim nme         As Name
    Dim wsh         As Worksheet
    Dim sKeyName    As String
    Dim sKeyRef     As String
    Dim sItem       As String
    Dim v           As Variant
    Dim i           As Long
    Dim lMaxName    As Long
    Dim lMaxRef     As Long
    
    If all_by_name Is Nothing Then Set all_by_name = New Dictionary
    If all_by_ref Is Nothing Then Set all_by_ref = New Dictionary
    
    For Each nme In all_wbk.Names
        sKeyName = mName.Mere(nme)
        sKeyRef = nme.RefersTo
        If all_valid_only Then
            If mName.IsValidUserRangeName(nme) Then
                If all_in_use_only Then
                    If mName.IsInUse(nme, all_wbk) Then
                        If Not all_by_name.Exists(sKeyName) Then
                            mDct.DctAdd all_by_name, sKeyName, nme, order_bykey, seq_ascending
                        End If
                        If Not all_by_ref.Exists(sKeyRef) Then
                            mDct.DctAdd all_by_ref, sKeyRef, nme, order_bykey, seq_ascending
                        End If
                    End If
                Else
                    If Not all_by_name.Exists(sKeyName) Then
                        mDct.DctAdd all_by_name, sKeyName, nme, order_bykey, seq_ascending
                    End If
                    If Not all_by_ref.Exists(sKeyRef) Then
                        mDct.DctAdd all_by_ref, sKeyRef, nme, order_bykey, seq_ascending
                    End If
                End If
            End If
        Else ' include invalid
            If all_in_use_only Then
                If mName.IsInUse(nme, all_wbk) Then
                    If Not all_by_name.Exists(sKeyName) Then
                        mDct.DctAdd all_by_name, sKeyName, nme, order_bykey, seq_ascending
                    End If
                    If Not all_by_ref.Exists(sKeyRef) Then
                        mDct.DctAdd all_by_ref, sKeyRef, nme, order_bykey, seq_ascending
                    End If
                End If
            Else
                If Not all_by_name.Exists(sKeyName) Then
                    mDct.DctAdd all_by_name, sKeyName, nme, order_bykey, seq_ascending
                End If
                If Not all_by_ref.Exists(sKeyRef) Then
                    mDct.DctAdd all_by_ref, sKeyRef, nme, order_bykey, seq_ascending
                End If
            End If
        End If
    Next nme

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function UniqueKey(ByVal uk_nme As Name, _
                  Optional ByVal uk_delimiter As String = "|") As String
' ----------------------------------------------------------------------------
' Returns a string with the syntax: name.range.[scope]
' - name:  the mere name string without any Worksheet-Name prefix
' - range: the RefersTo property of the Name object
' - scope: optional. No scope information means Workbook scope, else it is
'          the name of the "in scope Worksheet"
' The returned string is a disambiguated key for all Names in any number of
' Workbooks.
' ----------------------------------------------------------------------------
    UniqueKey = mSyncNames.ScopeName(uk_nme)
End Function

Public Function ScopeName(ByVal sn_nme As Name) As String
    Dim wbk As Workbook
    Dim wsh As Worksheet
    
    ScopeName = mName.Mere(sn_nme) & "|" & sn_nme.RefersTo & "|"
    If mName.Scope(sn_nme, wbk, wsh) = enWorksheet _
    Then ScopeName = ScopeName & wsh.Name
    
End Function

Public Sub Sync(ByRef sync_new As Dictionary, _
                ByRef sync_obsolete As Dictionary)
' ------------------------------------------------------------------------------
' Collects to be synchronized Names (obsolete and new) and displays them one by
' one in a mode-less dialog for being confirmed.
' Called by mSync.RunSync
' ------------------------------------------------------------------------------
    Const PROC = "Sync"
    
    On Error GoTo eh
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim cllButtons  As New Collection
    Dim AppRunArgs  As New Dictionary
    Dim i           As Long
    Dim v           As Variant
    Dim cll         As Collection
    
    If sync_obsolete.Count + sync_new.Count = 0 Then GoTo xt
    '~~ There's at least one Range Name in need of synchronization
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_NAMES)
    
    With Msg.Section(1)
        .Label.Text = "Remove Obsolete:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Removes any name in the Sync-Target-Workbook which is regarded obsolete because it either no " & _
                     "longer exists in the Sync-Source-Workbook or it is already obsolete in the Sync-Source-Workbook " & _
                     "not used in any VB-Component's code line or the Name's specification has changed so that it will " & _
                     "will be replaced by one which then is regarded new."
    End With
    With Msg.Section(2)
        .Label.Text = "Add New:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Add a Name which is regarded new in the Sync-Source-Workbook because it not exists in the " & _
                     "Sync-Target-Workbook." & vbLf & _
                     "Attention!" & vbLf & _
                     "When synchronizing new names it must be made sure the name refers-to the correct range in the Sync-Target-Workbook."
    End With
    With Msg.Section(5)
        With .Label
            .Text = "See: Synchronization of new Names"
            .FontColor = rgbBlue
            .OpenWhenClicked = mCompMan.README_URL & mSync.README_SYNC_CHAPTER_NEW_NAMES
        End With
        .Text.Text = "The chapter is providing additional information"
    End With
        
    If sync_obsolete.Count + sync_new.Count > 0 Then
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for obsolete Names for being removed
        For i = 1 To Min(7 - AppRunArgs.Count, sync_obsolete.Count)
            Set cll = sync_obsolete.Items()(i - 1)
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(APP_RUN_ARG_BUTTON_CAPTION), vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(APP_RUN_ARG_BUTTON_CAPTION) _
                                        , cll(APP_RUN_ARG_SERVICING_WORKBOOK) _
                                        , cll(APP_RUN_ARG_SERVICE) _
                                        , cll(APP_RUN_ARG_SERVICE_ARG1)
        Next i
        
        '~~ Prepare the Command-Buttons and their corresponding Application.Run action
        '~~ for new Names for being added
        For i = 1 To Min(7 - AppRunArgs.Count, sync_new.Count)
            Set cll = sync_new.Items()(i - 1)
            Set cllButtons = mMsg.Buttons(vbLf, cllButtons, cll(APP_RUN_ARG_BUTTON_CAPTION), vbLf)
            mMsg.ButtonAppRun AppRunArgs, cll(APP_RUN_ARG_BUTTON_CAPTION) _
                                        , cll(APP_RUN_ARG_SERVICING_WORKBOOK) _
                                        , cll(APP_RUN_ARG_SERVICE) _
                                        , cll(APP_RUN_ARG_SERVICE_ARG1)
        Next i
        
        Application.EnableEvents = True
        '~~ Display the mode-less dialog for the Names synchronization to run
        mMsg.Dsply dsply_title:=TITLE_SYNC_NAMES _
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

