Attribute VB_Name = "mSyncShapes"
Option Explicit
' -------------------------------------------------------------------
' Standard Module mShapes: Services to synchronize new, obsolete, and
'                          changed properties of sheet shapes and
'                          OLEObjects.
'
' -------------------------------------------------------------------
Private Const TITLE_SYNC_SHAPES = "VB-Project-Synchronization: Shapes"

Private dctKnownInSync      As Dictionary
Private dctKnownNew         As Dictionary
Private dctKnownObsolete    As Dictionary
Private dctKnownChanged     As Dictionary

Private Enum enCorrespondingShapesQuality
    enOrAny
    enAndAll
End Enum

Public Property Get KnownChanged(Optional ByVal nk_id As String) As Boolean
    If Not dctKnownChanged Is Nothing _
    Then KnownChanged = dctKnownChanged.Exists(nk_id)
End Property

Public Property Let KnownChanged(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownChanged, nk_id
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
    KnownObsolete = dctKnownObsolete.Exists(nk_id)
End Property

Public Property Let KnownObsolete(Optional ByVal nk_id As String, _
                                               ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownObsolete, nk_id
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function NoOfShapes(ByVal n_wbk As Workbook) As Long
    Dim wsh As Worksheet
    For Each wsh In n_wbk.Worksheets
        NoOfShapes = NoOfShapes + wsh.Shapes.Count
    Next wsh
End Function

Public Sub AppRunSyncAll()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "AppRunSyncAll"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If dctKnownChanged.Count > 0 Then mSyncShapes.AppRunChanged
    If dctKnownObsolete.Count > 0 Then mSyncShapes.AppRunObsolete
    If dctKnownNew.Count > 0 Then mSyncShapes.AppRunNew

xt: mBasic.EoP ErrSrc(PROC)
    If lSyncMode <> SyncSummarized Then
        mService.MessageUnload TITLE_SYNC_SHAPES
        mSync.RunSync
    End If
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

'Public Sub AppRunSyncAll()
'' ------------------------------------------------------------------------------
'' Called via Application.Run by CommonButton: Synchronizes all Sheets Shapes by
'' removing obsolete, adding new, and changing the properties when changed.
'' Precondition: All sheet synchronizations are done.
'' ------------------------------------------------------------------------------
'    Const PROC = "AppRunSyncAll"
'
'    On Error GoTo eh
'    Dim wshSource   As Worksheet
'    Dim wshTarget   As Worksheet
'    Dim shpSource   As Shape
'    Dim shpTarget   As Shape
'    Dim sShapeName  As String
'    Dim dctPrprtys  As New Dictionary
'    Dim v           As Variant
'    Dim wbkSource   As Workbook
'    Dim wbkTarget   As Workbook
'
'    mBasic.BoP ErrSrc(PROC)
'    If Not wsSyncLog.SummaryDone("Worksheets") _
'    Then Err.Raise AppErr(1), ErrSrc(PROC), "Shape synchronization cannot be done when there are Worksheet " & _
'                                            "synchronizations yet not done!"
'
'    lPropertyMaxLen = mSyncShapePrprtys.PropertyMaxLen
'
'    '~~ Synchronize, respectively remove obsolete Shapes
'    For Each wshTarget In mSync.TargetWorkingCopy.Worksheets
'        If mSyncSheets.Corresponding(c_wsh:=wshTarget _
'                                   , c_wbk:=mSync.Source _
'                                   , c_quality:=enCorrespondingShapesQuality.enOrNameCodeName _
'                                   , c_wsh_result:=wshSource) = "1:1" Then
'            For Each shpTarget In wshTarget.Shapes
'                If IsObsolete(shpTarget, wshSource) Then
'                    sShapeName = SyncId(shpTarget)
'                    shpTarget.Delete
'                    wsSyncLog.Done "obolete", "Shape", sShapeName, "removed", "Obsolete (deleted)", shpTarget
'                End If
'            Next shpTarget
'        End If
'    Next wshTarget
'
'    '~~ Synchronize new Shapes
'    For Each wshSource In mSync.Source.Worksheets
'        For Each shpSource In wshSource.Shapes
'            If mSyncSheets.Corresponding(c_wsh:=wshSource _
'                                       , c_wbk:=mSync.TargetWorkingCopy _
'                                       , c_quality:=enCorrespondingShapesQuality.enWeak _
'                                       , c_wsh_result:=wshTarget) = "1:1" Then
'                For Each shpSource In wshSource.Shapes
'                    If IsNew(shpSource, wshTarget) Then
'                        '~~ Synchronize new Shapes
'                        Srvc.ServicedItem = shpSource
'                        CopyShapeToTarget shpSource, wshTarget
'                    End If
'                    '~~ Synchronize the properties which have changed of all Shapes
'ns:             Next shpSource
'            End If
'        Next shpSource
'    Next wshSource
'
'    '~~ Synchronize Shape/OOB Properties
'    For Each wshSource In mSync.Source.Worksheets
'            mSyncSheets.Corresponding c_wsh:=wshSource _
'                                    , c_wbk:=mSync.TargetWorkingCopy _
'                                    , c_quality:=enCorrespondingShapesQuality.enWeak _
'                                    , c_wsh_result:=wshTarget
'            For Each shpSource In wshSource.Shapes
'                If Corresponding(c_shp:=shpSource _
'                               , c_quality:=enOrNameCodeName _
'                               , c_wbk_source:=
'                mSyncShapePrprtys.ShapeTarget = mSyncShapes.CorrespondingShape(shpSource, wshTarget)
'                mSyncShapePrprtys.SyncProperties dctPrprtys
'            Next shpSource
'        Next shpSource
'    Next wshSource
'
'    wsSyncLog.SummaryDone("Sheet-Shapes") = True
'    For Each v In dctPrprtys
'        wsSyncLog.SummaryDoneShapeProperties v
'    Next v
'
'    '~~ Re-display the synchronization dialog for still to be synchronized items
'    mService.MessageUnload TITLE_SYNC_SHAPES
'    mSync.RunSync
'
'xt: mBasic.EoP ErrSrc(PROC)
'    Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Private Sub AppRunChanged()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all Sheets Shapes by
' removing obsolete, adding new, and changing the properties when changed.
' Precondition: All sheet synchronizations are done.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunChanged"
    
    On Error GoTo eh
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    Dim shpSource   As Shape
    Dim shpTarget   As Shape
    Dim dctPrprtys  As New Dictionary
    Dim vSource     As Variant
    Dim vTarget     As Variant
    Dim sPrgrss     As String
    Dim i           As Long
    
    mBasic.BoP ErrSrc(PROC)
    If Not wsSyncLog.SummaryDone("Worksheets") _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The precondition of done Worksheet synchronizations is not met!"
            
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    vSource = Split(AppRunChangedIdsSource, ",")
    vTarget = Split(AppRunChangedIdsTarget, ",")
    sPrgrss = "Synchronizing Worksheets > Changed Name or CodeName " & String(UBound(vSource), ".")
    mService.DsplyStatus sPrgrss
    
    For i = LBound(vSource) To UBound(vSource)
        Set shpSource = GetShape(vSource(i), wbkSource, wshSource)
        Set shpTarget = GetShape(vTarget(i), wbkTarget, wshTarget)
        mSyncShapePrprtys.ShapeSource = shpSource
        mSyncShapePrprtys.ShapeTarget = shpTarget
        mSyncShapePrprtys.SyncProperties dctPrprtys
    Next i
    dctKnownChanged.RemoveAll ' indicates done
        
xt: mService.MessageUnload TITLE_SYNC_SHAPES
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppRunChangedIdsSource() As String
    AppRunChangedIdsSource = DueSyncIdsByAction(enSyncObjectKindShape, "Change Property", "to")
    If Right(AppRunChangedIdsSource, 1) = "," Then AppRunChangedIdsSource = Left(AppRunChangedIdsSource, Len(AppRunChangedIdsSource) - 1)
End Function

Private Function AppRunChangedIdsTarget() As String
    AppRunChangedIdsTarget = DueSyncIdsByAction(enSyncObjectKindShape, "Change Property", "from")
    If Right(AppRunChangedIdsTarget, 1) = "," Then AppRunChangedIdsTarget = Left(AppRunChangedIdsTarget, Len(AppRunChangedIdsTarget) - 1)
End Function

Private Sub AppRunNew()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all Sheets Shapes by
' removing obsolete, adding new, and changing the properties when changed.
' The service considers Worksheets yet not synchronized in order to provide
' the independance which services not only testing but also the sequence in
' which synchronization is done.
' ------------------------------------------------------------------------------
    Const PROC = "AppRunNew"
    
    On Error GoTo eh
    Dim i           As Long
    Dim shpSource   As Shape
    Dim v           As Variant
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    v = Split(mSync.AppRunNewIds(enSyncObjectKindShape), ",")
    mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepSyncing, enSyncActionAddNew, 0)
    
    For i = LBound(v) To UBound(v)
        Set shpSource = GetShape(v(i), wbkSource, wshSource)
        mSyncSheets.Corresponding c_wsh:=wshSource _
                                , c_wbk:=wbkTarget _
                                , c_quality:=enCorrespondingSheetsQuality.enAndNameCodeName _
                                , c_wsh_result:=wshTarget
        CopyShapeToTarget shpSource, wshTarget
        mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepSyncing, enSyncActionAddNew, i + 1)
    Next i
    dctKnownNew.RemoveAll ' indicates done
        
xt: mService.MessageUnload TITLE_SYNC_SHAPES
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AppRunObsolete()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes a shape (sync_shp_target)
' identified by its Name property, from a Worksheet (sync_wsh_target).
' ------------------------------------------------------------------------------
    Const PROC = "AppRunObsolete"
    
    On Error GoTo eh
    Dim shpTarget   As Shape
    Dim wshTarget   As Worksheet
    Dim v           As Variant
    Dim i           As Long
    Dim sPrgrss     As String
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mSync.TargetWorkingCopy
    Set wbkSource = mSync.source
    v = Split(mSync.AppRunObsoleteIds(enSyncObjectKindShape), ",")
    sPrgrss = "Synchronizing Shapes > Obsolete " & String(UBound(v) + 1, ".")
    mService.DsplyStatus sPrgrss
    
    For i = LBound(v) To UBound(v)
        Set shpTarget = GetShape(v(i), wbkTarget, wshTarget)
        Srvc.ServicedItem = shpTarget
        shpTarget.Delete
        If Err.Number = 0 Then
            If RunRemoveAsserted(shpTarget.Name, wshTarget) Then
                wsSyncLog.Done "obsolete", "Shape", SyncId(shpTarget), "removed", "Removed from Sync-Target_Workbook's working copy!"
            Else
                Stop
            End If
        Else
            wsSyncLog.Done "obsolete", "Shape", SyncId(shpTarget), "failed!", "Remove from Sync-Target_Workbook's working failed!"
        End If
        mService.DsplyStatus Left(sPrgrss, Len(sPrgrss) - (i + 1))
    Next i
    dctKnownObsolete.RemoveAll ' indicates done

xt: mService.MessageUnload TITLE_SYNC_SHAPES
    mBasic.BoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Collect(ByVal c_wbk_source As Workbook, _
                   ByVal c_wbk_target As Workbook)
' ------------------------------------------------------------------------------
' Returns a collection of all those sheet controls at least on property has
' changed.
' ------------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim shpTarget   As Shape
    Dim shpSource   As Shape
    Dim wshTarget   As Worksheet
    Dim wshSource   As Worksheet
    Dim lCount      As Long
    Dim sId         As String
    Dim wbkSource   As Workbook
    Dim wbkTarget   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    If mSyncShapes.AllDone(c_wbk_source, c_wbk_target) Then GoTo xt
    If lSyncMode = SyncByKind Then mSync.InitDueSyncs

    '~~ Collect obsolete (shapes in target worksheet not existing in corresponding source sheet)
    If dctKnownObsolete Is Nothing Then
        Set dctKnownObsolete = New Dictionary
        lCount = GetShapesCount(wbkTarget)
        mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepCollecting, enSyncActionChanged, 0)
        
        For Each wshTarget In c_wbk_target.Worksheets
            mSyncSheets.Corresponding c_wsh:=wshTarget _
                                    , c_wbk:=wbkSource _
                                    , c_quality:=enOrNameCodeName _
                                    , c_wsh_result:=wshSource
            For Each shpTarget In wshTarget.Shapes
                If Not shpTarget.Name Like "Comment *" Then
                    sId = SyncId(shpTarget)
                    If Not Exists(shpTarget, wshSource) Then
                        If Not KnownObsolete(sId) Then
                            mSync.DueSyncLet , enSyncObjectKindShape, enSyncActionRemoveObsolete, , sId
                        End If
                    End If
                End If
                mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepCollecting, enSyncActionRemoveObsolete, dctKnownChanged.Count)
            Next shpTarget
        Next wshTarget
    End If
    
    '~~ Collect new
    If dctKnownNew Is Nothing Then
        Set dctKnownNew = New Dictionary
        lCount = GetShapesCount(wbkSource)
        mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepCollecting, enSyncActionAddNew, 0)
        
        For Each wshSource In c_wbk_source.Worksheets
            mSyncSheets.Corresponding c_wsh:=wshSource _
                                    , c_wbk:=c_wbk_target _
                                    , c_quality:=enCorrespondingSheetsQuality.enOrNameCodeName _
                                    , c_wsh_result:=wshTarget
            
        Next wshSource
    End If
    
    '~~ Collect changed
    If dctKnownChanged Is Nothing Then
        Set dctKnownChanged = New Dictionary
        mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepCollecting, enSyncActionChanged, 0)
        
        For Each wshSource In wbkSource.Worksheets
            mSyncSheets.Corresponding c_wsh:=wshSource _
                                    , c_wbk:=c_wbk_target _
                                    , c_quality:=enCorrespondingSheetsQuality.enOrNameCodeName _
                                    , c_wsh_result:=wshTarget
            For Each shpSource In wshSource.Shapes
                If Exists(shpSource, wshTarget, shpTarget) Then
                    mSyncShapePrprtys.CollectChanged shpSource
                End If
                mService.DsplyStatus mSync.Progress(enSyncObjectKindShape, enSyncStepCollecting, enSyncActionChanged, dctKnownChanged.Count)
            Next shpSource
        Next wshSource
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Collected(ByVal c_action As enSyncAction) As Long
    Select Case True
        Case mSync.SyncActionIsChange(c_action):    Collected = dctKnownChanged.Count
        Case c_action = enSyncActionRemoveObsolete: Collected = dctKnownObsolete.Count
        Case c_action = enSyncActionAddNew:         Collected = dctKnownNew.Count
    End Select
End Function

Private Function GetShapesCount(ByVal c_wbk As Workbook) As Long
    Dim wsh As Worksheet
    For Each wsh In c_wbk.Worksheets
        GetShapesCount = GetShapesCount + wsh.Shapes.Count
    Next wsh
End Function

'Public Sub CollectAllItems()
'' ------------------------------------------------------------------------------
'' Writes: - the Worksheet Shapes potentially synchronized and
''         - the Shape-Properties able to be synchronized (writeable)
''         to the wsSynch sheet.
'' ------------------------------------------------------------------------------
'    Const PROC = "CollectAllItems"
'
'    On Error GoTo eh
'    Dim shp         As Shape
'    Dim wsh         As Worksheet
'    Dim ShapeId     As String
'    Dim dctShapes   As New Dictionary
'    Dim dctPrprtys  As New Dictionary
'    Dim sPrgrss   As String
'    Dim wshTarget   As Worksheet
'
'    mBasic.BoP ErrSrc(PROC)
'    wsService.MaxLenShapeName = MaxLenShapeName(mSync.TargetWorkingCopy, mSync.Source)
'
'    For Each wsh In mSync.Source.Worksheets
'        mSyncSheets.Corresponding c_wsh:=wsh _
'                                , c_wbk:=mSync.TargetWorkingCopy _
'                                , c_quality:=enCorrespondingShapesQuality.enWeak _
'                                , c_wsh_result:=wshTarget
'            If Not wshTarget Is Nothing Then
'            For Each shp In wsh.Shapes
'                mSync.MonitorStep "Collecting Sheet Shapes " & wsh.Name & sPrgrss
'                With shp
'                    If InStr(.Name, " ") Then
'                        Srvc.ServicedItem = shp
'                        Srvc.LogEntry = "Will not be synchronized (shape does not have a user specified name)!"
'                        Debug.Print "shp.Name: " & SyncId(shp) & " (not synchronized)"
'                        GoTo ns1  ' Comments are not synced as shapes
'                    ElseIf .Type = msoComment Then
'                        Srvc.ServicedItem = shp
'                        Srvc.LogEntry = "Type of Shape intentionally not synchronized!"
'                        Debug.Print "shp.Name: " & SyncId(shp) & " (not synchronized by intention)"
'                        GoTo ns1  ' next shape
'                    End If
'                    ShapeId = SyncId(shp)
'                    If Not dctShapes.Exists(ShapeId) Then
'                        Srvc.ServicedItem = shp
'                        mDct.DctAdd dctShapes, ShapeId, shp, order_bykey, seq_ascending
''                        mSyncShapePrprtys.PropertiesWriteable shp, dctPrprtys
'                    End If
'                End With
'                sPrgrss = sPrgrss & "."
'ns1:         Next shp
'        End If
'        sPrgrss = vbNullString
'    Next wsh
'
'    Set dctShapes = Nothing
'    Set dctPrprtys = Nothing
'
'xt: mBasic.EoP ErrSrc(PROC)
'    Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Private Sub CopyShape(ByVal shp_source As Shape, _
                      ByVal wsh_target As Worksheet, _
                      ByRef shp_target As Shape)
' ------------------------------------------------------------------------------
' Copies the Shape (shp_source) to Worksheet (wsh_target) and returns the target
' Shape (shp_target). The procedure tries to place the  Shape on the target
' sheet at the same cell (row/column) as the source Shape. In case the copy had
' failed the returned Shape (shp_target) is Nothing!
' ------------------------------------------------------------------------------
    Dim Rng As Range
    
    Set Rng = wsh_target.Cells(shp_source.TopLeftCell.row, shp_source.TopLeftCell.Column)
    shp_source.Copy
    wsh_target.Paste Rng
    Set shp_target = wsh_target.Shapes(wsh_target.Shapes.Count)
    With shp_target
        .Name = shp_source.Name
        .Top = shp_source.Top
        .Left = shp_source.Left
    End With
    If shp_source.Type = msoOLEControlObject Then
        shp_target.OLEFormat.Object.Name = shp_source.OLEFormat.Object.Name
    End If
    
End Sub

Private Sub CopyShapeToTarget(ByVal shp_source As Shape, _
                              ByVal wsh_target As Worksheet)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "CopyShapeToTarget"
    
    On Error GoTo eh
    Dim sSyncId As String
    Dim wshSource   As Worksheet
    Dim shpTarget   As Shape
    
    Set wshSource = shp_source.Parent
    sSyncId = SyncId(shp_source)
    
    CopyShape shp_source, wsh_target, shpTarget
    
    If SyncNewAsserted(shp_source, wsh_target) Then
        wsSyncLog.Done vbNullString, "Shape", sSyncId, "Copied from 'Sync-Source-Workbook'", shpTarget
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function CorrespondingOOB(ByVal c_shp_source As Shape, _
                                 ByVal c_wsh_target As Worksheet, _
                                 ByRef c_oob_result As OLEObject) As OLEObject
' ------------------------------------------------------------------------------
' Returns the OleObject in the Sync-Target-Workbook's Worksheet which corres-
' ponds to the source Shape's (c_shp_source) OOB.
' ------------------------------------------------------------------------------
    Dim oob             As OLEObject
    Dim sShapeName      As String
    Dim sOOBCodeName    As String
    
    mSyncShapes.ShapeName c_shp_source, sShapeName, sOOBCodeName
    
    For Each oob In c_wsh_target.OLEObjects
        If oob.Name = sOOBCodeName Or oob.ShapeRange.Name = sShapeName Then
            '~~ Note: The names (shape-name and code-name) may not have been synchronized yet
            Set c_oob_result = oob
            Set CorrespondingOOB = oob
            Exit For
        End If
    Next oob

End Function

Public Function CorrespondingShape(ByVal c_shp_source As Shape, _
                                   ByVal c_wsh_target As Worksheet, _
                          Optional ByRef c_shp_result As Shape) As Shape
' ------------------------------------------------------------------------------
' Returns the Shape in the Worksheet (c_wsh_target) which corresponds to the
' source Shape (c_shp_source). I.e. the Shape with either the same Name or -
' in case of an ActiveX-Control - the object Name as object (c_shp_result).
' ------------------------------------------------------------------------------
    Set CorrespondingShape = Nothing
    If Exists(c_shp_source, c_wsh_target, c_shp_result) Then
        Set CorrespondingShape = c_shp_result
    End If
End Function
                             
Public Function AllDone(ByVal d_wbk_source As Workbook, _
                        ByVal d_wbk_target As Workbook) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when there are no more sheet synchronizations outstanding plus
' the collections of the outstanding items.
' ----------------------------------------------------------------------------
    Const PROC = "AllDone"
    
    On Error GoTo eh
    Dim lInSync         As Long
    Dim lTarget         As Long
    Dim lSource         As Long
    Dim wsh             As Worksheet
    Dim shp             As Shape
    Dim lMaxLen         As Long
    
    mBasic.BoP ErrSrc(PROC)
    AllDone = Not DueSyncKindOfObject(enSyncObjectKindShape)
    If Not AllDone Then
        For Each wsh In d_wbk_source.Worksheets
            For Each shp In wsh.Shapes
                lSource = lSource + 1
                If wsService.MaxLenShapeSyncId = 0 Then lMaxLen = mBasic.Max(lMaxLen, Len(SyncId(shp)))
            Next shp
        Next wsh
        For Each wsh In d_wbk_target.Worksheets
            For Each shp In wsh.Shapes
                lTarget = lTarget + 1
                If wsService.MaxLenShapeSyncId = 0 Then lMaxLen = mBasic.Max(lMaxLen, Len(SyncId(shp)))
            Next shp
        Next wsh
        lInSync = CollectInSync(d_wbk_source, d_wbk_target)
    
        If lTarget = lInSync _
        And lSource = lInSync Then
            AllDone = True
            mMsg.MsgInstance TITLE_SYNC_SHAPES, True
            wsSyncLog.SummaryDone("Shape") = True
            mSync.DueSyncKindOfObjects.DeQueue , enSyncObjectKindShape
        Else
            If wsService.MaxLenShapeSyncId = 0 Then wsService.MaxLenShapeSyncId = lMaxLen
        End If
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Property Get KnownInSync(Optional ByVal k_nme_id As String) As Boolean
    If Not dctKnownInSync Is Nothing _
    Then KnownInSync = dctKnownInSync.Exists(k_nme_id)
End Property

Private Property Let KnownInSync(Optional ByVal k_nme_id As String, _
                                          ByVal b As Boolean)
    If b Then mSync.CollectKnown dctKnownInSync, k_nme_id
End Property


Private Function CollectInSync(ByVal c_wbk_source As Workbook, _
                               ByVal c_wbk_target As Workbook) As Long
' ----------------------------------------------------------------------------
' Collects the already in sync Names (in dctItemsAlreadyInSync) and returns
' their number.
' ----------------------------------------------------------------------------
    Const PROC = "CollectInSync"
    
    On Error GoTo eh
    Dim bNoneDiffers    As Boolean
    Dim dctSource       As Dictionary
    Dim dctTarget       As Dictionary

    Dim shpSource       As Shape
    Dim sIdSource       As String
    Dim wshSource       As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    Set dctKnownInSync = Nothing
    Set dctKnownInSync = New Dictionary
    
    For Each wshSource In c_wbk_source.Worksheets
        For Each shpSource In wshSource.Shapes
            sIdSource = mSyncShapes.SyncId(shpSource)
            If Corresponding(c_shp:=shpSource _
                           , c_quality:=enOrNameCodeName _
                           , c_wbk_source:=c_wbk_source _
                           , c_wbk_target:=c_wbk_target _
                           , c_dct_source:=dctSource _
                           , c_dct_target:=dctTarget _
                           , c_none_differs:=bNoneDiffers _
                           ) = "Source-1:Target-1" Then
                If Not KnownInSync(sIdSource) And bNoneDiffers Then
                    KnownInSync(sIdSource) = True
                End If
            End If
        Next shpSource
    Next wshSource

xt: CollectInSync = dctKnownInSync.Count
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Corresponding(ByVal c_shp As Shape, _
                               ByVal c_quality As enCorrespondingShapesQuality, _
                      Optional ByVal c_wbk_source As Workbook = Nothing, _
                      Optional ByVal c_wbk_target As Workbook = Nothing, _
                      Optional ByRef c_dct_source As Dictionary, _
                      Optional ByRef c_dct_target As Dictionary, _
                      Optional ByRef c_none_differs As Boolean, _
                      Optional ByRef c_shp_corresponding_target As Shape, _
                      Optional ByRef c_shp_corresponding_source As Shape) As String
' ------------------------------------------------------------------------
' Returns all Shapes in the Sync-Source-Workbook (c_wbk_source) - when
' provided - and in the Sync-Target-Workbook (c_wbk_target) - when
' provided - which correspond in accordance with the requested quality
' (c_quality).
' When the result is expressed as a string Source-n:Target-n. When the
' result is corresponding result is Source-1:Target-1 all property
' differencies are returned.
' When the mode (c_quality) is enWeak a different scope still corresponds.
' When the mode is enStrong correspondence is equal Shape, RefersTo and
' Scope.
' The syntax of the returned string is tn:m, whereby n = target- and m =
' source correspondencies.
' Note: While it is possible to provide a range with different names it is
'       not pssible to have the same name referring different ranges. This
'       is impossible neither by using the ShapeManager nor via VBA because
'       a subsequently added Shape with the same name ends up with the
'       previous added Shape object.
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
        mSyncShapes.CorrespondingShapes c_shp:=c_shp _
                                      , c_in:=c_wbk_source _
                                      , c_quality:=c_quality _
                                      , c_dct_corresponding:=c_dct_source
        lSource = c_dct_source.Count
    End If
    If Not c_wbk_target Is Nothing Then
        sTarget = "Target-"
        mSyncShapes.CorrespondingShapes c_shp:=c_shp _
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
        If lSource = 1 Then Set c_shp_corresponding_source = c_dct_source.Items()(0)
    End If
    If Not c_wbk_target Is Nothing Then
        Select Case lTarget
            Case 0, 1: sTarget = sTarget & lTarget
            Case Else: sTarget = sTarget & "n"
        End Select
        If lTarget = 1 Then Set c_shp_corresponding_target = c_dct_target.Items()(0)
    End If
        
xt: Corresponding = sSource & sRelate & sTarget
'    If lTarget = 1 And lSource = 1 Then
'        PropertiesDiffer c_shp_corresponding_source, c_shp_corresponding_target, c_name_differs, c_refto_differs, c_scope_differs, c_none_differs
'    End If
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CorrespondingShapes(ByVal c_shp As Shape, _
                                     ByVal c_in As Variant, _
                                     ByVal c_quality As enCorrespondingShapesQuality, _
                            Optional ByRef c_dct_corresponding As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all Shape objects in the provided Workbook (c_wbk)
' With a name equal to the provided Shape object's (c_shp) name. Shape objects
' with an invalid Shape or invalid RefersTo property are ignored.
' - In weak mode (c_quality) corresponding are all with the same name but with
'   different RefersTo and or different scope.
' - In strong mode only same name, same RefersTo and same Scope corresponds.
' Note: The correspondence has to consider yet not synchronized Worksheets
'       which means that the corresponding target sheet name may differ from
'       the source sheet name but when the CodeShapes are equal are still
'       corresponding.
' ----------------------------------------------------------------------------
    Const PROC = "CorrespondingShapes"
    
    On Error GoTo eh
    Dim shp             As Shape
    Dim dct             As New Dictionary
    Dim sId             As String
    Dim sIdWbk          As String
    
    mBasic.BoP ErrSrc(PROC)
    sIdWbk = SyncId(c_shp)
    For Each shp In c_in.Shapes
        sId = SyncId(shp)
'        PropertiesDiffer c_shp, shp, bShapeDiffers, bNoneDiffers
'        If c_quality = enOrNameCodeName Then
'            If bShapeDiffers And bRefToDiffers Then
'                '~~ when both are different this is not a corressponding Shape
'            Else
'                If Not dct.Exists(sId) Then
'                    dct.Add sId, shp
'                End If
'            End If
'        ElseIf c_quality = enAndNameCodeName Then ' all properties are equal
'            If Not bShapeDiffers And Not bRefToDiffers And Not bScopeDiffers Then
'                If Not dct.Exists(sId) Then
'                    dct.Add sId, shp
'                End If
'            End If
'        End If
    Next shp
    
    Set CorrespondingShapes = dct
    Set c_dct_corresponding = dct
    Set dct = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function







Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncShapes." & s
End Function

Public Function Exists(ByVal x_shp As Variant, _
                       ByVal x_wsh As Worksheet, _
              Optional ByRef x_shp_result As Shape) As Boolean
' -------------------------------------------------------------------------
' When the Shape (x_shp) exists in the Worksheet (x_wsh) either under
' its name or - and in case it is a Type msoOLEControlObject Shape - under
' the OLEControlObject's Name the function returns TRUE and the found Shape
' (x_shp_result).
' -------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim shp             As Shape
    Dim sNameWsh        As String
    Dim sNameWshOOB     As String
    Dim sName           As String
    Dim sOOBCodeName    As String
                   
    If TypeName(x_shp) = "Shape" Then
        sName = x_shp.Name
    Else
        sName = x_shp
    End If
    
    For Each shp In x_wsh.Shapes
        ShapeName shp, sNameWsh, sNameWshOOB
        If sNameWsh = sName Then
            Set x_shp_result = shp
            Exists = True
            Exit For
        Else
            If TypeName(x_shp) = "Shape" Then
                If x_shp.Type = msoOLEControlObject Then
                    If shp.OLEFormat.Object.Name = sOOBCodeName Then
                        Set x_shp_result = shp
                        Exists = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next shp
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function GetShape(ByVal g_shp_id As String, _
                          ByVal g_wbk As Workbook, _
                          ByRef g_wsh As Worksheet) As Shape
' ------------------------------------------------------------------------------
' Returns the Shape object (g_shp_id) and the Worksheet (g_wsh) it has been
' found.
' ------------------------------------------------------------------------------
    Dim shp As Shape
    Dim wsh As Worksheet
    
    For Each wsh In g_wbk.Worksheets
        For Each shp In wsh.Shapes
            If SyncId(shp) = g_shp_id Then
                Set g_wsh = wsh
                Set GetShape = shp
                Exit For
            End If
        Next shp
    Next wsh
                          
End Function
          
Public Sub Initialize()
    Set dctKnownChanged = Nothing
    Set dctKnownNew = Nothing
    Set dctKnownObsolete = Nothing
End Sub

Private Function IsNew(ByVal in_shp_source As Shape, _
                       ByVal in_wsh_target As Worksheet) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "IsNew"
    
    On Error GoTo eh
    IsNew = Not mSyncShapes.Exists(in_shp_source, in_wsh_target)

xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsObsolete(ByVal in_shp_target As Shape, _
                            ByVal in_wsh_source As Worksheet) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "IsObsolete"
    
    On Error GoTo eh
    IsObsolete = Not mSyncShapes.Exists(in_shp_target, in_wsh_source)

xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function MaxLenShapeName(ByVal wbk_target As Workbook, _
                                 ByVal wbk_source As Workbook) As Long
    Dim wsh     As Worksheet
    Dim shp     As Shape
    Dim lMaxLen As Long
    
    For Each wsh In wbk_source.Worksheets
        For Each shp In wsh.Shapes
            lMaxLen = Max(lMaxLen, Len(wsh.Name & "." & shp.Name))
        Next shp
    Next wsh
    
    For Each wsh In wbk_target.Worksheets
        For Each shp In wsh.Shapes
            lMaxLen = Max(lMaxLen, Len(wsh.Name & "." & shp.Name))
        Next shp
    Next wsh
    MaxLenShapeName = lMaxLen
    
End Function

Private Function RunRemoveAsserted(ByVal sync_shp_name As String, _
                                   ByVal sync_wsh As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Shape (sync_shp_name) not exists in Sheet (sync_wsh).
' ------------------------------------------------------------------------------
    Dim shp As Shape
    
    RunRemoveAsserted = True
    For Each shp In sync_wsh.Shapes
        If shp.Name = sync_shp_name Then
            RunRemoveAsserted = False
            Exit For
        End If
    Next shp

End Function

Private Function ShapeName(ByVal sn_shp As Shape, _
                           ByRef sn_shp_name As String, _
                           ByRef sn_oob_code_name As String) As String
' ------------------------------------------------------------------------------
' Returns the Name of a Shape object (sn-shp_name). In case the Shape is a type
' msoOLEControlObject the OOB-Object Name is returned (sn_shp_oob), else a
' vbNullString.
' ------------------------------------------------------------------------------

    sn_shp_name = sn_shp.Name
    ShapeName = sn_shp_name
    If sn_shp.Type = msoOLEControlObject _
    Then sn_oob_code_name = sn_shp.OLEFormat.Object.Name

End Function

Public Function ShapeNames(ByVal sn_obj As Variant) As String
' ------------------------------------------------------------------------------
' Returns the Name of a Shape and - in case Shape is a type msoOLEControlObject
' - the OOB-Object Name (Code Name) added separated with a semicolon.
' ------------------------------------------------------------------------------
    Const PROC = "ShapeNames"
    
    On Error GoTo eh
    Dim shp As Shape
    Dim oob As OLEObject
    
    Select Case TypeName(sn_obj)
        Case "Shape"
            Set shp = sn_obj
            ShapeNames = shp.Name
            If shp.Type = msoOLEControlObject Then ShapeNames = ShapeNames & " (" & shp.OLEFormat.Object.Name & ")"
        Case "OLEObject"
            Set oob = sn_obj
            ShapeNames = oob.ShapeRange.Name
            ShapeNames = ShapeNames & " (" & oob.Name & ")"
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "The provided object iss neither a 'Shape' nor an 'OleObject'!"
    End Select
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SyncKind()
' ------------------------------------------------------------------------------
' Called by mSync.RunSync: Collects to be synchronized Sheet Controls and
' displays them in a mode-less dialog for being confirmed one by one.
' ------------------------------------------------------------------------------
    Const PROC = "SyncKind"
    
    On Error GoTo eh
    Dim AppRunArgs  As New Dictionary
    Dim cllButtons  As New Collection
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim v           As Variant
    Dim dctNew      As Dictionary
    Dim dctObsolete As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    
    mSync.MonitorStep "Synchronizing Sheet Shapes"
    mService.MessageUnload TITLE_SYNC_SHAPES
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_SHAPES)
    With Msg.Section(1)
        .Label.Text = "Obsolete Shapes:"
        .Label.FontColor = rgbDarkGreen
        .Text.MonoSpaced = True
        For Each v In dctObsolete
            .Text.Text = .Text.Text & vbLf & v & dctObsolete(v)
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(2)
        .Label.Text = "New Shapes:"
        .Label.FontColor = rgbDarkGreen
        .Text.MonoSpaced = True
        For Each v In dctNew
            .Text.Text = .Text.Text & vbLf & v & dctNew(v)
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(3)
        .Label.Text = "About Shape synchronization:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Properties of the Shape - new or existing - are synchronized when changed. " & vbLf & _
                     "When the Shape is a Type msoOLEControlObject the OOB's properties are synchronized " & _
                     "in addition."
    End With
    With Msg.Section(8).Label
        .Text = "See in README chapter CompMan's VB-Project-Synchronization service (GitHub README):"
        .FontColor = rgbBlue
        .OpenWhenClicked = mCompMan.README_URL & mCompMan.README_SYNC_CHAPTER
    End With
               
    '~~ Prepare a Command-Buttonn with an Application.Run action for the synchronization of all Worksheets
    Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN)
    mMsg.BttnAppRun AppRunArgs, SYNC_ALL_BTTN _
                                , ThisWorkbook _
                                , "mSyncShapes.AppRunSyncAll"
               
    '~~ Display the mode-less dialog for the confirmation which Sheet synchronization to run
    mMsg.Dsply dsply_title:=TITLE_SYNC_SHAPES _
             , dsply_msg:=Msg _
             , dsply_buttons:=cllButtons _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=AppRunArgs _
             , dsply_width_min:=45 _
             , dsply_pos:=DialogTop & ";" & DialogLeft

xt: mBasic.BoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function SyncId(ByVal shp As Shape) As String
' ------------------------------------------------------------------------------
' Unified synchronization Id.
' ------------------------------------------------------------------------------
    SyncId = shp.Parent.Name & "." & shp.Name & " (" & TypeString(shp) & ")"
End Function

Private Function SyncNewAsserted(ByVal sna_shp_source As Shape, _
                                 ByVal sna_wsh_target As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when Name and Type of target shape and source shape are the same.
' ------------------------------------------------------------------------------
    Dim oobSource   As OLEObject
    Dim oobTarget   As OLEObject
    Dim shpTarget   As Shape
    Dim wbkTarget   As Workbook
    
    Set wbkTarget = sna_wsh_target.Parent
    Set shpTarget = GetShape(SyncId(sna_shp_source), wbkTarget, sna_wsh_target)
    If Not shpTarget Is Nothing Then
        If shpTarget.Type = msoOLEControlObject Then
            Set oobSource = sna_shp_source.OLEFormat.Object
            Set oobTarget = shpTarget.OLEFormat.Object
            SyncNewAsserted = oobTarget.Name = oobSource.Name _
                         And oobTarget.OLEType = oobSource.OLEType
        Else
            SyncNewAsserted = shpTarget.Name = sna_shp_source.Name _
                          And shpTarget.Type = sna_shp_source.Type
        End If
    End If
    
End Function

Public Function TypeString(ByVal shp As Shape) As String
' ------------------------------------------------------------------------------
' Returns the Shape's (shp) Type as string.
' ------------------------------------------------------------------------------
    Const PROC = "TypeString"
        
    On Error GoTo eh
    Dim oob As OLEObject
    
    If shp.Type = msoOLEControlObject Then Set oob = shp.OLEFormat.Object
    
    Select Case shp.Type
'        Case mso3DModel:                TypeString = "3dModel"
        Case msoAutoShape:              TypeString = "AutoShape " & TypeStringAutoShape(shp)
        Case msoCallout:                TypeString = "CallOut"
        Case msoCanvas:                 TypeString = "Canvas"
        Case msoChart:                  TypeString = "Chart"
        Case msoComment:                TypeString = "Comment"
'        Case msoContentApp:             TypeString = "ContentApp"
        Case msoDiagram:                TypeString = "Diagram"
        Case msoEmbeddedOLEObject:      TypeString = "EmbeddedOLEObject"
        Case msoFormControl:            TypeString = "FormControl " & TypeStringFormControl(shp)
        Case msoFreeform:               TypeString = "Freeform"
'        Case msoGraphic:                TypeString = "Graphic"
        Case msoGroup:                  TypeString = "Group"
        Case msoInk:                    TypeString = "Ink"
        Case msoInkComment:             TypeString = "InkComment"
        Case msoLine:                   TypeString = "Line"
'        Case msoLinked3DModel:          TypeString = "Linked3DModel"
'        Case msoLinkedGraphic:          TypeString = "LinkedGraphic"
        Case msoLinkedOLEObject:        TypeString = "LinkedOLEObject"
        Case msoLinkedPicture:          TypeString = "LinkedPicture"
        Case msoMedia:                  TypeString = "Media"
        Case msoOLEControlObject:       TypeString = "ActiveX-" & TypeName(oob.Object)
        Case msoPicture:                TypeString = "Picture"
        Case msoPlaceholder:            TypeString = "Placeholder"
        Case msoScriptAnchor:           TypeString = "ScriptAnchor"
        Case msoShapeTypeMixed:         TypeString = "ShapeTypeMixed"
        Case msoSlicer:                 TypeString = "Slicer"
        Case msoTable:                  TypeString = "Table"
        Case msoTextBox:                TypeString = "TextBox"
        Case msoTextEffect:             TypeString = "TextEffect"
'        Case msoWebVideo:               TypeString = "WebVideo"
        Case Else
            Debug.Print "Shape-Type: '" & shp.Type & "' Not implemented"
    End Select

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function TypeStringAutoShape(ByVal shp As Shape) As String
 
    Select Case shp.AutoShapeType
        Case msoShape10pointStar:                          TypeStringAutoShape = "10-point star"
        Case msoShape12pointStar:                          TypeStringAutoShape = "12-point star"
        Case msoShape16pointStar:                          TypeStringAutoShape = "16-point star"
        Case msoShape24pointStar:                          TypeStringAutoShape = "24-point star"
        Case msoShape32pointStar:                          TypeStringAutoShape = "32-point star"
        Case msoShape4pointStar:                           TypeStringAutoShape = "4-point star"
        Case msoShape5pointStar:                           TypeStringAutoShape = "5-point star"
        Case msoShape6pointStar:                           TypeStringAutoShape = "6-point star"
        Case msoShape7pointStar:                           TypeStringAutoShape = "7-point star"
        Case msoShape8pointStar:                           TypeStringAutoShape = "8-point star"
        Case msoShapeActionButtonBackorPrevious:           TypeStringAutoShape = "Back or Previous button"
        Case msoShapeActionButtonBeginning:                TypeStringAutoShape = "Beginning button"
        Case msoShapeActionButtonCustom:                   TypeStringAutoShape = "Button with no default picture or text"
        Case msoShapeActionButtonDocument:                 TypeStringAutoShape = "Document button"
        Case msoShapeActionButtonEnd:                      TypeStringAutoShape = "End button"
        Case msoShapeActionButtonForwardorNext:            TypeStringAutoShape = "Forward or Next button"
        Case msoShapeActionButtonHelp:                     TypeStringAutoShape = "Help button"
        Case msoShapeActionButtonHome:                     TypeStringAutoShape = "Home button"
        Case msoShapeActionButtonInformation:              TypeStringAutoShape = "Information button"
        Case msoShapeActionButtonMovie:                    TypeStringAutoShape = "Movie button"
        Case msoShapeActionButtonReturn:                   TypeStringAutoShape = "Return button"
        Case msoShapeActionButtonSound:                    TypeStringAutoShape = "Sound button"
        Case msoShapeArc:                                  TypeStringAutoShape = "Arc"
        Case msoShapeBalloon:                              TypeStringAutoShape = "Balloon"
        Case msoShapeBentArrow:                            TypeStringAutoShape = "Block arrow that follows a curved 90-degree angle."
        Case msoShapeBentUpArrow:                          TypeStringAutoShape = "Block arrow that follows a sharp 90-degree angle. Points up by default."
        Case msoShapeBevel:                                TypeStringAutoShape = "Bevel"
        Case msoShapeBlockArc:                             TypeStringAutoShape = "Block arc"
        Case msoShapeCan:                                  TypeStringAutoShape = "Can"
        Case msoShapeChartPlus:                            TypeStringAutoShape = "Square divided vertically and horizontally into four quarters"
        Case msoShapeChartStar:                            TypeStringAutoShape = "Square divided into six parts along vertical and diagonal lines"
        Case msoShapeChartX:                               TypeStringAutoShape = "Square divided into four parts along diagonal lines"
        Case msoShapeChevron:                              TypeStringAutoShape = "Chevron"
        Case msoShapeChord:                                TypeStringAutoShape = "Circle with a line connecting two points on the perimeter through the interior of the circle; a circle with a chord"
        Case msoShapeCircularArrow:                        TypeStringAutoShape = "Block arrow that follows a curved 180-degree angle"
        Case msoShapeCloud:                                TypeStringAutoShape = "Cloud shape"
        Case msoShapeCloudCallout:                         TypeStringAutoShape = "Cloud callout"
        Case msoShapeCorner:                               TypeStringAutoShape = "Rectangle with rectangular-shaped hole."
        Case msoShapeCornerTabs:                           TypeStringAutoShape = "Four right triangles aligning along a rectangular path; four 'snipped' corners."
        Case msoShapeCross:                                TypeStringAutoShape = "Cross"
        Case msoShapeCube:                                 TypeStringAutoShape = "Cube"
        Case msoShapeCurvedDownArrow:                      TypeStringAutoShape = "Block arrow that curves down"
        Case msoShapeCurvedDownRibbon:                     TypeStringAutoShape = "Ribbon banner that curves down"
        Case msoShapeCurvedLeftArrow:                      TypeStringAutoShape = "Block arrow that curves left"
        Case msoShapeCurvedRightArrow:                     TypeStringAutoShape = "Block arrow that curves right"
        Case msoShapeCurvedUpArrow:                        TypeStringAutoShape = "Block arrow that curves up"
        Case msoShapeCurvedUpRibbon:                       TypeStringAutoShape = "Ribbon banner that curves up"
        Case msoShapeDecagon:                              TypeStringAutoShape = "Decagon"
        Case msoShapeDiagonalStripe:                       TypeStringAutoShape = "Rectangle with two triangles-shapes removed; a diagonal stripe"
        Case msoShapeDiamond:                              TypeStringAutoShape = "Diamond"
        Case msoShapeDodecagon:                            TypeStringAutoShape = "Dodecagon"
        Case msoShapeDonut:                                TypeStringAutoShape = "Donut"
        Case msoShapeDoubleBrace:                          TypeStringAutoShape = "Double brace"
        Case msoShapeDoubleBracket:                        TypeStringAutoShape = "Double bracket"
        Case msoShapeDoubleWave:                           TypeStringAutoShape = "Double wave"
        Case msoShapeDownArrow:                            TypeStringAutoShape = "Block arrow that points down"
        Case msoShapeDownArrowCallout:                     TypeStringAutoShape = "Callout with arrow that points down"
        Case msoShapeDownRibbon:                           TypeStringAutoShape = "Ribbon banner with center area below ribbon ends"
        Case msoShapeExplosion1:                           TypeStringAutoShape = "Explosion"
        Case msoShapeExplosion2:                           TypeStringAutoShape = "Explosion"
        Case msoShapeFlowchartAlternateProcess:            TypeStringAutoShape = "Alternate process flowchart symbol"
        Case msoShapeFlowchartCard:                        TypeStringAutoShape = "Card flowchart symbol"
        Case msoShapeFlowchartCollate:                     TypeStringAutoShape = "Collate flowchart symbol"
        Case msoShapeFlowchartConnector:                   TypeStringAutoShape = "Connector flowchart symbol"
        Case msoShapeFlowchartData:                        TypeStringAutoShape = "Data flowchart symbol"
        Case msoShapeFlowchartDecision:                    TypeStringAutoShape = "Decision flowchart symbol"
        Case msoShapeFlowchartDelay:                       TypeStringAutoShape = "Delay flowchart symbol"
        Case msoShapeFlowchartDirectAccessStorage:         TypeStringAutoShape = "Direct access storage flowchart symbol"
        Case msoShapeFlowchartDisplay:                     TypeStringAutoShape = "Display flowchart symbol"
        Case msoShapeFlowchartDocument:                    TypeStringAutoShape = "Document flowchart symbol"
        Case msoShapeFlowchartExtract:                     TypeStringAutoShape = "Extract flowchart symbol"
        Case msoShapeFlowchartInternalStorage:             TypeStringAutoShape = "Internal storage flowchart symbol"
        Case msoShapeFlowchartMagneticDisk:                TypeStringAutoShape = "Magnetic disk flowchart symbol"
        Case msoShapeFlowchartManualInput:                 TypeStringAutoShape = "Manual input flowchart symbol"
        Case msoShapeFlowchartManualOperation:             TypeStringAutoShape = "Manual operation flowchart symbol"
        Case msoShapeFlowchartMerge:                       TypeStringAutoShape = "Merge flowchart symbol"
        Case msoShapeFlowchartMultidocument:               TypeStringAutoShape = "Multi-document flowchart symbol"
        Case msoShapeFlowchartOfflineStorage:              TypeStringAutoShape = "Offline storage flowchart symbol"
        Case msoShapeFlowchartOffpageConnector:            TypeStringAutoShape = "Off-page connector flowchart symbol"
        Case msoShapeFlowchartOr:                          TypeStringAutoShape = "'Or' flowchart symbol"
        Case msoShapeFlowchartPredefinedProcess:           TypeStringAutoShape = "Predefined process flowchart symbol"
        Case msoShapeFlowchartPreparation:                 TypeStringAutoShape = "Preparation flowchart symbol"
        Case msoShapeFlowchartProcess:                     TypeStringAutoShape = "Process flowchart symbol"
        Case msoShapeFlowchartPunchedTape:                 TypeStringAutoShape = "Punched tape flowchart symbol"
        Case msoShapeFlowchartSequentialAccessStorage:     TypeStringAutoShape = "Sequential access storage flowchart symbol"
        Case msoShapeFlowchartSort:                        TypeStringAutoShape = "Sort flowchart symbol"
        Case msoShapeFlowchartStoredData:                  TypeStringAutoShape = "Stored data flowchart symbol"
        Case msoShapeFlowchartSummingJunction:             TypeStringAutoShape = "Summing junction flowchart symbol"
        Case msoShapeFlowchartTerminator:                  TypeStringAutoShape = "Terminator flowchart symbol"
        Case msoShapeFoldedCorner:                         TypeStringAutoShape = "Folded corner"
        Case msoShapeFrame:                                TypeStringAutoShape = "Rectangular picture frame"
        Case msoShapeFunnel:                               TypeStringAutoShape = "Funnel"
        Case msoShapeHalfFrame:                            TypeStringAutoShape = "Half of a rectangular picture frame"
        Case msoShapeHeart:                                TypeStringAutoShape = "Heart"
        Case msoShapeHeptagon:                             TypeStringAutoShape = "Heptagon"
        Case msoShapeHexagon:                              TypeStringAutoShape = "Hexagon"
        Case msoShapeHorizontalScroll:                     TypeStringAutoShape = "Horizontal scroll"
        Case msoShapeIsoscelesTriangle:                    TypeStringAutoShape = "Isosceles triangle"
        Case msoShapeLeftArrow:                            TypeStringAutoShape = "Block arrow that points left"
        Case msoShapeLeftArrowCallout:                     TypeStringAutoShape = "Callout with arrow that points left"
        Case msoShapeLeftBrace:                            TypeStringAutoShape = "Left brace"
        Case msoShapeLeftBracket:                          TypeStringAutoShape = "Left bracket"
        Case msoShapeLeftCircularArrow:                    TypeStringAutoShape = "Circular arrow pointing counter-clockwise"
        Case msoShapeLeftRightArrow:                       TypeStringAutoShape = "Block arrow with arrowheads that point both left and right"
        Case msoShapeLeftRightArrowCallout:                TypeStringAutoShape = "Callout with arrowheads that point both left and right"
        Case msoShapeLeftRightCircularArrow:               TypeStringAutoShape = "Circular arrow pointing clockwise and counter-clockwise; a curved arrow with points at both ends"
        Case msoShapeLeftRightRibbon:                      TypeStringAutoShape = "Ribbon with an arrow at both ends"
        Case msoShapeLeftRightUpArrow:                     TypeStringAutoShape = "Block arrow with arrowheads that point left, right, and up"
        Case msoShapeLeftUpArrow:                          TypeStringAutoShape = "Block arrow with arrowheads that point left and up"
        Case msoShapeLightningBolt:                        TypeStringAutoShape = "Lightning bolt"
        Case msoShapeLineCallout1AccentBar:                TypeStringAutoShape = "Callout with horizontal accent bar"
        Case msoShapeLineCallout1BorderandAccentBar:       TypeStringAutoShape = "Callout with border and horizontal accent bar"
        Case msoShapeLineCallout1NoBorder:                 TypeStringAutoShape = "Callout with horizontal line"
        Case msoShapeLineCallout2AccentBar:                TypeStringAutoShape = "Callout with diagonal callout line and accent bar"
        Case msoShapeLineCallout2BorderandAccentBar:       TypeStringAutoShape = "Callout with border, diagonal straight line, and accent bar"
        Case msoShapeLineCallout2NoBorder:                 TypeStringAutoShape = "Callout with no border and diagonal callout line"
        Case msoShapeLineCallout3AccentBar:                TypeStringAutoShape = "Callout with angled callout line and accent bar"
        Case msoShapeLineCallout3BorderandAccentBar:       TypeStringAutoShape = "Callout with border, angled callout line, and accent bar"
        Case msoShapeLineCallout3NoBorder:                 TypeStringAutoShape = "Callout with no border and angled callout line"
        Case msoShapeLineCallout4AccentBar:                TypeStringAutoShape = "Callout with accent bar and callout line segments forming a U-shape"
        Case msoShapeLineCallout4BorderandAccentBar:       TypeStringAutoShape = "Callout with border, accent bar, and callout line segments forming a U-shape"
        Case msoShapeLineCallout4NoBorder:                 TypeStringAutoShape = "Callout with no border and callout line segments forming a U-shape"
        Case msoShapeLineInverse:                          TypeStringAutoShape = "Line inverse"
        Case msoShapeMathDivide:                           TypeStringAutoShape = "Division symbol ÷"
        Case msoShapeMathEqual:                            TypeStringAutoShape = "Equivalence symbol ="
        Case msoShapeMathMinus:                            TypeStringAutoShape = "Subtraction symbol -"
        Case msoShapeMathMultiply:                         TypeStringAutoShape = "Multiplication symbol x"
        Case msoShapeMathNotEqual:                         TypeStringAutoShape = "Non-equivalence symbol ?"
        Case msoShapeMathPlus:                             TypeStringAutoShape = "Addition symbol +"
        Case msoShapeMixed:                                TypeStringAutoShape = "- Return value only; indicates a combination of the other states."
        Case msoShapeMoon:                                 TypeStringAutoShape = "Moon"
        Case msoShapeNonIsoscelesTrapezoid:                TypeStringAutoShape = "Trapezoid with asymmetrical non-parallel sides"
        Case msoShapeNoSymbol:                             TypeStringAutoShape = "'No' symbol"
        Case msoShapeNotchedRightArrow:                    TypeStringAutoShape = "Notched block arrow that points right"
        Case Not msoShapeNotPrimitive:                     TypeStringAutoShape = "Not supported"
        Case msoShapeOctagon:                              TypeStringAutoShape = "Octagon"
        Case msoShapeOval:                                 TypeStringAutoShape = "Oval"
        Case msoShapeOvalCallout:                          TypeStringAutoShape = "Oval-shaped callout"
        Case msoShapeParallelogram:                        TypeStringAutoShape = "Parallelogram"
        Case msoShapePentagon:                             TypeStringAutoShape = "Pentagon"
        Case msoShapePie:                                  TypeStringAutoShape = "Circle ('pie') with a portion missing"
        Case msoShapePieWedge:                             TypeStringAutoShape = "Quarter of a circular shape"
        Case msoShapePlaque:                               TypeStringAutoShape = "Plaque"
        Case msoShapePlaqueTabs:                           TypeStringAutoShape = "Four quarter-circles defining a rectangular shape"
        Case msoShapeQuadArrow:                            TypeStringAutoShape = "Block arrows that point up, down, left, and right"
        Case msoShapeQuadArrowCallout:                     TypeStringAutoShape = "Callout with arrows that point up, down, left, and right"
        Case msoShapeRectangle:                            TypeStringAutoShape = "Rectangle"
        Case msoShapeRectangularCallout:                   TypeStringAutoShape = "Rectangular callout"
        Case msoShapeRegularPentagon:                      TypeStringAutoShape = "Pentagon"
        Case msoShapeRightArrow:                           TypeStringAutoShape = "Block arrow that points right"
        Case msoShapeRightArrowCallout:                    TypeStringAutoShape = "Callout with arrow that points right"
        Case msoShapeRightBrace:                           TypeStringAutoShape = "Right brace"
        Case msoShapeRightBracket:                         TypeStringAutoShape = "Right bracket"
        Case msoShapeRightTriangle:                        TypeStringAutoShape = "Right triangle"
        Case msoShapeRound1Rectangle:                      TypeStringAutoShape = "Rectangle with one rounded corner"
        Case msoShapeRound2DiagRectangle:                  TypeStringAutoShape = "Rectangle with two rounded corners, diagonally-opposed"
        Case msoShapeRound2SameRectangle:                  TypeStringAutoShape = "Rectangle with two-rounded corners that share a side"
        Case msoShapeRoundedRectangle:                     TypeStringAutoShape = "Rounded rectangle"
        Case msoShapeRoundedRectangularCallout:            TypeStringAutoShape = "Rounded rectangle-shaped callout"
        Case msoShapeSmileyFace:                           TypeStringAutoShape = "Smiley face"
        Case msoShapeSnip1Rectangle:                       TypeStringAutoShape = "Rectangle with one snipped corner"
        Case msoShapeSnip2DiagRectangle:                   TypeStringAutoShape = "Rectangle with two snipped corners, diagonally-opposed"
        Case msoShapeSnip2SameRectangle:                   TypeStringAutoShape = "Rectangle with two snipped corners that share a side"
        Case msoShapeSnipRoundRectangle:                   TypeStringAutoShape = "Rectangle with one snipped corner and one rounded corner"
        Case msoShapeSquareTabs:                           TypeStringAutoShape = "Four small squares that define a rectangular shape"
        Case msoShapeStripedRightArrow:                    TypeStringAutoShape = "Block arrow that points right with stripes at the tail"
        Case msoShapeSun:                                  TypeStringAutoShape = "Sun"
        Case msoShapeSwooshArrow:                          TypeStringAutoShape = "Curved arrow"
        Case msoShapeTear:                                 TypeStringAutoShape = "Water droplet"
        Case msoShapeTrapezoid:                            TypeStringAutoShape = "Trapezoid"
        Case msoShapeUpArrow:                              TypeStringAutoShape = "Block arrow that points up"
        Case msoShapeUpArrowCallout:                       TypeStringAutoShape = "Callout with arrow that points up"
        Case msoShapeUpDownArrow:                          TypeStringAutoShape = "Block arrow that points up and down"
        Case msoShapeUpDownArrowCallout:                   TypeStringAutoShape = "Callout with arrows that point up and down"
        Case msoShapeUpRibbon:                             TypeStringAutoShape = "Ribbon banner with center area above ribbon ends"
        Case msoShapeUTurnArrow:                           TypeStringAutoShape = "Block arrow forming a U shape"
        Case msoShapeVerticalScroll:                       TypeStringAutoShape = "Vertical scroll"
        Case msoShapeWave:                                 TypeStringAutoShape = "Wave"
    End Select
    
End Function

Private Function TypeStringFormControl(ByVal shp As Shape) As String

    Select Case shp.FormControlType
        Case xlButtonControl:   TypeStringFormControl = "CommandButton"
        Case xlCheckBox:        TypeStringFormControl = "CheckBox"
        Case xlDropDown:        TypeStringFormControl = "DropDown"
        Case xlEditBox:         TypeStringFormControl = "EditBox"
        Case xlGroupBox:        TypeStringFormControl = "GroupBox"
        Case xlLabel:           TypeStringFormControl = "Label"
        Case xlListBox:         TypeStringFormControl = "ListBox"
        Case xlOptionButton:    TypeStringFormControl = "OptionButton"
        Case xlScrollBar:       TypeStringFormControl = "ScrollBar"
        Case xlSpinner:         TypeStringFormControl = "Spinner"
    End Select

End Function

