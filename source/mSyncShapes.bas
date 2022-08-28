Attribute VB_Name = "mSyncShapes"
Option Explicit
' -------------------------------------------------------------------
' Standard Module mSyncSheetControls
'          Services to synchronize new, obsolete, and properties of
'          sheet shapes and OLEObjects.
'
' -------------------------------------------------------------------
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

Public Function CorrespTargetOOB(ByVal cto_shp_source As Shape, _
                                 ByVal cto_wsh_target As Worksheet, _
                                 ByRef cto_oob_result As OLEObject) As OLEObject
' ------------------------------------------------------------------------------
' Returns the OleObject in the Sync-Target-Workbook's Worksheet which corres-
' ponds to the source Shape's (cto_shp_source) OOB.
' ------------------------------------------------------------------------------
    Dim oob             As OLEObject
    Dim sShapeName      As String
    Dim sOOBCodeName    As String
    
    mSyncShapes.ShapeName cto_shp_source, sShapeName, sOOBCodeName
    
    For Each oob In cto_wsh_target.OLEObjects
        If oob.Name = sOOBCodeName Or oob.ShapeRange.Name = sShapeName Then
            '~~ Note: The names (shape-name and code-name) may not have been synchronized yet
            Set cto_oob_result = oob
            Set CorrespTargetOOB = oob
            Exit For
        End If
    Next oob

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

Public Sub CollectAllItems()
' ------------------------------------------------------------------------------
' Writes: - the Worksheet Shapes potentially synchronized and
'         - the Shape-Properties able to be synchronized (writeable)
'         to the wsSynch sheet.
' ------------------------------------------------------------------------------
    Const PROC = "CollectAllItems"
    
    On Error GoTo eh
    Dim shp         As Shape
    Dim wsh         As Worksheet
    Dim ShapeId     As String
    Dim v           As Variant
    Dim dctShapes   As New Dictionary
    Dim dctPrprtys  As New Dictionary
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim wshTarget   As Worksheet
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    wsService.MaxLenShapeName = MaxLenShapeName(wbkTarget, wbkSource)

    For Each wsh In wbkSource.Worksheets
        mSyncSheets.CorrespSheet wsh, wbkTarget, wshTarget
        If Not wshTarget Is Nothing Then
            For Each shp In wsh.Shapes
                With shp
                    If InStr(.Name, " ") Then
                        Log.ServicedItem = shp
                        Log.Entry = "Will not be synchronized (shape does not have a user specified name)!"
                        Debug.Print "shp.Name: " & ItemSyncName(shp) & " (not synchronized)"
                        GoTo ns1  ' Comments are not synced as shapes
                    ElseIf .Type = msoComment Then
                        Log.ServicedItem = shp
                        Log.Entry = "Type of Shape intentionally not synchronized!"
                        Debug.Print "shp.Name: " & ItemSyncName(shp) & " (not synchronized by intention)"
                        GoTo ns1  ' next shape
                    End If
                    ShapeId = ItemSyncName(shp)
                    If Not dctShapes.Exists(ShapeId) Then
                        Log.ServicedItem = shp
                        mDct.DctAdd dctShapes, ShapeId, shp, order_bykey, seq_ascending
'                        mSyncShapePrprtys.PropertiesWriteable shp, dctPrprtys
                    End If
                End With
ns1:         Next shp
        End If
    Next wsh
            
    For Each v In dctShapes
        wsSync.ShpItemAll(v) = True
    Next
    
    Set dctShapes = Nothing
    Set dctPrprtys = Nothing

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function CollectChangedStillRelevant(ByVal ir_sheet_shape As String) As Boolean
    With wsSync
        CollectChangedStillRelevant = IsSyncItem(ir_sheet_shape) _
                         And Not .ShpItemChangedDone(ir_sheet_shape) _
                         And Not .ShpItemChangedFailed(ir_sheet_shape)
    End With
End Function

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

Public Function CollectNew() As Dictionary
' ------------------------------------------------------------------------------
' Returns a collection of all those sheet controls which exist in the source
' but not in the target Workbook, for each Shape the Application.Run arguments:
' ------------------------------------------------------------------------------
    Const PROC = "CollectNew"
    
    On Error GoTo eh
    Dim dct             As New Dictionary
    Dim shpSource       As Shape
    Dim wshSource       As Worksheet
    Dim wshTarget       As Worksheet
    Dim v               As Variant
    Dim TargetSheetName As String
    Dim wbkTarget       As Workbook
    Dim wbkSource       As Workbook
    Dim sProgress       As String
    
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    For Each wshSource In wbkSource.Worksheets
        If Not mSyncSheets.CorrespSheet(wshSource, wbkTarget, wshTarget) Is Nothing Then
            '~~ There is a corresponding sheet with an equal Name or CodeName
            For Each shpSource In wshSource.Shapes
                mSync.MonitorStep "Collecting Sheet Shapes new" & sProgress
                If IsNew(shpSource, wshTarget) Then
                    Log.ServicedItem = shpSource
                    mDct.DctAdd dct, ItemSyncName(shpSource), " (new " & TypeString(shpSource) & ")", order_bykey, seq_ascending, sense_casesensitive
                End If
                sProgress = sProgress & "."
ns:         Next shpSource
        End If
    Next wshSource
    
    If wsSync.ShpNumberNew = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.ShpItemNew(v) = True
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
          
Private Function CollectNewStillRelevant(ByVal ir_sheet_shape As String) As Boolean
    With wsSync
        If IsSyncItem(ir_sheet_shape) Then
            CollectNewStillRelevant = Not .ShpItemNewDone(ir_sheet_shape) _
                                  And Not .ShpItemNewFailed(ir_sheet_shape)
'            If ir_sheet_shape = "Test_A.cmbSheetTestA1" Then Stop
        End If
    End With
End Function

Public Function CollectObsolete() As Dictionary
' ------------------------------------------------------------------------------
' Returns a collection of all those sheet controls at least on property has
' changed.
' ------------------------------------------------------------------------------
    Const PROC = "CollectObsolete"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim dct         As New Dictionary
    Dim ShapeId     As String
    Dim shpTarget   As Shape
    Dim v           As Variant
    Dim wshTarget   As Worksheet
    Dim wshSource   As Worksheet
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim sProgress   As String
    
    sProgress = " "
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
   
    For Each wshTarget In wbkTarget.Worksheets
        If Not mSyncSheets.CorrespSheet(wshTarget, wbkSource, wshSource) Is Nothing Then
            For Each shpTarget In wshTarget.Shapes
                mSync.MonitorStep "Collecting Sheet Shapes obsolete" & sProgress
                If shpTarget.Name Like "Comment *" Then GoTo ns ' next shape
                If IsObsolete(shpTarget, wshSource) Then
                    '~~ the shape not exists in the corresponding source sheet
                    ShapeId = ItemSyncName(shpTarget)
                    mDct.DctAdd dct, ShapeId, " (obsolete " & TypeString(shpTarget), order_bykey, seq_ascending, sense_casesensitive
                End If
                sProgress = sProgress & "."
                
ns:         Next shpTarget
        End If
    Next wshTarget

    If wsSync.ShpNumberObsolete = 0 Then
        '~~ Write not yet registered itens to wsSync sheet
        For Each v In dct
            wsSync.ShpItemObsolete(v) = True
        Next v
    End If

xt: Set CollectObsolete = dct
    Set cll = Nothing
    Set dct = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CollectObsoleteStillRelevant(ByVal ir_sheet_shape As String) As Boolean
    With wsSync
        CollectObsoleteStillRelevant = IsSyncItem(ir_sheet_shape) _
                         And Not .ShpItemObsoleteDone(ir_sheet_shape) _
                         And Not .ShpItemObsoleteFailed(ir_sheet_shape)
    End With
End Function

Private Sub CopyShape(ByVal shp_source As Shape, _
                      ByVal wsh_target As Worksheet, _
                      ByRef shp_target As Shape)
' ------------------------------------------------------------------------------
' Copies the Shape (shp_source) to Worksheet (wsh_target) and returns the target
' Shape (shp_target). The procedure tries to place the  Shape on the target
' sheet at the same cell (row/column) as the source Shape. In case the copy had
' failed the returned Shape (shp_target) is Nothing!
' ------------------------------------------------------------------------------
    Dim rng As Range
    
    Set rng = wsh_target.Cells(shp_source.TopLeftCell.row, shp_source.TopLeftCell.Column)
    shp_source.Copy
    wsh_target.Paste rng
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
    Dim oobSource   As OLEObject
    Dim oobTarget   As OLEObject
    Dim sSheetShape As String
    Dim wshSource   As Worksheet
    Dim shpTarget   As Shape
    Dim iBefore     As Long
    Dim iAfter      As Long
    
    Set wshSource = shp_source.Parent
    sSheetShape = ItemSyncName(shp_source)
    
    CopyShape shp_source, wsh_target, shpTarget
    
    If SyncNewAsserted(shp_source, wsh_target, wshSource, shpTarget) Then
        Log.Entry = "Copied from 'Sync-Source-Workbook'"
        wsSync.ShpItemNewDone(sSheetShape) = True
    Else
        wsSync.ShpItemNewFailed(sSheetShape) = True
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function CorrespTargetShape(ByVal cr_shp_source As Shape, _
                                   ByVal cr_wsh_target As Worksheet, _
                          Optional ByRef cr_shp_result As Shape) As Shape
' ------------------------------------------------------------------------------
' Returns the Shape in the Worksheet (cr_wsh_target) which corresponds to the
' source Shape (cr_shp_source). I.e. the Shape with either the same Name or -
' in case of an ActiveX-Control - the object Name as object (cr_shp_result).
' ------------------------------------------------------------------------------
    Set CorrespTargetShape = Nothing
    If Exists(cr_shp_source, cr_wsh_target, cr_shp_result) Then
        Set CorrespTargetShape = cr_shp_result
    End If
End Function
                              
Public Function Done(ByRef sync_new As Dictionary, _
                     ByRef sync_obsolete As Dictionary) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when there are no more sheet synchronizations outstanding plus
' the collections of the outstanding items.
' ----------------------------------------------------------------------------
    Set sync_new = mSyncShapes.CollectNew
    Set sync_obsolete = mSyncShapes.CollectObsolete
    
    Done = sync_new.Count _
         + sync_obsolete.Count = 0
    If Done Then mMsg.MsgInstance TITLE_SYNC_SHEET_SHAPES, True
    
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncShapes." & s
End Function

Public Function Exists(ByVal xst_shp As Variant, _
                       ByVal xst_wsh As Worksheet, _
              Optional ByRef xst_shp_result As Shape) As Boolean
' -------------------------------------------------------------------------
' When the Shape (xst_shp) exists in the Worksheet (xst_wsh) either under
' its name or - and in case it is a Type msoOLEControlObject Shape - under
' the OLEControlObject's Name the function returns TRUE and the found Shape
' (xst_shp_result).
' -------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim shp             As Shape
    Dim wbkTarget       As Workbook
    Dim wbkSource       As Workbook
    Dim sNameWsh        As String
    Dim sNameWshOOB     As String
    Dim sName           As String
    Dim sOOBCodeName    As String
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
                
    If TypeName(xst_shp) = "Shape" Then
        sName = xst_shp.Name
    Else
        sName = xst_shp
    End If
    
    For Each shp In xst_wsh.Shapes
        ShapeName shp, sNameWsh, sNameWshOOB
        If sNameWsh = sName Then
            Set xst_shp_result = shp
            Exists = True
            Exit For
        Else
            If TypeName(xst_shp) = "Shape" Then
                If xst_shp.Type = msoOLEControlObject Then
                    If shp.OLEFormat.Object.Name = sOOBCodeName Then
                        Set xst_shp_result = shp
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

Private Function GetShape(ByVal shp_name As String, _
                          ByVal shp_sheet As Worksheet) As Shape
' ------------------------------------------------------------------------------
' Returns the Shape object in Worksheet (shp_sheet) nsmed (shp_name).
' ------------------------------------------------------------------------------
    Dim shp As Shape
    
    For Each shp In shp_sheet.Shapes
        If shp.Name = shp_name Then
            Set GetShape = shp
            Exit For
        End If
    Next shp
                          
End Function

Private Function GetSheet(ByVal gs_wsh As Variant, _
                          ByVal gs_wbk As Workbook) As Worksheet
' ------------------------------------------------------------------------------
' Returns the Worksheet object in Workbook (gs_wb) which corresponds to Sheet
' (gs_ws).
' I.e. either the Name or the CodeName is identical.
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet
    
    For Each wsh In gs_wbk.Worksheets
        If IsObject(gs_wsh) Then
            If wsh.Name = gs_wsh.Name Or wsh.CodeName = gs_wsh.CodeName Then
                Set GetSheet = wsh
                Exit For
            End If
        Else
            If wsh.Name = gs_wsh Then
                Set GetSheet = wsh
                Exit For
            End If
        End If
    Next wsh
                          

End Function

Private Function IsSyncItem(ByVal ir_sheet_shape As String) As Boolean
    With wsSync
        IsSyncItem = .ShpNumberAll > 0 _
              And .ShpItemAll(ir_sheet_shape)
    End With
End Function

Public Sub SyncAllShapes()
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Synchronizes all Sheets Shapes by
' removing obsolete, adding new, and changing the properties when changed.
' Precondition: All sheet synchronizations are done.
' ------------------------------------------------------------------------------
    Const PROC = "SyncAllShapes"
    
    On Error GoTo eh
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim wshSource   As Worksheet
    Dim wshTarget   As Worksheet
    Dim shpSource   As Shape
    Dim shpTarget   As Shape
    Dim sShapeName  As String
    Dim rTarget     As Range
    Dim dctPrprtys  As New Dictionary
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    If Not wsSync.WshSyncDone _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Shape synchronization cannot be done when there are Worksheet " & _
                                            "synchronizations yet not done!"
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    lPropertyMaxLen = mSyncShapePrprtys.PropertyMaxLen
    
    '~~ Synchronize, respectively remove obsolete Shapes
    For Each wshTarget In wbkTarget.Worksheets
        mSyncSheets.CorrespSheet wshTarget, wbkSource, wshSource
        If Not wshSource Is Nothing Then
            For Each shpTarget In wshTarget.Shapes
                If IsObsolete(shpTarget, wshSource) Then
                    Log.ServicedItem = shpTarget
                    sShapeName = ItemSyncName(shpTarget)
                    shpTarget.Delete
                    Log.Entry = "Obsolete (deleted)"
                    wsSync.ShpItemObsoleteDone(sShapeName) = True
                End If
            Next shpTarget
        End If
    Next wshTarget

    '~~ Synchronize new Shapes
    For Each wshSource In wbkSource.Worksheets
        mSyncSheets.CorrespSheet wshSource, wbkTarget, wshTarget
        If Not wshTarget Is Nothing Then
            For Each shpSource In wshSource.Shapes
                If IsNew(shpSource, wshTarget) Then
                    '~~ Synchronize new Shapes
                    Log.ServicedItem = shpSource
                    CopyShapeToTarget shpSource, wshTarget
                End If
                '~~ Synchronize the properties which have changed of all Shapes
ns:         Next shpSource
        End If
    Next wshSource
        
    '~~ Synchronize Shape/OOB Properties
    For Each wshSource In wbkSource.Worksheets
        mSyncSheets.CorrespSheet wshSource, wbkTarget, wshTarget
        For Each shpSource In wshSource.Shapes
            mSyncShapePrprtys.ShapeSource = shpSource
            mSyncShapePrprtys.ShapeTarget = mSyncShapes.CorrespTargetShape(shpSource, wshTarget)
            mSyncShapePrprtys.SyncProperties dctPrprtys
        Next shpSource
    Next wshSource
    
    wsSync.ShpSyncDone = True
    For Each v In dctPrprtys
        wsSync.ShpPrprtys(v) = True
        wsSync.ShpPrprtysSyncExample(dctPrprtys(v)) = True
    Next v
    
    '~~ Re-display the synchronization dialog for still to be synchronized items
    UnloadSyncMessage TITLE_SYNC_SHEET_SHAPES
    mSync.RunSync

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub



Private Function SyncNewAsserted(ByVal sna_shp_source As Shape, _
                                 ByVal sna_wsh_target As Worksheet, _
                                 ByVal sna_wsh_source As Worksheet, _
                        Optional ByRef sna_shp_source_target As Shape) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when Name and Type of target shape and source shape are the same.
' ------------------------------------------------------------------------------
    Dim oobSource   As OLEObject
    Dim oobTarget   As OLEObject
    Dim shpTarget   As Shape
    
    Set shpTarget = GetShape(sna_shp_source.Name, sna_wsh_target)
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

Public Sub RunRemove(ByVal sync_shp_target_name As String, _
                     ByVal sync_wsh_target_name As String)
' ------------------------------------------------------------------------------
' Called via Application.Run by CommonButton: Removes a shape (sync_shp_target)
' identified by its Name property, from a Worksheet (sync_wsh_target).
' ------------------------------------------------------------------------------
    Const PROC = "RunRemove"
    
    On Error GoTo eh
    Dim shpTarget   As Shape
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    Dim wshTarget   As Worksheet
    
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    mService.EstablishServiceLog wbkTarget, mCompManClient.SRVC_SYNCHRONIZE
    
    Set wshTarget = GetSheet(sync_wsh_target_name, wbkTarget)
    Set shpTarget = GetShape(sync_shp_target_name, wshTarget)
    Log.ServicedItem = shpTarget
    shpTarget.Delete
    If Err.Number = 0 Then
        If RunRemoveAsserted(sync_shp_target_name, wshTarget) Then
            Log.Entry = "Removed!"
            wsSync.ShpItemObsoleteDone(ItemSyncName(shpTarget)) = True
        Else
            Stop
        End If
    Else
        Log.Entry = "Remove failed!"
        wsSync.ShpItemObsoleteFailed(ItemSyncName(shpTarget)) = True
    End If
    
    With wsService
        .SyncDialogLeft = mMsg.MsgInstance(.SyncDialogTitle).Left
        .SyncDialogTop = mMsg.MsgInstance(.SyncDialogTitle).Top
    End With
    
    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
    mSync.RunSync

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

Public Function ShapeName(ByVal sn_shp As Shape, _
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

Private Function ShapeType(ByVal go_shp As Shape) As String
' ------------------------------------------------------------------------------
' Returns the Name of a Shape object. In case the Shape is a type
' msoOLEControlObject not the Name but the code-name is returned.
' ------------------------------------------------------------------------------
    Dim oob As OLEObject
    
    Select Case go_shp.Type
        Case msoOLEControlObject
            Set oob = go_shp.OLEFormat.Object
            ShapeType = oob.OLEType
        Case Else
            ShapeType = go_shp.Type
    End Select

End Function

Public Function ItemSyncName(ByVal shp As Shape) As String
' ------------------------------------------------------------------------------
' Unified name for loging
' ------------------------------------------------------------------------------
    ItemSyncName = Align(shp.Parent.Name & "." & ShapeNames(shp), wsService.MaxLenShapeName) & " " & TypeString(shp)
End Function

Public Sub Sync(ByRef sync_new As Dictionary, _
                ByRef sync_obsolete As Dictionary)
' ------------------------------------------------------------------------------
' Called by mSync.RunSync: Collects to be synchronized Sheet Controls and
' displays them in a mode-less dialog for being confirmed one by one.
' ------------------------------------------------------------------------------
    Const PROC = "Sync"
    
    On Error GoTo eh
    Dim AppRunArgs  As New Dictionary
    Dim cllButtons  As New Collection
    Dim fSync       As fMsg
    Dim i           As Long
    Dim Msg         As TypeMsg
    Dim v           As Variant
    Dim wbkTarget   As Workbook
    Dim wbkSource   As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbkTarget = mWbk.GetOpen(wsService.SyncTargetWorkbookName)
    Set wbkSource = mWbk.GetOpen(wsService.SyncSourceWorkbookName)
    
    wsService.SyncDialogTitle = TITLE_SYNC_SHEET_SHAPES
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_SHEET_SHAPES)
    With Msg.Section(1)
        .Label.Text = "Obsolete Shapes:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_obsolete
            .Text.Text = .Text.Text & vbLf & v & sync_obsolete(v)
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(2)
        .Label.Text = "New Shapes:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        For Each v In sync_new
            .Text.Text = .Text.Text & vbLf & v & sync_new(v)
        Next v
        .Text.Text = Replace(.Text.Text, vbLf, vbNullString, 1, 1)
    End With
    With Msg.Section(3)
        .Label.Text = "About Shape synchronization:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Properties of the Shape - new or existing - are synchronized when changed. " & vbLf & _
                     "When the Shape is a Type msoOLEControlObject the OOB's properties are synchronized " & _
                     "in addition."
    End With
               
    '~~ Prepare a Command-Buttonn with an Application.Run action for the synchronization of all Worksheets
    Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN)
    mMsg.ButtonAppRun AppRunArgs, SYNC_ALL_BTTN _
                                , ThisWorkbook _
                                , "mSyncShapes.SyncAllShapes"
               
    '~~ Display the mode-less dialog for the confirmation which Sheet synchronization to run
    mMsg.Dsply dsply_title:=TITLE_SYNC_SHEET_SHAPES _
             , dsply_msg:=Msg _
             , dsply_buttons:=cllButtons _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=AppRunArgs _
             , dsply_width_min:=45 _
             , dsply_pos:=wsService.SyncDialogTop & ";" & wsService.SyncDialogLeft

xt: mBasic.BoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

Public Function TypeString(ByVal shp As Shape) As String
' ------------------------------------------------------------------------------
' Returns the Shape's (shp) Type as string.
' ------------------------------------------------------------------------------
    Const PROC = "TypeString"
        
    On Error GoTo eh
    Dim oob As OLEObject
    Dim shpAutoShape As Shape
    
    If shp.Type = msoOLEControlObject Then Set oob = shp.OLEFormat.Object
    
    Select Case shp.Type
        Case mso3DModel:                TypeString = "3dModel"
        Case msoAutoShape:              TypeString = "AutoShape " & TypeStringAutoShape(shp)
        Case msoCallout:                TypeString = "CallOut"
        Case msoCanvas:                 TypeString = "Canvas"
        Case msoChart:                  TypeString = "Chart"
        Case msoComment:                TypeString = "Comment"
        Case msoContentApp:             TypeString = "ContentApp"
        Case msoDiagram:                TypeString = "Diagram"
        Case msoEmbeddedOLEObject:      TypeString = "EmbeddedOLEObject"
        Case msoFormControl:            TypeString = "FormControl " & TypeStringFormControl(shp)
        Case msoFreeform:               TypeString = "Freeform"
        Case msoGraphic:                TypeString = "Graphic"
        Case msoGroup:                  TypeString = "Group"
        Case msoInk:                    TypeString = "Ink"
        Case msoInkComment:             TypeString = "InkComment"
        Case msoLine:                   TypeString = "Line"
        Case msoLinked3DModel:          TypeString = "Linked3DModel"
        Case msoLinkedGraphic:          TypeString = "LinkedGraphic"
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
        Case msoWebVideo:               TypeString = "WebVideo"
        Case Else
            Debug.Print "Shape-Type: '" & shp.Type & "' Not implemented"
    End Select

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

