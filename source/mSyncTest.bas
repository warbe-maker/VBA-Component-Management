Attribute VB_Name = "mSyncTest"
Option Explicit

Private Const TEST_BOOK_SYNC        As String = "CompManSyncTest.xlsb"
Private Const TEST_SHEET_SOURCE     As String = "Test_A"
Private Const TEST_SHAPE_SOURCE     As String = "CommandButton2_Test_A"
Private Const TEST_OOB_SOURCE       As String = "CommandButtonActiveX_Test_A"
Private Const TEST_SHEET_TARGET     As String = "Test_A"
Private Const TEST_SHAPE_TARGET     As String = "CommandButton2_Test_A"
Private Const TEST_OOB_TARGET       As String = "CommandButtonActiveX_Test_A"

Private oobSource   As OLEObject
Private oobTarget   As OLEObject
Private shpSource   As Shape
Private shpTarget   As Shape
Private wbkServiced As Workbook
Private wbkSource   As Workbook
Private wbkTarget   As Workbook
Private wshSource   As Worksheet
Private wshTarget   As Worksheet

Private Property Get TestBookTarget() As String
    TestBookTarget = mConfig.ServicedSyncTargetFolder & "\" & TEST_BOOK_SYNC
End Property

Private Property Get TestBookTargetCopy() As String
    TestBookTargetCopy = mConfig.ServicedSyncTargetFolder & "\" & TEST_BOOK_SYNC & mSync.SYNC_COPY_SUFFIX
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mSyncTest." & sProc
End Function

Public Sub Test_00_RegressionTest()
    Const PROC = "Test_00_RegressionTest"

    On Error Resume Next
    Application.Workbooks(wsService.SyncSourceWorkbookName).Close False
    Application.Workbooks(wsService.SyncTargetWorkbookName).Close False
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    Application.EnableEvents = True
    mWbk.GetOpen TestBookTarget ' initiates the synchronization
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_99_mSyncShapeProperties_Name_Property()
' ----------------------------------------------------------------------------
' Changing the Name Property of the Shape or the OOB always changes both!
' Note: The possibility to change only one of the two and thereby making them
'       inconsistent could not be proved.
' ----------------------------------------------------------------------------
    Const PROC = "Test_99_mSyncShapeProperties_Name_Property"
        
    On Error GoTo eh
    Test_EnvironmentProvide
        
    '~~ Test 1: Change Shape.Name property (changing the Shape Name changes the OOB's CodeName accordingly!)
    shpSource.Name = "CommandButtonActiveX_Test_A_X"
    Debug.Assert mSyncShapes.ShapeNames(shpSource) = "CommandButtonActiveX_Test_A_X (CommandButtonActiveX_Test_A_X)"
    Debug.Assert mSyncShapes.ShapeNames(oobSource) = "CommandButtonActiveX_Test_A_X (CommandButtonActiveX_Test_A_X)"
    '~~ Undo
    shpSource.Name = "CommandButtonActiveX_Test_A"
    Debug.Assert mSyncShapes.ShapeNames(shpSource) = "CommandButtonActiveX_Test_A (CommandButtonActiveX_Test_A)"
    Debug.Assert mSyncShapes.ShapeNames(oobSource) = "CommandButtonActiveX_Test_A (CommandButtonActiveX_Test_A)"
    
    '~~ Test 2: Change the OLEObject.Name property  (changing the Code-Name changes the Shape's Name accordingly!)
    oobSource.Name = "cmbActiveX_Test_A"
    Debug.Assert mSyncShapes.ShapeNames(shpSource) = "cmbActiveX_Test_A (cmbActiveX_Test_A)"
    Debug.Assert mSyncShapes.ShapeNames(oobSource) = "cmbActiveX_Test_A (cmbActiveX_Test_A)"
    '~~ Undo
    oobSource.Name = "CommandButtonActiveX_Test_A"
    Debug.Assert mSyncShapes.ShapeNames(shpSource) = "CommandButtonActiveX_Test_A (CommandButtonActiveX_Test_A)"
    Debug.Assert mSyncShapes.ShapeNames(oobSource) = "CommandButtonActiveX_Test_A (CommandButtonActiveX_Test_A)"
        
xt: Test_EnvironmentCleanup
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_99_mSyncShapeProperties_OOBObjectProperties()
' ----------------------------------------------------------------------------
' Synchronize all applicable (red/write and different) Properties of oobTarget
' with oobSource.
' ----------------------------------------------------------------------------
    Const PROC = "Test_99_mSyncShapeProperties_OOBObjectProperties"
    
    Dim enProperty      As enProperties
    Dim lMaxLen         As Long
    Dim vTarget         As Variant
    Dim vSource         As Variant
    Dim sProperty       As String
    Dim dctRW           As New Dictionary
    Dim v               As Variant
    
    On Error GoTo eh
    Application.ScreenUpdating = False
    Test_EnvironmentProvide
    
    Test_ProvideControls wbkSource, "Test_A"
            
    mSyncShapePrprtys.ShapeSource = shpSource
    mSyncShapePrprtys.ShapeTarget = shpTarget
    mSyncShapePrprtys.SheetSource = wshSource
    mSyncShapePrprtys.SheetTarget = wshTarget
    mSyncShapePrprtys.OLEObjectSource = oobSource
    mSyncShapePrprtys.OLEObjectTarget = oobTarget
        
    lMaxLen = mSyncShapePrprtys.PropertyMaxLen(mSyncShapePrprtys.enPropertiesOOBFirst, mSyncShapePrprtys.enPropertiesOOBLast)
    
    '~~ Do synch for a specific OOB's properties
    Debug.Print "Synchronizing the properties for : " & wshTarget.Name & "." & ShapeNames(oobTarget)
    For enProperty = mSyncShapePrprtys.enPropertiesOOBFirst To mSyncShapePrprtys.enPropertiesOOBLast
        sProperty = mSyncShapePrprtys.PropertyName(enProperty)
        vTarget = mSyncShapePrprtys.PropertyValue(enProperty, oobTarget)
        vSource = mSyncShapePrprtys.PropertyValue(enProperty, oobSource)
        If vTarget <> vSource Then
            mSyncShapePrprtys.SyncProperty enProperty
            If mSyncShapePrprtys.PropertyValue(enProperty, oobTarget) = mSyncShapePrprtys.PropertyValue(enProperty, oobSource) Then
                If Not dctRW.Exists(sProperty) Then
                    '~~ synchronizability proved
                    mDct.DctAdd dctRW, sProperty, ": synchronizability proved (example " & wshSource.Name & "." & ShapeNames(oobSource) & ")", order_bykey, seq_ascending, sense_casesensitive
                End If
            Else
                Debug.Print mBasic.Align(sProperty, lMaxLen, , , ".") & " not changed " & mSyncShapePrprtys.PropertyChange(vTarget, vSource)
                Stop
            End If
        Else
            mSyncShapePrprtys.SynchabilityCheckOOB sProperty, dctRW
        End If
    Next enProperty
    
    For enProperty = mSyncShapePrprtys.enPropertiesOOBFirst To mSyncShapePrprtys.enPropertiesOOBLast
        sProperty = mSyncShapePrprtys.PropertyName(enProperty)
        If Not dctRW.Exists(sProperty) Then
            mDct.DctAdd dctRW, sProperty, ": synchronizability  n o t  proved!", order_bykey, seq_ascending, sense_casesensitive
        End If
    Next enProperty
    
    For Each v In dctRW
        Debug.Print mBasic.Align(v, lMaxLen, , , ".") & dctRW(v)
    Next v

xt: Test_EnvironmentCleanup
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_99_mSyncShapeProperties_ShapeProperties()
' ----------------------------------------------------------------------------
' Synchronize all Properties of all non OOB shapes in the source sheet which
' do have a corresponding shape in the target sheet.
' ----------------------------------------------------------------------------
    Const PROC = "Test_99_mSyncShapeProperties_ShapeProperties"
    
    Dim enProperty      As enProperties
    Dim lMaxLen         As Long
    Dim vTarget         As Variant
    Dim vSource         As Variant
    Dim sProperty       As String
    Dim dctRW           As New Dictionary
    Dim v               As Variant
    
    On Error GoTo eh
    Application.ScreenUpdating = False
    Test_EnvironmentProvide
    
    Test_ProvideControls wbkSource, "Test_A"
            
    mSyncShapePrprtys.ShapeSource = shpSource
    mSyncShapePrprtys.ShapeTarget = shpTarget
    mSyncShapePrprtys.SheetSource = wshSource
    mSyncShapePrprtys.SheetTarget = wshTarget
    mSyncShapePrprtys.OLEObjectSource = oobSource
    mSyncShapePrprtys.OLEObjectTarget = oobTarget
        
    lMaxLen = mSyncShapePrprtys.PropertyMaxLen(mSyncShapePrprtys.enPropertiesOOBFirst, mSyncShapePrprtys.enPropertiesOOBLast)
    
    For Each shpSource In wshSource.Shapes
        If shpSource.Type = msoOLEControlObject Then GoTo ns
        If mSyncShapes.CorrespTargetShape(shpSource, wshTarget, shpTarget) Is Nothing Then GoTo ns
        Debug.Print "Synchronizing the properties for : " & mSyncShapes.ItemSyncName(shpSource)
        
        mSyncShapePrprtys.ShapeSource = shpSource
        mSyncShapePrprtys.ShapeTarget = shpTarget
        For enProperty = mSyncShapePrprtys.enPropertiesShapeFirst To mSyncShapePrprtys.enPropertiesShapeLast
            sProperty = mSyncShapePrprtys.PropertyName(enProperty)
            On Error Resume Next
            vTarget = mSyncShapePrprtys.PropertyValue(enProperty, shpTarget)
            If Err.Number <> 0 Then
                '~~ The Property or Method is not supported or the Property has not a valid value
                GoTo np
            End If
            On Error GoTo eh
            vSource = mSyncShapePrprtys.PropertyValue(enProperty, shpSource)
            If vTarget <> vSource Then
                On Error Resume Next
                mSyncShapePrprtys.SyncProperty enProperty
                If Err.Number = 0 Then
                    If mSyncShapePrprtys.PropertyValue(enProperty, shpTarget) = mSyncShapePrprtys.PropertyValue(enProperty, shpSource) Then
                        If Not dctRW.Exists(sProperty) Then
                            '~~ synchronizability proved
                            mDct.DctAdd dctRW, sProperty, ": synchronizability proved (example " & mSyncShapes.ItemSyncName(shpSource) & ")", order_bykey, seq_ascending, sense_casesensitive
                        End If
                    Else
                        Debug.Print mBasic.Align(sProperty, lMaxLen, , , ".") & ": " & enProperty & " not changed " & mSyncShapePrprtys.PropertyChange(vTarget, vSource) & "  (without error!)"
                    End If
                Else
                    Debug.Print mBasic.Align(sProperty, lMaxLen, , , ".") & ": " & enProperty & " not changed " & mSyncShapePrprtys.PropertyChange(vTarget, vSource) & "  (Error: " & Err.Description & " " & Err.Number & ")"
                End If
            Else
                '~~ For a equal value check if it would have been the synchronizabel
                mSyncShapePrprtys.SynchabilityCheckShape enProperty, sProperty, dctRW
            End If
np:      Next enProperty
ns:  Next shpSource
    
    For enProperty = mSyncShapePrprtys.enPropertiesShapeFirst To mSyncShapePrprtys.enPropertiesShapeLast
        sProperty = mSyncShapePrprtys.PropertyName(enProperty)
        If Not dctRW.Exists(sProperty) Then
            mDct.DctAdd dctRW, sProperty, ": synchronizability n o t proved!", order_bykey, seq_ascending, sense_casesensitive
        End If
    Next enProperty
    
    For Each v In dctRW
        Debug.Print mBasic.Align(v, lMaxLen, , , ".") & dctRW(v)
    Next v

xt: Test_EnvironmentCleanup
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub



Public Sub Test_99_Read_Write_Properties_check()
    Const PROC = "Test_99_Read_Write_Properties_check"
    
    On Error GoTo eh
    Dim enProperty  As enProperties
    Dim lMaxLen     As Long
    Dim wsh         As Worksheet
    Dim oob         As OLEObject
    Dim dctRW       As New Dictionary
    Dim sProperty   As String
    Dim shp         As Shape
    Dim v           As Variant
    Dim sFile       As String
    
    Application.ScreenUpdating = False
    Test_EnvironmentProvide
    
    Test_ProvideControls wbkSource, "Test_A"
    
    '~~ Get the max property name lenght for Debug.Print and save all names to a Dictionary
    lMaxLen = mBasic.Max(mSyncShapePrprtys.PropertyMaxLen(mSyncShapePrprtys.enPropertiesOOBFirst, mSyncShapePrprtys.enPropertiesShapeLast))
    
    '~~ Do read/write test for all Controls in the Sync-Source-Workbook
    For Each wsh In wbkSource.Worksheets
        For Each oob In wsh.OLEObjects
            Set oobTarget = oob
            Set oobSource = oob
            For enProperty = mSyncShapePrprtys.enPropertiesOOBFirst To mSyncShapePrprtys.enPropertiesOOBLast
                sProperty = mSyncShapePrprtys.PropertyName(enProperty)
                On Error Resume Next
                mSyncShapePrprtys.SyncProperty enProperty
                If Err.Number = 0 Then
                    If Not dctRW.Exists(sProperty) Then
                        '~~ read/write proved
                        mDct.DctAdd dctRW, sProperty, ": synchronizability proved  (example " & wsh.Name & "." & ShapeNames(oob) & ")", order_bykey, seq_ascending, sense_casesensitive
                    End If
                End If
            Next enProperty
        Next oob
        
        For Each shp In wsh.Shapes
            Set shpTarget = shp
            Set shpSource = shp
            If shpTarget.Type = msoOLEControlObject Then GoTo ns
            For enProperty = mSyncShapePrprtys.enPropertiesShapeFirst To mSyncShapePrprtys.enPropertiesShapeLast
                sProperty = mSyncShapePrprtys.PropertyName(enProperty)
                On Error Resume Next
                mSyncShapePrprtys.SyncProperty enProperty
                If Err.Number = 0 Then
                    If Not dctRW.Exists(sProperty) Then
                        '~~ read/write proved
                        mDct.DctAdd dctRW, sProperty, ": synchronizability proved (example " & wsh.Name & "." & ShapeNames(shp) & ")", order_bykey, seq_ascending, sense_casesensitive
                    End If
                End If
            Next enProperty
ns:      Next shp
    Next wsh
    
    Debug.Print dctRW.Count
    sFile = mFile.Temp(, "txt")
    For Each v In dctRW
        mFile.txt(sFile) = mBasic.Align(v, lMaxLen, , , ".") & dctRW(v)
    Next v
    mMsg.ShellRun sFile, WIN_NORMAL
    Stop: mFile.Delete sFile
    
    '~~ Add all still missing (synchronizability not proved) Properties
    For enProperty = mSyncShapePrprtys.enPropertiesOOBFirst To mSyncShapePrprtys.enPropertiesShapeLast
        If enProperty <> mSyncShapePrprtys.enPropertiesOOBLast And _
           enProperty <> mSyncShapePrprtys.enPropertiesShapeFirst Then
            sProperty = mSyncShapePrprtys.PropertyName(enProperty)
            If Not dctRW.Exists(sProperty) Then
                mDct.DctAdd dctRW, sProperty, ": synchronizability not proved yet", order_bykey, seq_ascending, sense_casesensitive
            End If
        End If
    Next enProperty
    
    sFile = mFile.Temp(, "txt")
    For Each v In dctRW
        mFile.txt(sFile) = mBasic.Align(v, lMaxLen, , , ".") & dctRW(v)
    Next v
    mMsg.ShellRun sFile, WIN_NORMAL
    Stop: mFile.Delete sFile
        
xt: Set dctRW = Nothing
    Test_EnvironmentCleanup
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
Public Sub Test_99_mSyncShapes_ShapeName()
    Const PROC = "Test_99_mSyncShapes_ShapeName"
        
    On Error GoTo eh
    Test_EnvironmentProvide
        
    Debug.Assert mSyncShapes.ShapeNames(shpSource) = TEST_SHAPE_SOURCE & " (" & TEST_OOB_SOURCE & ")"
    Debug.Assert mSyncShapes.ShapeNames(oobSource) = TEST_SHAPE_SOURCE & " (" & TEST_OOB_SOURCE & ")"
    
xt: Test_EnvironmentCleanup
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_EnvironmentCleanup()
    Const PROC = "Test_EnvironmentCleanup"
    
    Dim fso As New FileSystemObject
    Dim wbk As Workbook
    
    With fso
        If .FileExists(TestBookTargetCopy) Then
            If mWbk.IsOpen(TestBookTargetCopy, wbk) Then wbk.Close False
            .DeleteFile TestBookTargetCopy
        End If
    End With
    If mWbk.IsOpen(TEST_BOOK_SYNC, wbk) Then wbk.Close False
    
    Application.EnableEvents = False
    On Error Resume Next
    wbkServiced.Close False
    
    On Error Resume Next
    wbkTarget.Close False
    wbkSource.Close False
    
    Set wshSource = Nothing
    Set wshTarget = Nothing
    Set shpSource = Nothing
    Set shpTarget = Nothing
    Set oobSource = Nothing
    Set oobTarget = Nothing
        
    Application.EnableEvents = True
    Set fso = Nothing
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_EnvironmentProvide()
    Const PROC = "Test_EnvironmentProvide"
    
    On Error Resume Next
    Test_EnvironmentCleanup
    
    On Error GoTo eh
    Application.EnableEvents = False
    Set wbkServiced = mWbk.GetOpen(TestBookTarget)
    Set wbkTarget = mSync.SyncTargetWorkingCopy(wbkServiced)
    Set wbkSource = mWbk.GetOpen(mSync.SyncTargetsSource(wbkServiced), False) ' for testing not opened read-only
    If wshSource Is Nothing Then Set wshSource = wbkSource.Worksheets(TEST_SHEET_SOURCE)
    If wshTarget Is Nothing Then Set wshTarget = wbkTarget.Worksheets(TEST_SHEET_TARGET)
    If shpSource Is Nothing Then Set shpSource = wshSource.Shapes(TEST_SHAPE_SOURCE)
    If shpTarget Is Nothing Then Set shpTarget = wshTarget.Shapes(TEST_SHAPE_TARGET)
    If oobSource Is Nothing Then Set oobSource = wshSource.OLEObjects(TEST_OOB_SOURCE)
    If oobTarget Is Nothing Then Set oobTarget = wshTarget.OLEObjects(TEST_OOB_TARGET)
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_TestEnvironent()
    Test_EnvironmentCleanup
    Test_EnvironmentProvide
    Test_EnvironmentCleanup
End Sub

Private Sub Test_ProvideControls(ByVal tpc_wbk As Workbook, _
                                 ByVal tpc_wsh As Variant)
' ----------------------------------------------------------------------------
' Adds any possibly still missing Test-Controls to the Worksheet (tpc_wsh) -
' an object or an object's name - in Workbook (tcp_wbk).
' ----------------------------------------------------------------------------
    Const PROC = "Test_ProvideControls"
    
    On Error GoTo eh
    Dim sName   As String
    Dim wsh     As Worksheet
    
    If TypeName(tpc_wsh) = "Worksheet" _
    Then Set wsh = tpc_wsh _
    Else Set wsh = tpc_wbk.Worksheets(tpc_wsh)
    
    sName = "Line1_Test_A"
    If Not mSyncShapes.Exists(sName, wsh) Then
        With wsh.Shapes.AddLine(10, 10, 250, 250).Line
            .Parent.Name = sName
            .DashStyle = msoLineDashDotDot
            .ForeColor.RGB = rgbRed
            .BeginArrowheadLength = msoArrowheadShort
            .BeginArrowheadStyle = msoArrowheadOval
            .BeginArrowheadWidth = msoArrowheadNarrow
            .EndArrowheadLength = msoArrowheadLong
            .EndArrowheadStyle = msoArrowheadTriangle
            .EndArrowheadWidth = msoArrowheadWide
            .InsetPen = False
        End With
    End If
        
    sName = "DropDown1_Test_A"
    If Not mSyncShapes.Exists(sName, wsh) Then
        With wsh.Shapes.AddFormControl(Type:=xlDropDown, Left:=10, Top:=10, Width:=100, Height:=10)
            .Name = sName
            With .ControlFormat
                 .DropDownLines = 10
                 .Enabled = True
                 .ListFillRange = wsh.Range("rngDropDownCells").Address
                 .MultiSelect = xlExtended
            End With
        End With
    End If
        
    sName = "ListBox1_Test_A"
    If Not mSyncShapes.Exists(sName, wsh) Then
        With wsh.Shapes.AddFormControl(Type:=xlListBox, Left:=100, Top:=10, Width:=100, Height:=100)
            .Name = sName
            With .ControlFormat
                .Enabled = False
                .ListFillRange = wsh.Range("rngDropDownCells").Address
            End With
        End With
    End If

    sName = "ScrollBar1_Test_A"
    If Not mSyncShapes.Exists(sName, wsh) Then
        With wsh.Shapes.AddFormControl(xlScrollBar, Left:=10, Top:=10, Width:=10, Height:=200)
            With .ControlFormat
                .LinkedCell = Replace(wsh.Range("cellLinked").Address, "$", vbNullString)
                .Max = 100
                .Min = 0
                .LargeChange = 10
                .SmallChange = 2
            End With
        End With
    End If
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ListShapes()
    Dim shp As Shape
    Dim wbk As Workbook
    Dim wsh As Worksheet
    
    Set wbk = mWbk.GetOpen("CompManSyncTest.xlsb")
    Set wsh = wbk.Worksheets("TesT_A")
    
    For Each shp In wsh.Shapes
        Debug.Print mSyncShapes.ShapeNames(shp)
        With shp
            If .Name = "Elbow Connector 3" Then
                .Visible = True
                .Top = 50
                .Left = 50
            End If
        End With
    Next shp
    
End Sub

