Attribute VB_Name = "mSyncTest"
Option Explicit

Private Const TEST_BOOK_SYNC        As String = "CompManSyncTest.xlsb"
Private Const TEST_SHEET_SOURCE     As String = "Test_A"
Private Const TEST_SHAPE_SOURCE     As String = "CommandButton2_Test_A"
Private Const TEST_OOB_SOURCE       As String = "CommandButtonActiveX_Test_A"
Private Const TEST_SHEET_TARGET     As String = "Test_A"
Private Const TEST_SHAPE_TARGET     As String = "CommandButton2_Test_A"
Private Const TEST_OOB_TARGET       As String = "CommandButtonActiveX_Test_A"

Private oobSource                   As OLEObject
Private oobTarget                   As OLEObject
Private shpSource                   As Shape
Private shpTarget                   As Shape
Private wshSource                   As Worksheet
Private wshTarget                   As Worksheet

Public wbkSource                    As Workbook
Public wbkTarget                    As Workbook

Private Property Get TestSyncTargetFullName() As String
    TestSyncTargetFullName = wsConfig.FolderSyncTarget & "\" & "CompManSyncTest\" & TEST_BOOK_SYNC
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mSyncTest." & sProc
End Function

Private Sub Test_00_RegressionTest()
' ------------------------------------------------------------------------------
' Test-Sync-Target-Workbook: <FolderSyncTarget>\CompManSyncTest\CompManSyncTest.xlsb
' Test-Sync_Source_Workbook: <FolderCompManRoot>\Common-VBA-Excel-Component-Management-Services\SyncTest\SyncSource\CompManSyncTest.xlsb
' ------------------------------------------------------------------------------
    Const PROC = "Test_00_RegressionTest"

    On Error Resume Next
    Dim wbk As Workbook
    
    Application.EnableEvents = False
    If mWbk.IsOpen(wsService.SyncSourceFullName, wbk) Then wbk.Close False
    If mWbk.IsOpen(wsService.CurrentServicedWorkbookFullName(), wbk) Then wbk.Close False
    If mWbk.IsOpen(TestSyncTargetFullName, wbk) Then wbk.Close False
    Application.EnableEvents = True
    
    Test_00_RegressionTest_AssertSyncTargetRegressionTestPreparation
    Test_00_RegressionTest_AssertSyncSourceRegressionTestPreparation
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    mWbk.GetOpen TestSyncTargetFullName ' regular initiation of the synchronization
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

    
Private Sub Test_00_RegressionTest_AssertSyncSourceRegressionTestPreparation()

End Sub

Private Sub Test_00_RegressionTest_AssertSyncTargetRegressionTestPreparation()

End Sub

Private Sub Test_99_mSyncShapeProperties_Name_Property()
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

Private Sub Test_99_mSyncShapes_ShapeName()
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
    
    Dim fso         As New FileSystemObject
    Dim wbk         As Workbook
    Dim sTargetWorkingCopy As String
    
    With fso
        sTargetWorkingCopy = mSync.TargetWorkingCopyFullName(TestSyncTargetFullName)
        If .FileExists(sTargetWorkingCopy) Then
            If mWbk.IsOpen(sTargetWorkingCopy, wbk) Then wbk.Close False
            .DeleteFile sTargetWorkingCopy
        End If
    End With
    
    Application.EnableEvents = False
    On Error Resume Next
    mSync.TargetWorkingCopy.Close False
    mSync.source.Close False
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
    mSync.TargetWorkingCopy = mWbk.GetOpen(TestSyncTargetFullName)
    Application.EnableEvents = True
'    mSync.Source = mSync.Target
    If wshSource Is Nothing Then Set wshSource = mSync.source.Worksheets(TEST_SHEET_SOURCE)
    If wshTarget Is Nothing Then Set wshTarget = mSync.TargetWorkingCopy.Worksheets(TEST_SHEET_TARGET)
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

Private Sub Regression():        TestSync True:          End Sub

Private Sub TestSelectedOnly():  TestSync False:         End Sub

Private Sub TestSync(Optional ByVal t_regression As Boolean = False)
    Const PROC = "TestSync"
    
    On Error GoTo eh
    
    wsService.CurrentServiceName = ErrSrc(PROC)
        
    mBasic.BoP ErrSrc(PROC)
    Application.ScreenUpdating = False
    With wsSyncTest
        .SetupTestWorkbooksEstablish t_regression
        .Setup "source", t_regression
        .wbkSource.Save
        If t_regression Then .wbkSource.Close ' will be re-opened by the service
        .Setup "target", t_regression
    End With
    
    If Not t_regression Then
        mService.Initiate mCompManClient.SRVC_SYNCHRONIZE, wsSyncTest.wbkTarget, False
        wsSyncTest.wbkSource.Save
        mSync.source = mWbk.GetOpen(wsSyncTest.SyncTestSourceFullName)
        mSync.TargetWorkingCopy = mWbk.GetOpen(wsSyncTest.SyncTestTargetFullName)
    End If
    
    If t_regression Then
        mSync.Initialize i_sync_refs:=True _
                       , i_sync_sheets:=True _
                       , i_sync_names:=True _
                       , i_sync_shapes:=False _
                       , i_sync_comps:=True
    Else
        mSync.Initialize i_sync_refs:=wsSyncTest.DoSync("Reference") _
                       , i_sync_sheets:=wsSyncTest.DoSync("Worksheet") _
                       , i_sync_names:=wsSyncTest.DoSync("Name") _
                       , i_sync_shapes:=wsSyncTest.DoSync("Shape") _
                       , i_sync_comps:=wsSyncTest.DoSync("VBComponent")
    End If

xt: mBasic.EoP ErrSrc(PROC)
    If t_regression Then
        '~~ Simmulates the Workbook_Open events's action:
        '~~ mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED
        mCompMan.SynchronizeVBProjects wsSyncTest.wbkTarget
    Else
        mSync.SyncMode
        mSync.RunSync
    End If
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Private Sub Test_TestEnvironent()
    Test_EnvironmentCleanup
    Test_EnvironmentProvide
    Test_EnvironmentCleanup
End Sub

