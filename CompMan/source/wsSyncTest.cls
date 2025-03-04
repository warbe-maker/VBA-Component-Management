VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsSyncTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public wbkTarget                            As Workbook
Public wbkSource                            As Workbook
Private Const SYNC_TEST_WBK_REGRESSION      As String = "SyncTestRegression.xlsb"
Private Const SYNC_TEST_WBK_SOURCE          As String = "SyncTestSource.xlsb"
Private Const SYNC_TEST_WBK_TARGET          As String = "SyncTestTarget.xlsb"
Private Const SYNC_TEST_REGRESSION_FOLDER   As String = "SyncTestRegression"
Private rngRow                              As Range
Private wbkSyncTestSource                   As Workbook
Private wbkSyncTestTarget                   As Workbook
Private sSyncTestSourceFullName             As String
Private sSyncTestTargetFullName             As String

Private Property Get ItemAddress() As String
    ItemAddress = Intersect(rngRow.EntireRow, Me.Range("colItemAddress").EntireColumn).Value
End Property

Private Property Get ItemCodeName() As String
    ItemCodeName = Intersect(rngRow.EntireRow, Me.Range("colItemCodeName").EntireColumn).Value
End Property

Private Property Get ItemComment() As String
    ItemComment = Intersect(rngRow.EntireRow, Me.Range("colItemComment").EntireColumn).Value
End Property

Private Property Get ItemCompType() As String
    ItemCompType = Intersect(rngRow.EntireRow, Me.Range("colItemCompType").EntireColumn).Value
End Property

Private Property Get ItemKind() As String
    ItemKind = Intersect(rngRow.EntireRow, Me.Range("colItemKind").EntireColumn).Value
End Property

Private Property Get ItemName() As String
    ItemName = Intersect(rngRow.EntireRow, Me.Range("colItemName").EntireColumn).Value
End Property

Private Property Get ItemRange(ByVal i_wbk As Workbook) As Range
    Const PROC = "ItemRange"
    
    On Error GoTo eh
    Dim rng As Range
    Dim wsh As Worksheet
    
    Set wsh = i_wbk.Worksheets(SheetName)
    Set rng = wsh.Range(ItemAddress)
    If rng.row = 1 _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The setup range for a Test-Name must not be row 1 which is reserved as header!"
    Set ItemRange = rng

xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg. | Installed component |
'         |------------------|---------------------|
'         | mErH = 1         | mErH                |
'         | mTrc = 1         | mTrc                |
'         | clsTrc = 1       | clsTrc              |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ The error handling will also hand over to the Common VBA Execution Trace
    '~~ provided one is installed  (mTrc/clsTrc) and activated

    mErH.EoP e_proc
#ElseIf clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf mTrc = 1 Then
    mTrc.EoP e_proc, e_inf
#End If

End Sub

Private Property Get ItemScope(ByVal i_wbk As Workbook) As Variant
    Dim s   As String
    s = Intersect(rngRow.EntireRow, Me.Range("colItemScopeName").EntireColumn).Value
    If s = vbNullString _
    Then Set ItemScope = i_wbk _
    Else Set ItemScope = i_wbk.Worksheets(s)
End Property

Private Property Get SheetName() As String
    SheetName = Intersect(rngRow.EntireRow, Me.Range("colSheetName").EntireColumn).Value
End Property

Private Property Get SourceOrTarget() As String
    SourceOrTarget = Intersect(rngRow.EntireRow, Me.Range("colItemSourceOrTarget").EntireColumn).Value
End Property

Private Property Get SyncTestSource() As Workbook:                               Set SyncTestSource = wbkSyncTestSource:                                     End Property

Private Property Let SyncTestSource(ByVal wbk As Workbook):                      Set wbkSyncTestSource = wbk:                                                End Property

Public Property Get SyncTestSourceFullName() As String:                          SyncTestSourceFullName = sSyncTestSourceFullName:                           End Property

Private Property Let SyncTestSourceFullName(ByVal s As String):                  sSyncTestSourceFullName = s:                                                End Property

Private Property Get SyncTestTarget() As Workbook:                               Set SyncTestTarget = wbkSyncTestTarget:                                     End Property

Private Property Let SyncTestTarget(ByVal wbk As Workbook):                      Set wbkSyncTestTarget = wbk:                                                End Property

Public Property Get SyncTestTargetFullName() As String:                          SyncTestTargetFullName = sSyncTestTargetFullName:                           End Property

Private Property Let SyncTestTargetFullName(ByVal s As String):                  sSyncTestTargetFullName = s:                                                End Property

Private Property Get TestSetupData() As Range
    Set TestSetupData = Intersect(Me.Range("colsTestData").EntireColumn, Me.Range("rwsTestData"), Me.UsedRange.EntireRow)
End Property

Public Function DoSync(ByVal d_kind As String) As Boolean
    Dim cel As Range
    
    For Each cel In Intersect(Me.Range("rngDoSyncs").EntireRow, Me.Range("colDoSyncKind").EntireColumn).Cells
        If cel.Value = d_kind Then
            DoSync = cel.Offset(0, -1).Value = "x"
            Exit For
        End If
    Next cel
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Me.CodeName & "." & sProc
End Function

Private Sub FlipState(ByVal cel As Range)
    If cel.Cells.CountLarge = 1 Then
        If Not IsEmpty(cel.Offset(0, 1).Value) Then
            mRng.FlipStages cel, True, cel.Offset(0, 1), "x", vbNullString
        End If
    End If
End Sub

Private Function ItemFromFile() As String
    ItemFromFile = Intersect(rngRow.EntireRow, Me.Range("colItemFromFile").EntireColumn).Value
End Function

Private Function ItemIsActive() As Boolean
    ItemIsActive = Intersect(rngRow.EntireRow, Me.Range("colItemActive").EntireColumn).Value = "x"
End Function

Public Sub Setup(ByVal s_source_or_target As String, _
        Optional ByVal s_regression As Boolean = False)
' ----------------------------------------------------------------------------
' Attention: For a non-regression test the Sync-Test-Source and the Sync-Test-
'            Target-Workbook both reside in a dedicated sub-folder of
'            ThisWorbook.
'            For a regression test the Sync-Test-Target-Workbook resides in a
'            dedicated folder in the configured Sy-Target-Folder!
' ----------------------------------------------------------------------------
    Const PROC = "Setup"
    
    On Error GoTo eh
    Dim bFirstSource    As Boolean: bFirstSource = True
    Dim bFirstTarget    As Boolean: bFirstTarget = True
    Dim wbk             As Workbook
    
    Set wbk = Nothing
    
    '~~ 1. setup all source
    For Each rngRow In TestSetupData.Rows
        If ItemKind = vbNullString Then Exit For
        If SourceOrTarget = s_source_or_target Then
            If wbk Is Nothing Then Set wbk = Wrkbk(s_source_or_target)
            If ItemIsActive() Or s_regression Then
                Select Case ItemKind
                    Case "Name":        SetupName n_name:=ItemName _
                                                , n_rng:=ItemRange(wbk) _
                                                , n_scope:=ItemScope(wbk) _
                                                , n_comment:=ItemComment
                    Case "Reference":
                                        SetupReference r_name:=ItemName _
                                                     , r_wbk:=wbk
                    Case "Shape"        ' not finally implemented
                    Case "VBComponent": SetupComp c_source_or_target:=s_source_or_target _
                                                , c_name:=ItemName _
                                                , c_type:=ItemCompType _
                                                , c_from_file:=ItemFromFile _
                                                , c_wbk:=wbk
                                                
                    Case "Worksheet":   SetupSheet s_name:=ItemName _
                                                 , s_code_name:=ItemCodeName _
                                                 , s_wbk:=wbk _
                                                 , s_comment:=ItemComment _
                                                 , s_source_or_target:=SourceOrTarget _
                                                 , s_first_source:=bFirstSource _
                                                 , s_first_target:=bFirstTarget
                End Select
            End If
        End If
    Next rngRow
    
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SetupComp(ByVal c_source_or_target As String, _
                      ByVal c_name As String, _
                      ByVal c_type As String, _
                      ByVal c_from_file As String, _
                      ByVal c_wbk As Workbook)
' ----------------------------------------------------------------------------
' Creates a VBComponent named (c_name) of type (c_type) in the Workbook
' (c_wbk) with the content (c_from_file).
' ----------------------------------------------------------------------------
    Const PROC = "SetupComp"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim sFromFile   As String
    
    '~~ The file serving as the source-code for the VBComponent to be set up is the
    '~~ dedicated test sync workbook source or target folder
    sFromFile = SetupCompFromFile(c_source_or_target, c_from_file)
    
    Select Case FSo.GetExtensionName(sFromFile)
        Case "cls", "frm", "bas"
            c_wbk.VBProject.VBComponents.Import sFromFile
        Case Else ' content from data/file
            With c_wbk.VBProject.VBComponents
                Select Case c_type
                    Case "Document Module (Worksheet)" _
                       , "Document Module (Workbook)"
                                                        SetupCompDocMod s_wbk:=c_wbk, s_comp_name:=c_name, s_comp_type:=c_type, s_vbc:=vbc
                    Case "Standard Module":             Set vbc = .Add(vbext_ct_StdModule)
                    Case "Class Module":                Set vbc = .Add(vbext_ct_ClassModule)
                    Case "UserForm":                    Set vbc = .Add(vbext_ct_MSForm)
                End Select
            End With
            If Not vbc Is Nothing Then
                vbc.CodeModule.AddFromFile sFromFile
                If c_type <> "Document Module (Workbook)" Then
                    '~~ The Workbook's module name cannot be changed other than with VBA's IDE.
                    '~~ But even when this is done the changed CodeName cannot be synchronized.
                    vbc.Name = c_name
                End If
            End If
    End Select
    
xt: Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SetupCompDocMod(ByVal s_wbk As Workbook, _
                            ByVal s_comp_name As String, _
                            ByVal s_comp_type As String, _
                            ByRef s_vbc As VBComponent)
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    
    On Error Resume Next ' in case the Worksheet has yet not been setup
    Select Case s_comp_type
        Case "Document Module (Worksheet)": Set s_vbc = s_wbk.VBProject.VBComponents(s_comp_name)
        Case "Document Module (Workbook)":  Set s_vbc = s_wbk.VBProject.VBComponents(s_wbk.CodeName)
    End Select

End Sub

Private Function SetupCompFromFile(ByVal s_source_or_target As String, _
                                   ByVal s_from_file As String) As String
' ----------------------------------------------------------------------------
' Returns the file which is the source for the setup of the VBComponent.
' ----------------------------------------------------------------------------
    Const PROC = "Setup"
    
    On Error GoTo eh
    Dim sExt        As String
    Dim sFromFile1  As String
    Dim sFromFile2  As String
    Dim sExtension  As String
    Dim sFromFile   As String
    Dim sPath1      As String
    Dim sPath2      As String
    
    sPath1 = ThisWorkbook.Path & "\source\"
    sExt = FSo.GetExtensionName(s_from_file)
    Select Case s_source_or_target
        Case "Sync_SourceWorkbook"
            sPath2 = ThisWorkbook.Path & "\SyncTestSource\setup\"
            If sExtension = "dat" Then
                sFromFile1 = sPath1 & s_from_file
                sFromFile2 = sPath2 & s_from_file
            Else
                sFromFile1 = sPath1 & s_from_file
                sFromFile2 = sPath2 & s_from_file
            End If
        
        Case "Sync_TargetWorkbook"
            sPath2 = ThisWorkbook.Path & "\SyncTestTarget\setup\"
            If sExtension = "dat" Then
                sFromFile1 = sPath1 & s_from_file
                sFromFile2 = sPath2 & s_from_file
            Else
                sFromFile1 = sPath1 & s_from_file
                sFromFile2 = sPath2 & s_from_file
            End If
    End Select
    If FSo.FileExists(sFromFile2) Then ' preferred when both exist
        sFromFile = sFromFile2
    ElseIf FSo.FileExists(sFromFile1) Then
        sFromFile = sFromFile1
    Else
        Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The provided file (" & s_from_file & ") neither exist in " & vbLf & vbLf & _
                                           "'" & sPath1 & "' nor in " & vbLf & vbLf & _
                                           "'" & sPath2 & "'!"
    End If
    
xt: SetupCompFromFile = sFromFile
    Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function SetupName(ByVal n_name As String, _
                           ByVal n_rng As Range, _
                           ByVal n_scope As Variant, _
                  Optional ByVal n_comment As String = vbNullString, _
                  Optional ByRef n_nme As Name) As Name
' ------------------------------------------------------------------------
' Returns a Name object, optionally (n_nme), with the provided properties
' and the provided scope (n_scope).
' ------------------------------------------------------------------------
    Const PROC = "SetupName"
    
    On Error GoTo eh
    Dim nme As Name
         
    Set nme = n_scope.Names.Add(n_name, n_rng)
    '~~ Note: The Name is unique! When the same name is added (n_name)
    '~~       the previous added Name's properties will be overwritten.
    nme.Comment = n_comment
    nme.RefersTo = n_rng
    
    With n_rng
        .Value = n_name ' exclusively for this test preparation
        .EntireColumn.AutoFit
    End With
    With n_rng.Offset(0, 1)
        .Value = n_comment
        .EntireColumn.AutoFit
    End With
    With n_rng.Offset(0, 2)
        If n_rng.Offset(0, 1).Value = "obsolete" Then
            .Formula = Replace("=IF(ISERROR(""Name '""&" & "<n>" & "&""' used for the '""&RC[-1]&""' test""),""Name '""&RC[-2]& ""' has been removed!"")", "<n>", n_name)
        Else
            .Formula = Replace("=""Name '""&" & "<n>" & "&""' used for the '""&RC[-1]&""' test""", "<n>", n_name)
        End If
        .EntireColumn.AutoFit
    End With
    
    Set n_nme = nme
    Set SetupName = nme
    Set nme = Nothing

xt: Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub SetupReference(ByVal r_name As String, _
                           ByVal r_wbk As Workbook)
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    Const PROC = "SetupReference"
    
    On Error GoTo eh
    Dim ref As Reference
    
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Description = r_name Then
            On Error Resume Next
            r_wbk.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
            If Err.Number <> 0 _
            Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Reference '" & ref.GUID & "' not added"
        End If
    Next ref

xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function SetupSheet(ByVal s_name As String, _
                            ByVal s_code_name As String, _
                            ByVal s_wbk As Workbook, _
                            ByVal s_source_or_target As String, _
                            ByRef s_first_source As Boolean, _
                            ByRef s_first_target As Boolean, _
                   Optional ByVal s_comment As String = vbNullString) As Worksheet
' ------------------------------------------------------------------------
' Returns a Name object, optionally (s_nme), with the provided properties
' and the provided scope (s_scope).
' ------------------------------------------------------------------------
    Const PROC = "SetupSheet"
    
    On Error GoTo eh
    Dim i   As Long
    Dim wsh As Worksheet
         
    Set wsh = s_wbk.Worksheets.Add
    wsh.Name = s_name
    mWsh.ChangeCodeName s_wbk, wsh.CodeName, s_code_name
    If s_comment Like "*owned-by-project*" Then
        '~~ sheet is protected an has no unlocked cell
        wsh.Protect
    End If
    Set SetupSheet = wsh
    Set wsh = Nothing
    
    Application.DisplayAlerts = False
    If s_first_target And s_source_or_target = "Sync_TargetWorkbook" Then
        For i = 2 To s_wbk.Worksheets.Count
            Set wsh = s_wbk.Worksheets(2)
            wsh.Delete
        Next i
        s_first_target = False
    End If
    If s_first_source And s_source_or_target = "Sync_SourceWorkbook" Then
        For i = 2 To s_wbk.Worksheets.Count
            Set wsh = s_wbk.Worksheets(2)
            wsh.Delete
        Next i
        s_first_source = False
    End If
    Application.DisplayAlerts = True
    
xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub SetupTestWorkbooksCleanup(ByVal s_sync_source As String, _
                                        ByVal s_sync_target As String, _
                               Optional ByVal s_sync_target_copy As String = vbNullString)
' ------------------------------------------------------------------------
'
' ------------------------------------------------------------------------
    Const PROC = "SetupTestWorkbooksCleanup"

    On Error GoTo eh
    Dim wbk As Workbook
    
    With FSo
        If .FileExists(s_sync_source) Then
            If mWbk.IsOpen(s_sync_source, wbk) Then wbk.Close False
            .DeleteFile s_sync_source
        
        End If
        If .FileExists(s_sync_target) Then
            If mWbk.IsOpen(s_sync_target, wbk) Then wbk.Close False
            .DeleteFile s_sync_target
        End If
        
        If Not s_sync_target_copy = vbNullString Then
            If .FileExists(s_sync_target_copy) Then
                If mWbk.IsOpen(s_sync_target_copy, wbk) Then wbk.Close False
                .DeleteFile s_sync_target_copy
            End If
        End If
    End With
    
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SetupTestWorkbooksEstablish(Optional ByVal s_regression As Boolean = False)
' ------------------------------------------------------------------------
' Establishes a Sync-Source-Workbook in a dedicated sub-folders of this
' this Workbook and a Sync-Target-Workbook:
' - for regression-test (s_regression = True) in a dedicated folder in the
'   configured Sync-Target-Folder
' - for unit-test (s_regression = False) in a dedicated sub-folders of
'   this Workbook.
' ------------------------------------------------------------------------
    Const PROC = "SetupTestWorkbooksEstablish"
    
    On Error GoTo eh
    Dim sSourcePath         As String
    Dim sTargetPath         As String
    Dim sSourceFullName     As String
    Dim sTargetFullName     As String
    Dim sTrgtCpyFullName    As String
    Dim sExt                As String
    
    Application.ScreenUpdating = False
    With FSo
        If s_regression Then
            '~~ For the regression test the source Workbook is established
            '~~ in the configured "Serviced Development and Test folder"
            sSourcePath = ThisWorkbook.Path & "\" & SYNC_TEST_REGRESSION_FOLDER
            sSourceFullName = sSourcePath & "\" & SYNC_TEST_WBK_REGRESSION
            '~~ For the regression test the target Workbook is established
            '~~ in the configured "Serviced Sync-Target-Folder"
            sTargetPath = wsConfig.FolderSyncTarget & "\" & SYNC_TEST_REGRESSION_FOLDER
            sTargetFullName = sTargetPath & "\" & SYNC_TEST_WBK_REGRESSION
            sExt = "." & .GetExtensionName(sTargetFullName)
            sTrgtCpyFullName = Replace(sTargetFullName, sExt, SYNC_TARGET_SUFFIX & sExt)
        Else
            '~~ For the selected tests the Workbooks are established within the CompMan folder
            sSourcePath = ThisWorkbook.Path & "\SyncTestSource"
            sTargetPath = ThisWorkbook.Path & "\SyncTestTarget"
            sSourceFullName = sSourcePath & "\" & SYNC_TEST_WBK_SOURCE
            sTargetFullName = sTargetPath & "\" & SYNC_TEST_WBK_TARGET
        End If
        If Not .FolderExists(sSourcePath) Then .CreateFolder sSourcePath
        If Not .FolderExists(sTargetPath) Then .CreateFolder sTargetPath
    End With
    
    SetupTestWorkbooksCleanup sSourceFullName, sTargetFullName, sTrgtCpyFullName
    
    '~~ Create and save source Workbook
    SyncTestSource = Application.Workbooks.Add
    SyncTestSource.SaveAs FileName:=sSourceFullName _
                         , FileFormat:=xlExcel12
    SyncTestSourceFullName = SyncTestSource.FullName
    If s_regression Then SyncTestSource.Close True
    
    '~~ Create and save target Workbook
    SyncTestTarget = Application.Workbooks.Add
    SyncTestTarget.SaveAs FileName:=sTargetFullName _
                         , FileFormat:=xlExcel12
    SyncTestTargetFullName = SyncTestTarget.FullName
    If s_regression Then SyncTestTarget.Close True
    
    '~~ Replaces the Services.Initiate which is inappropriate for individual and regression tests
    If Not s_regression Then
        wsService.SyncSourceFullName = SyncTestSource.FullName
        wsService.SyncTargetFullNameCopy = SyncTestTarget.FullName
        wsService.CurrentServicedWorkbookFullName = SyncTestTarget.FullName
    End If
        
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Worksheet_Activate()
    Const PROC = "Worksheet_Activate"
    
    Dim rng As Range
    
    mCompManClient.Events ErrSrc(PROC), False
    Set rng = Intersect(Me.Range("rwsTestData").EntireRow, Me.UsedRange, Me.Range("colsTestData").EntireColumn)
    On Error Resume Next
    rng.EntireRow.RowHeight = 11
    rng.Select
    mCompManClient.Events ErrSrc(PROC), True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Const PROC = "Worksheet_SelectionChange"
    
    Dim celCurrent As Range
    
    mCompManClient.Events ErrSrc(PROC), False
    If Target.Cells.CountLarge > 1 Then GoTo xt
    
    Application.ScreenUpdating = False
    Select Case True
        Case Not Intersect(Target, Me.Range("colItemActive").EntireColumn) Is Nothing
            Set celCurrent = Target
            FlipState Target
            celCurrent.Select
        Case Not Intersect(Target, Me.Range("rngDoSyncs").EntireRow, Me.Range("colDoSyncs").EntireColumn) Is Nothing
            Set celCurrent = Target
            FlipState Target
            celCurrent.Select
    End Select

xt: mCompManClient.Events ErrSrc(PROC), True

End Sub

Public Function Wrkbk(ByVal w_source_or_target As String) As Workbook
    Select Case w_source_or_target
        Case "Sync_SourceWorkbook"
            Set wbkSource = mWbk.GetOpen(wsSyncTest.SyncTestSourceFullName)
            Set Wrkbk = wbkSource
        Case "Sync_TargetWorkbook"
            Set wbkTarget = mWbk.GetOpen(wsSyncTest.SyncTestTargetFullName)
            Set Wrkbk = wbkTarget
    End Select
End Function

