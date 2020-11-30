Attribute VB_Name = "mTest"
Option Explicit
Option Private Module

Private Const TEST_CHANGE = "' Test code change"
Private cTest   As clsTestService
Private cComp   As New clsComp
Private wbTest   As Workbook
Private wbSrc   As Workbook
Private wbTrgt  As Workbook
Private vbc     As VBComponent
Private vbcm    As CodeModule

Public Sub Test_Temp()

    Set wbTest = ThisWorkbook
    Set vbc = wbTest.VBProject.VBComponents("mFile")
    Set vbcm = vbc.CodeModule
    
    Debug.Print vbc.Properties(1)
    Debug.Print vbcm.Parent.Properties(1)
    
    Cleanup
    
End Sub
    
Public Sub Test_CompOriginHasChanged()

End Sub

Public Sub Test_02_ExportChangedComponents()
    Const PROC = "Test_ExportChangedComponents"
    
    mErH.BoP ErrSrc(PROC)
    mCompMan.ExportChangedComponents ThisWorkbook
    mErH.EoP ErrSrc(PROC)
    
End Sub


Public Sub Regression()
' -----------------------------------------------------------------
'
' -----------------------------------------------------------------
    Const PROC = "Regression"
    
    Set cTest = New clsTestService
    cTest.Regression = True
    
    mErH.BoP ErrSrc(PROC)
    Test_01_01_CodeChanged
    Test_01_02_CodeChanged
    Test_01_03_CodeChanged
    Test_01_04_CodeChanged
    mErH.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_01_01_CodeChanged()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_01_01_CodeChanged"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim wbTest      As Workbook
    Dim sExpFile    As String
    Dim vResult     As Variant
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Begin of test preparation
    With cTest
        .TestProcedure = ThisWorkbook.Name & ": " & ErrSrc(PROC)
        .TestItem = ThisWorkbook.Name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        Set wbTest = mWrkbk.GetOpen(ThisWorkbook.Path & "\" & "Test\Test1.xlsm")
        .ResultExpected = True
        .Details = "Export File does not exist, CodeModule not/never exported"
    End With
    
    With cComp
        .Host = wbTest
        .VBComp = .Host.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If fso.FileExists(sExpFile) Then fso.DeleteFile (sExpFile)
    End With
    
    ' Test: Code is regarded changed because there is no export file
    mErH.BoP ErrSrc(PROC)
    vResult = cComp.CodeChanged
    mErH.EoP ErrSrc(PROC)
    
    ' Evaluating the result
    If cTest.Evaluated(vResult) = cTest.FAILED Then Stop
    
xt: Cleanup exp_file:=sExpFile, vbc:=cComp.VBComp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_01_02_CodeChanged()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_01_02_CodeChanged"
    
    On Error GoTo eh
    Dim sExpFile    As String
    Dim vResult     As Variant
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Test preparation
    With cTest
        .TestProcedure = ThisWorkbook.Name & ": " & ErrSrc(PROC)
        .TestItem = ThisWorkbook.Name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        Set wbTest = mWrkbk.GetOpen(ThisWorkbook.Path & "\" & "Test\Test1.xlsm")
        .ResultExpected = False
        .Details = "Export File is identical with CodeModule"
    End With
    
    With cComp
        .Host = wbTest
        .VBComp = .Host.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If .VBComp.CodeModule.Lines(1, 1) = TEST_CHANGE Then .VBComp.CodeModule.DeleteLines 1, 1
        .VBComp.Export sExpFile
    End With
    
    ' Test: Code has not changed because it is identical with the export file
    mErH.BoP ErrSrc(PROC)
    vResult = cComp.CodeChanged
    mErH.EoP ErrSrc(PROC)
    
    ' Evaluating the result
    If cTest.Evaluated(vResult) = cTest.FAILED Then Stop
           
xt: Cleanup exp_file:=sExpFile, vbc:=cComp.VBComp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_01_03_CodeChanged()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_01_03_CodeChanged"
    
    On Error GoTo eh
    Dim sExpFile    As String
    Dim vResult     As Variant
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Test preparation
    With cTest
        .TestProcedure = ThisWorkbook.Name & "." & ErrSrc(PROC)
        .TestItem = ThisWorkbook.Name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        .Details = "Export File outdated, CodeModule changed (additional line)"
        .ResultExpected = True
    End With
    Set wbTest = mWrkbk.GetOpen(ThisWorkbook.Path & "\" & "Test\Test1.xlsm")
    
    With cComp
        .Host = wbTest
        .VBComp = .Host.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If .VBComp.CodeModule.Lines(1, 1) = TEST_CHANGE Then .VBComp.CodeModule.DeleteLines 1, 1
        .VBComp.Export sExpFile ' Overwrites any existing
        .VBComp.CodeModule.InsertLines 1, TEST_CHANGE
    End With
    
    '~~ Test: Code is regarded changed because it is not identical with the Export File
    mErH.BoP ErrSrc(PROC)
    vResult = cComp.CodeChanged
    mErH.EoP ErrSrc(PROC)
    
    ' Evaluating the result
    If cTest.Evaluated(vResult) = cTest.FAILED Then Stop
           
xt: Cleanup exp_file:=sExpFile, vbc:=cComp.VBComp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub


Public Sub Test_01_04_CodeChanged()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_01_04_CodeChanged"
    
    On Error GoTo eh
    Dim sExpFile    As String
    Dim vResult     As Variant
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Test preparation
    With cTest
        .TestProcedure = ThisWorkbook.Name & "." & ErrSrc(PROC)
        .TestItem = ThisWorkbook.Name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        .Details = "Additional empty line in CodeModule is not considered a change though the Export File is outdated"
        .ResultExpected = False
    End With
    Set wbTest = mWrkbk.GetOpen(ThisWorkbook.Path & "\" & "Test\Test1.xlsm")
    
    With cComp
        .Host = wbTest
        .VBComp = .Host.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If .VBComp.CodeModule.Lines(1, 1) = TEST_CHANGE Then .VBComp.CodeModule.DeleteLines 1, 1
        .VBComp.Export sExpFile ' Overwrites any existing
        .VBComp.CodeModule.InsertLines 1, vbLf
    End With
    
    '~~ Test: Code is regarded changed because it is not identical with the Export File
    mErH.BoP ErrSrc(PROC)
    vResult = cComp.CodeChanged(ignore_empty_lines:=True)
    mErH.EoP ErrSrc(PROC)
    
    ' Evaluating the result
    If cTest.Evaluated(vResult) = cTest.FAILED Then Stop
           
xt: Cleanup exp_file:=sExpFile, vbc:=cComp.VBComp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_UpdateUsedCommCompsTheOriginHasChanged()
    Const PROC  As String = "Test_UpdateCommonModules"
    
    Dim wb      As Workbook

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    
    Application.StatusBar = vbNullString
    Set wb = ThisWorkbook
    mCompMan.UpdateUsedCommCompsTheOriginHasChanged wb
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_Refs()
    
    Dim ref As Reference

    For Each ref In ThisWorkbook.VBProject.References
        With ref
            If InStr(.Description, "Applications Extensibility") <> 0 Then
                Debug.Print .Description
                Debug.Print .GUID
                Debug.Print "Major=" & .Major
                Debug.Print "Minor=" & .Minor
                Debug.Print TypeName(.GUID)
            End If
            If InStr(.Description, "Scripting Runtime") <> 0 Then
                Debug.Print .Description
                Debug.Print .GUID
                Debug.Print "Major=" & .Major
                Debug.Print "Minor=" & .Minor
                Debug.Print TypeName(.GUID)
            End If
        End With
    Next ref
    
End Sub

Public Sub Test_DisplayDiff()
    
    Dim cComp As New clsComp

    With cComp
        .Host = ThisWorkbook
        .VBComp = .Host.VBProject.VBComponents("mBasic")
        .DisplayDiff .ExportFile, .ExportFile, "Test title left", "Test title right"
    End With
    
End Sub

Private Sub Cleanup(Optional ByVal exp_file As String = vbNullString, _
                    Optional ByRef vbc As VBComponent = Nothing)
        
    If exp_file <> vbNullString Then
        With New FileSystemObject
            If .FileExists(exp_file) Then .DeleteFile exp_file
        End With
    End If
    
    If Not vbc Is Nothing Then
        With vbc.CodeModule
            If .Lines(1, 1) = TEST_CHANGE Then .DeleteLines 1, 1
            While Len(.Lines(1, 1)) = 0
                .DeleteLines 1.1
            Wend
        End With
    End If
    
    If Not cComp Is Nothing Then Set cComp = Nothing
    If Not vbcm Is Nothing Then Set vbcm = Nothing
    If Not vbc Is Nothing Then Set vbc = Nothing
    On Error Resume Next: wbTest.Close SaveChanges:=False
    On Error Resume Next: wbSrc.Close SaveChanges:=False
    On Error Resume Next: wbTrgt.Close SaveChanges:=False

End Sub

Public Sub Test_CleanExportFile()
    
    Dim cComp As New clsComp

    With cComp
        .Host = ThisWorkbook
        .VBComp = ThisWorkbook.VBProject.VBComponents("mFile")
        .CleanExportFile .ExportFile
    End With
    
End Sub

Public Sub Test_CodeModuleTrimm()

    Dim vbc As VBComponent
    Dim wb  As Workbook
    
    Set wb = ActiveWorkbook
    Set vbc = wb.VBProject.VBComponents("mCommon")
    Debug.Print "Trim CodeModule 'mCommon' in Workbook '" & wb.Name & "'"
    mVBP.CodeModuleTrim vbc, wb
    
    Set wb = ThisWorkbook
    Set vbc = wb.VBProject.VBComponents("mCommon")
    Debug.Print "Trim CodeModule 'mCommon' in Workbook '" & wb.Name & "'"
    mVBP.CodeModuleTrim vbc, wb
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function
