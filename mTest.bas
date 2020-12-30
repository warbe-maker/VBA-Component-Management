Attribute VB_Name = "mTest"
Option Explicit
Option Private Module

Private Const TEST_CHANGE = "' Test code change"

Private cTest   As clsTestService
Private cComp   As New clsComp
Private wbTest  As Workbook
Private wbSrc   As Workbook
Private wbTrgt  As Workbook
Private vbc     As VBComponent
Private vbcm    As CodeModule

Public Sub Cleanup(Optional ByVal exp_file As String = vbNullString, _
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function

Public Sub Regression()
' -----------------------------------------------------------------
'
' -----------------------------------------------------------------
    Const PROC = "Regression"
    
    Set cTest = New clsTestService
    cTest.Regression = True
    
    mErH.BoP ErrSrc(PROC)
    Test_01_KindOfComp
    Test_05_01_KindOfCodeChange_NoRaw_UsedOnly
    Test_05_02_KindOfCodeChange_NoRaw_NoCodeChange
    Test_05_03_KindOfCodeChange_NoRaw_UsedOnly
    Test_05_04_KindOfCodeChange_NoRaw_UsedOnly
    mErH.EoP ErrSrc(PROC)
    
End Sub

    
Public Sub Test()
    With New FileSystemObject
        Debug.Print "File-Path          : " & .GetFile(ThisWorkbook.FullName).PATH
        Debug.Print "File-Name          : " & .GetFile(ThisWorkbook.FullName).name
        Debug.Print "File-BaseName      : " & .GetBaseName(.GetFile(ThisWorkbook.FullName).PATH)
        Debug.Print "File-Extension     : " & .GetExtensionName(.GetFile(ThisWorkbook.FullName).PATH)
        Debug.Print "File-Parent-Folder : " & .GetParentFolderName(.GetFile(ThisWorkbook.FullName))
    End With
End Sub

Public Sub Test_01_KindOfComp()
    Const PROC = "Test_01_KindOfComp"

    Dim wb          As Workbook
    Dim fso         As New FileSystemObject
    Dim cComp       As clsComp
    Dim sComp       As String
    
    Set wb = mWrkbk.GetOpen(fso.GetParentFolderName(ThisWorkbook.PATH) & "\File\File.xlsm")

    sComp = "mFile"
    Set cComp = Nothing
    Set cComp = New clsComp
    With cComp
        .Wrkbk = wb
        .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enKindOfComp.enRawClone
    End With

    sComp = "fMsg"
    Set cComp = Nothing
    Set cComp = New clsComp
    With cComp
        .Wrkbk = wb
        .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enRawClone
    End With
    
    sComp = "mTest"
    Set cComp = Nothing
    Set cComp = New clsComp
    With cComp
        .Wrkbk = wb
        .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enInternal
    End With
    
xt: wb.Close SaveChanges:=False
    Set cComp = Nothing
    Set fso = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
    
End Sub

Public Sub Test_05_01_KindOfCodeChange_NoRaw_UsedOnly()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_05_01_KindOfCodeChange_NoRaw_UsedOnly"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim wbTest      As Workbook
    Dim sExpFile    As String
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Begin of test preparation
    With cTest
        .TestProcedure = ThisWorkbook.name & ": " & ErrSrc(PROC)
        .TestItem = ThisWorkbook.name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        Set wbTest = mWrkbk.GetOpen(ThisWorkbook.PATH & "\" & "Test\Test1.xlsm")
        .ResultExpected = True
        .Details = "Export File does not exist, CodeModule not/never exported"
    End With
    
    With cComp
        .Wrkbk = wbTest
        .VBComp = .Wrkbk.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If fso.FileExists(sExpFile) Then fso.DeleteFile (sExpFile)
        
        ' Test: Code is regarded changed because there is no export file
        mErH.BoP ErrSrc(PROC)
        Debug.Assert .KindOfCodeChange = enUsedOnly
        mErH.EoP ErrSrc(PROC)
    
    End With
        
xt: Cleanup exp_file:=sExpFile, vbc:=cComp.VBComp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_05_02_KindOfCodeChange_NoRaw_NoCodeChange()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_05_02_KindOfCodeChange_NoRaw_NoCodeChange"
    
    On Error GoTo eh
    Dim sExpFile    As String
    Dim vResult     As Variant
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Test preparation
    With cTest
        .TestProcedure = ThisWorkbook.name & ": " & ErrSrc(PROC)
        .TestItem = ThisWorkbook.name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        Set wbTest = mWrkbk.GetOpen(ThisWorkbook.PATH & "\" & "Test\Test1.xlsm")
        .ResultExpected = False
        .Details = "Export File is identical with CodeModule"
    End With
    
    With cComp
        .Wrkbk = wbTest
        .VBComp = .Wrkbk.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If .VBComp.CodeModule.Lines(1, 1) = TEST_CHANGE Then .VBComp.CodeModule.DeleteLines 1, 1
        .VBComp.Export sExpFile
    End With
    
    ' Test: Code has not changed because it is identical with the export file
    mErH.BoP ErrSrc(PROC)
    Debug.Assert cComp.KindOfCodeChange = enNoCodeChange
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

Public Sub Test_05_03_KindOfCodeChange_NoRaw_UsedOnly()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_05_03_KindOfCodeChange_NoRaw_UsedOnly"
    
    On Error GoTo eh
    Dim sExpFile    As String
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Test preparation
    With cTest
        .TestProcedure = ThisWorkbook.name & "." & ErrSrc(PROC)
        .TestItem = ThisWorkbook.name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        .Details = "Export File outdated, CodeModule changed (additional line)"
        .ResultExpected = True
    End With
    Set wbTest = mWrkbk.GetOpen(ThisWorkbook.PATH & "\" & "Test\Test1.xlsm")
    
    With cComp
        .Wrkbk = wbTest
        .VBComp = .Wrkbk.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If .VBComp.CodeModule.Lines(1, 1) = TEST_CHANGE Then .VBComp.CodeModule.DeleteLines 1, 1
        .VBComp.Export sExpFile ' Overwrites any existing
        .VBComp.CodeModule.InsertLines 1, TEST_CHANGE
    End With
    
    '~~ Test: Code is regarded changed because it is not identical with the Export File
    mErH.BoP ErrSrc(PROC)
    Debug.Assert cComp.KindOfCodeChange = enUsedOnly
    mErH.EoP ErrSrc(PROC)
               
xt: Cleanup exp_file:=sExpFile, vbc:=cComp.VBComp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_05_04_KindOfCodeChange_NoRaw_UsedOnly()
'-----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "Test_05_04_KindOfCodeChange_NoRaw_UsedOnly"
    
    On Error GoTo eh
    Dim sExpFile    As String
    Dim vResult     As Variant
    
    If cTest Is Nothing Then Set cTest = New clsTestService
    If cComp Is Nothing Then Set cComp = New clsComp
    
    '~~ Test preparation
    With cTest
        .TestProcedure = ThisWorkbook.name & "." & ErrSrc(PROC)
        .TestItem = ThisWorkbook.name & ".clsComp.CodeChanged"
        .TestedByTheWay = "clsComp.ExportFileFullName"
        .Details = "Additional empty line in CodeModule is not considered a change though the Export File is outdated"
        .ResultExpected = False
    End With
    Set wbTest = mWrkbk.GetOpen(ThisWorkbook.PATH & "\" & "Test\Test1.xlsm")
    
    With cComp
        .Wrkbk = wbTest
        .VBComp = .Wrkbk.VBProject.VBComponents("mTest")
        sExpFile = .ExportFileFullName
        If .VBComp.CodeModule.Lines(1, 1) = TEST_CHANGE Then .VBComp.CodeModule.DeleteLines 1, 1
        .VBComp.Export sExpFile ' Overwrites any existing
        .VBComp.CodeModule.InsertLines 1, vbLf
    End With
    
    '~~ Test: Code is regarded changed because it is not identical with the Export File
    mErH.BoP ErrSrc(PROC)
    Debug.Assert cComp.KindOfCodeChange(ignore_empty_lines:=True) = enUsedOnly
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

Public Sub Test_10_ExportChangedComponents()
    Const PROC = "Test_ExportChangedComponents"
    
    mErH.BoP ErrSrc(PROC)
    mCompMan.ExportChangedComponents ThisWorkbook
    mErH.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_CodeModuleTrimm()

    Dim vbc As VBComponent
    Dim wb  As Workbook
    
    Set wb = ActiveWorkbook
    Set vbc = wb.VBProject.VBComponents("mCommon")
    Debug.Print "Trim CodeModule 'mCommon' in Workbook '" & wb.name & "'"
    mVBP.CodeModuleTrim vbc, wb
    
    Set wb = ThisWorkbook
    Set vbc = wb.VBProject.VBComponents("mCommon")
    Debug.Print "Trim CodeModule 'mCommon' in Workbook '" & wb.name & "'"
    mVBP.CodeModuleTrim vbc, wb
    
End Sub

Public Sub Test_CompOriginHasChanged()

End Sub

Public Sub Test_File_sAreEqual()

    Debug.Assert _
    mFile.sAreEqual( _
                  fc_file1:="E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas" _
                , fc_file2:="E:\Ablage\Excel VBA\DevAndTest\Common\CompManDev\mFile.bas" _
                 ) = False
    
'    Debug.Assert _
'    mFile.sAreEqual( _
'                  fc_file1:="E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas" _
'                , fc_file2:="E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas" _
'                 ) = True

End Sub

Public Sub Test_File_Compare()
    
    Const FILE_LEFT = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas"
    Const FILE_RIGHT = "E:\Ablage\Excel VBA\DevAndTest\Common\CompManDev\mFile.bas"
    
    Debug.Print mFile.Compare(file_left_full_name:=FILE_LEFT _
                          , file_right_full_name:=FILE_RIGHT _
                          , file_left_title:=FILE_LEFT _
                          , file_right_title:=FILE_RIGHT _
                           )
    
    Debug.Print mFile.Compare(file_left_full_name:=FILE_LEFT _
                          , file_right_full_name:=FILE_LEFT _
                          , file_left_title:=FILE_LEFT _
                          , file_right_title:=FILE_LEFT _
                           )
    
End Sub

Public Sub Test_File_SectionNames()
    Const PROC = "Test_File_SectionNames"

    On Error GoTo eh
    Dim v   As Variant
    
    For Each v In mFile.SectionNames(sn_file:=mCfg.CompManAddinPath & "\CompMan.dat")
        Debug.Print "[" & v & "]"
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Test_File_ValueNames()
    Const PROC = "Test_File_ValueNames"

    On Error GoTo eh
    Dim v   As Variant
    
    For Each v In mFile.ValueNames(vn_file:=mCfg.CompManAddinPath & "\CompMan.dat")
        Debug.Print """" & v & """"
    Next v
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Test_File_Values()
    Const PROC = "Test_File_Values"
    
    On Error GoTo eh
    Dim dctValues   As Dictionary
    Dim v           As Variant
    Dim sFile       As String
    
    sFile = mCfg.CompManAddinPath & "\CompMan.dat"
    Set dctValues = mFile.Values(vl_file:=sFile _
                             , vl_section:=mFile.SectionNames(sn_file:=sFile).Items()(0))
    For Each v In dctValues
        Debug.Print v & " = " & dctValues(v)
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Test_File_Value()
' ------------------------------------------------
' This test relies on the Value (Let) service.
' ------------------------------------------------
    Const PROC = "Test_File_Value"
    Const vbTemporaryFolder = 2
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim sFile           As String
    Dim sComment        As String
    Dim cyValue         As Currency: cyValue = 12345.6789
    
    '~~ Test preparation
    sFile = fso.GetSpecialFolder(SpecialFolder:=vbTemporaryFolder) & "\" & fso.GetTempName
        
    mErH.BoP ErrSrc(PROC)
    
    '~~ Test step 1: Write commented values
    sComment = "My comment"
    mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-1", vl_comment:=sComment) = "Test Value"
    sComment = "This is a boolean True"
    mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-2", vl_comment:=sComment) = True
    sComment = "This is a boolean False"
    mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-3", vl_comment:=sComment) = False
    sComment = "This is a currency value"
    mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-4") = cyValue
    
    '~~ Display written test values
    mMsg.Box msg_title:="Content of file '" & sFile & "'" _
           , msg:=fso.OpenTextFile(sFile).ReadAll _
           , msg_monospaced:=True
    
    '~~ Test step 2: Read commented values
    sComment = vbNullString
    Debug.Print "Test.Value-1 = '" & mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-1", vl_comment:=sComment) & "'"
    Debug.Assert mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-1", vl_comment:=sComment) = "Test Value"
    Debug.Assert sComment = "My comment"
    
    sComment = vbNullString
    Debug.Assert mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-2", vl_comment:=sComment) = True
    Debug.Assert sComment = "This is a boolean True"
    
    sComment = vbNullString
    Debug.Assert mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-3", vl_comment:=sComment) = False
    Debug.Assert sComment = "This is a boolean False"
    
    sComment = vbNullString
    Debug.Assert mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-4", vl_comment:=sComment) = cyValue
    Debug.Assert sComment = vbNullString
    Debug.Assert VarType(mFile.value(vl_file:=sFile, vl_section:="Test", vl_value_name:="Test.Value-4", vl_comment:=sComment)) = vbCurrency
    
    mErH.EoP ErrSrc(PROC)
    
xt: '~~ Test cleanup
    With fso
        If .FileExists(sFile) Then .DeleteFile (sFile)
    End With
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Test_File_Sections_Transfer_By_LetGet()
' ------------------------------------------------
' This test relies on the Value (Let) service.
' ------------------------------------------------
    Const PROC = "Test_File_Sections_Transfer_By_LetGet"
    Const vbTemporaryFolder = 2
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim sFileGet        As String
    Dim sFileLet        As String
    Dim i               As Long
    Dim j               As Long
    Dim arSections()    As Variant
    Dim sSectionName    As String
    
    '~~ Test preparation
    sFileGet = fso.GetSpecialFolder(SpecialFolder:=vbTemporaryFolder) & "\" & fso.GetTempName
    sFileLet = fso.GetSpecialFolder(SpecialFolder:=vbTemporaryFolder) & "\" & fso.GetTempName
    
    For i = 1 To 3
        sSectionName = "Section-" & i
        ReDim Preserve arSections(i - 1)
        arSections(i - 1) = sSectionName
        For j = 1 To 5
            mFile.value(vl_file:=sFileGet _
                    , vl_section:=sSectionName _
                    , vl_value_name:="Value-" & j _
                     ) = CStr(i & "-" & j)
        Next j
    Next i
    
    '~~ Test
    mErH.BoP ErrSrc(PROC)
    
    mFile.SectionsCopy sc_section_names:=arSections, sc_file_from:=sFileGet, sc_file_to:=sFileLet
    Debug.Assert mFile.sDiffer(dif_file1:=fso.GetFile(sFileGet), dif_file2:=fso.GetFile(sFileLet)) = False
    
    mErH.EoP ErrSrc(PROC)
    
xt: '~~ Test cleanup
    With fso
        If .FileExists(sFileGet) Then .DeleteFile (sFileGet)
        If .FileExists(sFileGet) Then .DeleteFile (sFileLet)
    End With
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
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

Public Sub Test_Temp()
    Const PROC = "Test_Temp"
    
    On Error GoTo eh
    
    Set wbTest = ThisWorkbook
    Set vbc = wbTest.VBProject.VBComponents("mFile")
    Set vbcm = vbc.CodeModule
    
    Debug.Print vbc.Properties(1)
    Debug.Print vbcm.Parent.Properties(1)
    
    Cleanup

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_UpdateRawClonesTheRemoteRawHasChanged()
    Const PROC  As String = "Test_UpdateCommonModules"
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    
    Application.StatusBar = vbNullString
    mCompMan.UpdateClonesTheRawHasChanged ThisWorkbook
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_Log()
    Const PROC = "Test_Log"
    
    Dim cLog As New clsLog
    With cLog
        .Reset
        .Service = ErrSrc(PROC)
        .Serviced = ThisWorkbook.name & ": " & "Mine"
        .Action = "Tested"
    End With
    Set cLog = Nothing

End Sub

