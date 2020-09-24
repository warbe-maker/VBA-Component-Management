Attribute VB_Name = "mTest"
Option Explicit
Private Module As Variant

Private Sub Test_Temp()
Dim wb      As Workbook
Dim vbc     As VBComponent
Dim vbcm    As CodeModule

    Set wb = ThisWorkbook
    Set vbc = wb.VBProject.VBComponents("mFile")
    Set vbcm = vbc.CodeModule
    
    Debug.Print vbc.Properties(1)
    Debug.Print vbcm.Parent.Properties(1)
    
End Sub
    
Private Sub Test_ExportChangedComponents()
Const PROC = "Test_ExportChangedComponents"
    BoP ErrSrc(PROC)
    mCompMan.ExportChangedComponents ThisWorkbook
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_CodeChanged()
Const PROC      As String = "Test_CodeChanged"
Dim cComp       As New clsComp
Dim wbSource    As Workbook
Dim wbTarget    As Workbook
Dim vbcSource   As VBComponent
Dim vbcTarget   As VBComponent
Dim flSource    As File
Dim flTarget    As File

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    With cComp
        .Host = ThisWorkbook
    
        '~~ Test 1: Code has not changed
        .VBComp = .Host.VBProject.VBComponents("fMsg")
        Debug.Assert .CodeChanged = False
    End With

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Test_UpdateUsedCommCompsTheOriginHasChanged()
Const PROC  As String = "Test_UpdateCommonModules"
Dim wb      As Workbook

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    Application.StatusBar = vbNullString
    Set wb = ThisWorkbook
    mCompMan.UpdateUsedCommCompsTheOriginHasChanged wb
    
exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
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
'Private Sub Test_Remove()
'' ----------------------------------------------------------------
'' Precondition: Export and Import are already successfully tested.
'' This test is pretty much the same what ReorganizeComponent does.
'' ----------------------------------------------------------------
'Dim sComp       As String
'Dim wbTarget    As Workbook
'Dim vbc         As VBComponent
'Dim sFileName   As String
'
'    Set wbTarget = OpenWb("DevTestDemo_mCompMan.xlsm")
'    sComp = "mForRemoveTest"
'
'    ' Export the to-be-removed module
'    Debug.Assert ComponentExists(wbTarget, sComp)
'    sFileName = mExport.CodeBackUpFullName(wbTarget, sComp)
'    mExport.Component wbTarget, sComp, sFileName
'
'    Set vbc = wbTarget.VBProject.VBComponents(sComp)
'    mCompMan.Remove vbc, wbTarget
'    Debug.Assert Not ComponentExists(wbTarget, sComp)
'
'    '~~ Reset Workbok to its initial test status
'   mCompMan.Import sFileName, wbTarget
'    Debug.Assert ComponentExists(wbTarget, sComp)
'
'End Sub

Private Sub Test_DisplayDiff()
Dim cComp           As New clsComp

    With cComp
        .Host = ThisWorkbook
        .VBComp = .Host.VBProject.VBComponents("mBasic")
        .DisplayDiff .ExportFile, .ExportFile, "Test title left", "Test title right"
    End With
End Sub

Private Sub Test_CleanExportFile()
Dim cComp As New clsComp

    With cComp
        .Host = ThisWorkbook
        .VBComp = ThisWorkbook.VBProject.VBComponents("mFile")
        .CleanExportFile .ExportFile
    End With
End Sub

Private Sub Test_CodeModuleTrimm()
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
