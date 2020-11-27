Attribute VB_Name = "mTest"
Option Explicit
Option Private Module

Public Sub Test_Temp()

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
    
    mErH.BoP ErrSrc(PROC)
    mCompMan.ExportChangedComponents ThisWorkbook
    mErH.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_CodeChanged()
    Const PROC      As String = "Test_CodeChanged"
    
    Dim cComp       As New clsComp

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    
    With cComp
        .Host = ThisWorkbook
    
        '~~ Test 1: Code has not changed
        .VBComp = .Host.VBProject.VBComponents("fMsg")
        Debug.Assert .CodeChanged = False
    End With

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
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
    
eh: mErH.ErrMsg ErrSrc(PROC)
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
