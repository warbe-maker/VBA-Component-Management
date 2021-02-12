Attribute VB_Name = "mRenew"
Option Explicit

Public Sub ByImport(ByVal rn_wb As Workbook, _
                    ByVal rn_comp_name As String, _
                    ByVal rn_exp_file_full_name As String, _
           Optional ByVal rn_status As String = vbNullString)
' -----------------------------------------------------------
' Renews the component (rn_comp_name) in Workbook (rn_wb)
' by importing the Export File (rn_exp_file_full_name).
' -----------------------------------------------------------
    Dim sTempName       As String
    Dim sExpFilePath    As String
    Dim fso             As New FileSystemObject

    Debug.Print NowMsec & " =========================="
    If rn_status <> vbNullString Then Application.StatusBar = rn_status & "Save and wait"
    SaveWbk rn_wb
    DoEvents:  Application.Wait Now() + 0.0000001 ' wait for 10 milliseconds
    With rn_wb.VbProject
        If CompExists(ce_wb:=rn_wb, ce_comp_name:=rn_comp_name) Then
            '~~ Find a free/unused temporary name
            sTempName = GetTempName(ac_wb:=rn_wb, ac_comp_name:=rn_comp_name)
            '~~ Rename the component when it already exists
            If rn_status <> vbNullString Then Application.StatusBar = rn_status & "Rename '" & rn_comp_name & "' to '" & sTempName & "'"
            .VBComponents(rn_comp_name).name = sTempName
            Debug.Print NowMsec & " '" & rn_comp_name & "' renamed to '" & sTempName & "'"
'           DoEvents:  Application.Wait Now() + 0.0000001 ' wait for 10 milliseconds
            If rn_status <> vbNullString Then Application.StatusBar = rn_status & "Remove '" & sTempName & "'"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Debug.Print NowMsec & " '" & sTempName & "' removed (may be postponed by the system however)"
        End If
    
        '~~ (Re-)import the component
        If rn_status <> vbNullString Then Application.StatusBar = rn_status & "(Re-)import '" & rn_exp_file_full_name & "'"
        .VBComponents.Import rn_exp_file_full_name
        Debug.Print NowMsec & " '" & rn_comp_name & "' (re-)imported from '" & rn_exp_file_full_name & "'"
        sExpFilePath = rn_wb.Path & "\" & rn_comp_name & Extension(ext_wb:=rn_wb, ext_comp_name:=rn_comp_name)
        If Not fso.FileExists(sExpFilePath) Or rn_exp_file_full_name <> sExpFilePath Then
            If rn_status <> vbNullString Then Application.StatusBar = rn_status & "Export '" & rn_comp_name & "' to '" & sExpFilePath & "'"
            .VBComponents(rn_comp_name).Export sExpFilePath
            Debug.Print NowMsec & " '" & rn_comp_name & "' exported to '" & sExpFilePath & "'"
        End If
    End With
          
    Set fso = Nothing
    
End Sub

Private Sub SaveWbk(ByVal rs_wb As Workbook)
    Application.EnableEvents = False
    rs_wb.Save
    Application.EnableEvents = True
End Sub

Private Function GetTempName(ByVal ac_wb As Workbook, _
                             ByVal ac_comp_name As String) As String
' Return a temporary name for a component not already existing
' ------------------------------------------------------------------
    Dim sTempName   As String
    Dim i           As Long
    Dim vbc         As VBComponent
    
    sTempName = ac_comp_name & "_Temp"
    Do
        On Error Resume Next
        sTempName = ac_wb.VbProject.VBComponents(sTempName).name
        If Err.Number <> 0 Then Exit Do ' a component with sTempName does not exist
        i = i + 1: sTempName = sTempName & i
    Loop
    GetTempName = sTempName
End Function

Private Function WbkIsOpen(ByVal io_wb_full_name As String) As Boolean
' Retuns True when the Workbook (io_full_name) is open.
' --------------------------------------------------------------------
    Dim fso     As New FileSystemObject
    Dim xlApp   As Excel.Application
    
    If Not fso.FileExists(io_wb_full_name) Then Exit Function
    On Error Resume Next
    Set xlApp = GetObject(io_wb_full_name).Application
    WbkIsOpen = Err.Number = 0

End Function

Private Function WbkGetOpen(ByVal go_wb_full_name) As Workbook
    
    Dim fso     As New FileSystemObject
    Dim sWbName As String
    
    If Not fso.FileExists(go_wb_full_name) Then Exit Function
    If WbkIsOpen(go_wb_full_name) Then
        Set WbkGetOpen = Application.Workbooks(sWbName)
    Else
        Set WbkGetOpen = Application.Workbooks.Open(go_wb_full_name)
    End If

    Set fso = Nothing

End Function

Private Function CompExists(ByVal ce_wb As Workbook, _
                            ByVal ce_comp_name As String) As Boolean
' Returns TRUE when the component (ce_comp_name) exists in the
' Workbook ce_wb.
' ------------------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = ce_wb.VbProject.VBComponents(ce_comp_name).name
    CompExists = Err.Number = 0
End Function

Private Function Extension(ByVal ext_wb As Workbook, _
                           ByVal ext_comp_name As String) As String
' Returns the components Export File extension
' -----------------------------------------------------------------
    Select Case ext_wb.VbProject.VBComponents(ext_comp_name).Type
        Case vbext_ct_StdModule:    Extension = ".bas"
        Case vbext_ct_ClassModule:  Extension = ".cls"
        Case vbext_ct_MSForm:       Extension = ".frm"
        Case vbext_ct_Document:     Extension = ".cls"
    End Select
End Function

Private Property Get NowMsec() As String
    NowMsec = Format(Now(), "hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Property

