Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient, optionally used by any Workbook to:
' - update used 'Common-Components' (hosted, developed, tested,
'   and provided, by another Workbook) with the Workbook_open event
' - export any changed VBComponent with the Workbook_Before_Save event.
'
' W. Rauschenberger, Berlin March 2021
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------
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

Public Sub CompManService(ByVal cm_service As String, _
                          ByVal hosted As String)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (cm_service) preferably via the CompMan
' Development instance when available (assuming it is for testing). Only when
' not available the CompMan AddIn services (CompMan.xlam) are used.
' ----------------------------------------------------------------------------
    Const COMPMAN_BY_ADDIN = "CompMan.xlam!mCompMan."
    Const COMPMAN_BY_DEVLP = "CompMan.xlsb!mCompMan."
    Dim Done As Boolean
    
    On Error Resume Next
    Done = Application.Run(COMPMAN_BY_ADDIN & cm_service, ThisWorkbook, hosted)
    If Err.Number = 1004 Or Not Done Then
        On Error Resume Next
        Application.Run COMPMAN_BY_DEVLP & cm_service, ThisWorkbook, hosted
        If Err.Number = 1004 Then
            Application.StatusBar = "'" & cm_service & "' neither available by '" & COMPMAN_BY_ADDIN & "' nor by '" & COMPMAN_BY_DEVLP & "'!"
        End If
    End If

xt: Exit Sub

End Sub

