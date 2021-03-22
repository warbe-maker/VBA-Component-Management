Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mCompManClient
'                 Optionally used by any Workbook to:
'                 - automatically update used Common Components (hosted,
'                   developed, tested, and provided, by another Workbook)
'                   with the Workbook_open event
'                 - automatically export any changed VBComponent with
'                   the Workbook_Before_Save event.
'
' W. Rauschenberger, Berlin March 18 2021
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------------

Public Function CompManService(ByVal service As String, ByVal hosted As String) As Boolean
' ----------------------------------------------------------------------------
' Execution of the CompMan service (service) preferrably via the CompMan-Addin
' or when not available alternatively via the CompMan's development instance.
' ----------------------------------------------------------------------------
    Const COMPMAN_BY_ADDIN = "CompMan.xlam!mCompMan."
    Const COMPMAN_BY_DEVLP = "CompMan.xlsb!mCompMan."
    
    On Error Resume Next
    Application.Run COMPMAN_BY_ADDIN & service, ThisWorkbook, hosted
    If Err.Number = 1004 Then
        On Error Resume Next
        Application.Run COMPMAN_BY_DEVLP & service, ThisWorkbook, hosted
        If Err.Number = 1004 Then
            Application.StatusBar = "'" & service & "' neither available by '" & COMPMAN_BY_ADDIN & "' nor by '" & COMPMAN_BY_DEVLP & "'!"
        End If
    End If
End Function

