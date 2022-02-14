VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbSyncTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const HOSTED_RAWS = ""
Private Const COMPMAN_BY_ADDIN  As String = "CompMan.xlam!mCompMan."
Private Const COMPMAN_BY_DEVLP  As String = "CompMan.xlsb!mCompMan."

Private Sub Workbook_Open()
    
    '~~ ------------------------------------------------------------------
    '~~ CompMan Workbook_Open service 'UpdateRawClones':
    '~~ Executed by the Addin *) or via the development instance when open
    '~~ *) automatically available only when referenced by the VB-Project
    CompManService "UpdateRawClones"
    '~~ ------------------------------------------------------------------
    
xt: Exit Sub

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    '~~ ------------------------------------------------------------------
    '~~ CompMan Workbook_Before_Save service 'ExportChangedComponents':
    '~~ Executed by the Addin *) or via the development instance when open
    '~~ *) automatically available only when referenced by the VB-Project
    CompManService "ExportChangedComponents"
    '~~ ------------------------------------------------------------------

xt: Exit Sub

End Sub

Private Function CompManService(ByVal service As String) As Boolean
' ----------------------------------------------------------------------------
' Execution of the CompMan service (service) preferrably via the CompMan-Addin
' or when not available alternatively via the CompMan's development instance.
' ----------------------------------------------------------------------------
    On Error Resume Next
    Application.Run COMPMAN_BY_ADDIN & service, ThisWorkbook, HOSTED_RAWS
    If Err.Number = 1004 Then
        On Error Resume Next
        Application.Run COMPMAN_BY_DEVLP & service, ThisWorkbook, HOSTED_RAWS
        If Err.Number = 1004 Then
            Application.StatusBar = "CompMan-Service: '" & service & "' neither available by '" & COMPMAN_BY_ADDIN & "' nor by '" & COMPMAN_BY_DEVLP & "' available!"
        End If
    End If
End Function
