Attribute VB_Name = "mService"
Option Explicit

Public Function Denied(ByVal wb As Workbook) As Boolean
' -----------------------------------------------------
' Returns True when all preconditions for a service
' execution are fulfilled.
' -----------------------------------------------------

    If WbkIsRestoredBySystem(wb) Then
        cLog.Action = "Service denied for a Workbook restored by the system!"
        Denied = True
    ElseIf Not WbkInServicedRoot(wb) Then
        cLog.Action = "Service denied for a Workbook outside the configured 'ServicedRoot': " & mMe.RootServicedByCompMan & "!"
        Denied = True
    ElseIf mMe.AddInPaused Then
        cLog.Action = "Service denied because the CompMan Addin is currently paused!"
        Denied = True
    ElseIf FolderNotVbProjectExclusive(wb) Then
        cLog.Action = "Service denied because the Workbook is not an exclusive VB-Project in folder '" & wb.Path & "'!"
        Denied = True
    End If

End Function

Private Function WbkInServicedRoot(ByVal idr_wb As Workbook) As Boolean
    WbkInServicedRoot = InStr(idr_wb.Path, mMe.RootServicedByCompMan) <> 0
End Function

Private Function WbkIsRestoredBySystem(ByVal rbs_wb As Workbook) As Boolean
    WbkIsRestoredBySystem = InStr(ActiveWindow.Caption, "(") <> 0 _
                         Or InStr(rbs_wb.FullName, "(") <> 0
End Function

Private Function FolderNotVbProjectExclusive(ByVal wb As Workbook) As Boolean

    Dim fso As New FileSystemObject
    Dim fl  As File
    
    For Each fl In fso.GetFolder(wb.Path).Files
        If fl.Path <> wb.FullName And VBA.Left$(fso.GetFileName(fl.Path), 2) <> "~$" Then
            Select Case fso.GetExtensionName(fl.Path)
                Case "xlsm", "xlam", "xlsb" ' may contain macros, a VB-Project repsectively
                    FolderNotVbProjectExclusive = True
                    Exit For
            End Select
        End If
    Next fl

End Function

