Attribute VB_Name = "mService"
Option Explicit

Public Function Denied(ByVal wb As Workbook) As Boolean
' -----------------------------------------------------
' Returns True when all preconditions for a service
' execution are fulfilled.
' -----------------------------------------------------

    If WbkIsRestoredBySystem(wb) Then
        Denied = True
    ElseIf Not WbkInServicedRoot(wb) Then
        Denied = True
    ElseIf mMe.AddInPaused Then
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


