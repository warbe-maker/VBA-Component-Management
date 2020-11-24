Attribute VB_Name = "mCompLog"
Option Explicit

Private Const SBJCT_FILE_NAME   As String = "\Update.log"
Private log                     As clsAppData
Private lLogEntry               As Long

Private Sub InitLog()
    If log Is Nothing Then
        Set log = New clsAppData
        With log
            .Location = spp_File
            .Extension = ".log"
            .Subject = mCommDat.CommonComponentsBasePath & SBJCT_FILE_NAME
        End With
    End If
End Sub

Public Sub LogAction(ByVal sWbName As String, _
                     ByVal sLog As String)
' ---------------------------------------------
' Log the action (sLog) started at (dtLog).
' ---------------------------------------------
    InitLog
    With log
        .Aspect = sWbName
        lLogEntry = lLogEntry + 1
        .ValueLet sName:=sWbName & "_" & Format(lLogEntry, "00"), vValue:=sLog
    End With
End Sub

                       
Public Sub StartLog(ByRef sWbName As String)
    InitLog
    log.AspectRemove sAspect:=sWbName
    lLogEntry = 0 ' Reset dat entry counter
End Sub

