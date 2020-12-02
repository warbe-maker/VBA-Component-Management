Attribute VB_Name = "mCompLog"
Option Explicit

Private Const SBJCT_FILE_NAME   As String = "\Update.log"
Private Log                     As clsAppData
Private lLogEntry               As Long

Private Sub InitLog()
    If Log Is Nothing Then
        Set Log = New clsAppData
        With Log
            .Location = spp_File
            .Extension = ".log"
            .Subject = mConfig.CommonComponentsBasePath & SBJCT_FILE_NAME
        End With
    End If
End Sub

Public Sub LogAction(ByVal sWbName As String, _
                     ByVal sLog As String)
' ---------------------------------------------
' Log the action (sLog) started at (dtLog).
' ---------------------------------------------
    InitLog
    With Log
        .Aspect = sWbName
        lLogEntry = lLogEntry + 1
        .ValueLet sName:=sWbName & "_" & Format(lLogEntry, "00"), vValue:=sLog
    End With
End Sub

                       
Public Sub StartLog(ByRef sWbName As String)
    InitLog
    Log.AspectRemove sAspect:=sWbName
    lLogEntry = 0 ' Reset dat entry counter
End Sub

