VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LmAt As String
Public LmBy As String
Public LmIn As String
Public LmKo As String
Public LmOn As String
Public LmXf As String

Public Function DifferFrom(ByVal d_prprts As clsProperties) As Boolean

    DifferFrom = True
    With d_prprts
        Select Case True
            Case LmAt <> .LmAt
            Case LmBy <> .LmBy
            Case LmIn <> .LmIn
            Case .LmKo <> vbNullString And LmKo <> .LmKo
            Case LmOn <> .LmOn
            Case .LmXf <> vbNullString And LmXf <> .LmXf
            Case Else
                DifferFrom = False
        End Select
    End With
    
End Function
