VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sService    As String
Private sServiced   As String
Private sFile       As String
Private tsFile      As TextStream

Public Property Let Service(ByVal s As String)
    sService = ThisWorkbook.name & "/" & s
    sFile = ThisWorkbook.PATH & "\" & s & ".log"
    Reset
End Property

Public Property Let Serviced(ByVal s As String):    sServiced = s:                              End Property

Public Property Let Action(ByVal s As String)
' -------------------------------------------
' Appen an Action line to the log file.
' -------------------------------------------
    Dim fso As New FileSystemObject
    
    If Not fso.FileExists(sFile) Then fso.CreateTextFile sFile
    Set tsFile = fso.OpenTextFile(sFile, IOMode:=ForAppending)
    tsFile.WriteLine Now() & " " & sServiced & ": " & s
    tsFile.Close
    Set fso = Nothing
    Set tsFile = Nothing

End Property


Public Sub Reset()
    Dim fso As New FileSystemObject
    With fso
        If .FileExists(sFile) Then
            .DeleteFile sFile
        End If
        .CreateTextFile sFile
    End With
    Set fso = Nothing
End Sub

Private Sub Class_Initialize()
    
    Dim fso As New FileSystemObject
    
    sFile = ThisWorkbook.PATH & "\" & fso.GetBaseName(ThisWorkbook.FullName) & ".log"
    If Not fso.FileExists(sFile) Then fso.CreateTextFile (sFile)
    Set fso = Nothing
    
End Sub
