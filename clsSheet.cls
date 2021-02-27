VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private wsh         As Worksheet
Private wbk         As Workbook

Friend Property Set Sheet(ByRef sh As Worksheet):   Set wsh = sh:               End Property

Friend Property Get Sheet() As Worksheet:           Set Sheet = wsh:            End Property

Friend Property Set Wrkbk(ByRef wb As Workbook):    Set wbk = wb:               End Property

Friend Property Get Wrkbk() As Workbook:            Set Wrkbk = wbk:            End Property

Friend Property Get CodeName( _
                   Optional ByVal scn_name As String) As String
' -------------------------------------------------------------
' Returns the sheet's CodeName through its Name.
' -------------------------------------------------------------
    Dim sh As Worksheet
    For Each sh In wbk.Worksheets
        If sh.Name = scn_name Then
            CodeName = sh.CodeName
            Exit For
        End If
    Next sh
End Property

Friend Property Get Name( _
          Optional ByVal sn_codename As String) As String
' -------------------------------------------------------
' Returns the sheet's name through its CodeName.
' -------------------------------------------------------
    Dim sh As Worksheet
    For Each sh In wbk.Worksheets
        If sh.CodeName = sn_codename Then
            Name = sh.Name
            Exit For
        End If
    Next sh
End Property

Public Function ExistsByCodeName(ByVal eb_codename As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Me.Wrkbk
        ExistsByCodeName = ws.CodeName = eb_codename
        If ExistsByCodeName Then Exit For
    Next ws
End Function

Public Function ExistsByName(ByVal eb_name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Me.Wrkbk
        ExistsByName = ws.Name = eb_name
        If ExistsByName Then Exit For
    Next ws
End Function

