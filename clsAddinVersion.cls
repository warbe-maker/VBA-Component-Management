VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAddinVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim sVersion As String
Property Get Version() As String:           Version = sVersion: End Property
Property Let Version(ByVal s As String):    sVersion = s:       End Property
