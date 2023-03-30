VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------
' Class Module clsStats
' CompMan services statistics
' ----------------------------------
Private aCount() As Long

Private Sub Clear()
    ReDim aCount(0)
    aCount(0) = 0
End Sub

Public Sub Count(Optional ByVal c_sic As siCounter, _
                  Optional ByVal c_value As Long = 1)
' -------------------------------------------------------------
' Adds the (value) to the array aCount under the index (c_sic).
' -------------------------------------------------------------
    If UBound(aCount) < c_sic Then
        ReDim Preserve aCount(c_sic)
    End If
    aCount(c_sic) = aCount(c_sic) + c_value
End Sub

Private Sub Class_Initialize()
    Clear
End Sub
