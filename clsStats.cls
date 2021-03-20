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
Private lMaxLen As Long

Public Sub Clear()
    ReDim aCount(0)
    aCount(0) = 0
End Sub

Public Property Get AlignedValue(ByVal sic As siCounter) As String
    Dim v       As Variant
    Dim value   As Long
    
    If lMaxLen = 0 Then
        For Each v In aCount
            lMaxLen = Max(lMaxLen, Len(CStr(v)))
        Next v
    End If
    
    value = Me.Total(sic)
    If sic = sic_non_doc_mod_total Then
        value = Me.Total(sic_vbcomps_total) - Me.Total(sic_sheets_total) - 1
    End If
    
    AlignedValue = Align(CStr(value), lMaxLen, AlignRight)
End Property

Public Property Get Total(Optional ByVal sic As siCounter) As Long
    If UBound(aCount) < sic Then Total = 0 Else Total = aCount(sic)
End Property

Public Sub Count( _
  Optional ByVal sic As siCounter, _
  Optional ByVal value As Long = 1)
' ----------------------------------
' Adds the (value) to the array
' aCount under the index (sic).
' ----------------------------------
    If UBound(aCount) < sic Then
        ReDim Preserve aCount(sic)
    End If
    aCount(sic) = aCount(sic) + value
End Sub


Private Sub Class_Initialize()
    Clear
End Sub
