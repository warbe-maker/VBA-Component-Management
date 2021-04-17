VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsSyncTestB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Property Get ReadWrite(Optional property_row As Long) As String
    ReadWrite = Me.Range("colReadWrite").Cells(property_row).Value
End Property

Public Property Let ReadWrite(Optional property_row As Long, ByVal property_readwrite As String)
    Me.Range("colReadWrite").Cells(property_row).Value = property_readwrite
End Property
