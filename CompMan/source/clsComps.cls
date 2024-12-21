VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dctAll              As Dictionary

Public Property Get All() As Dictionary:            Set All = dctAll:                           End Property

Private Function Changed() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all components of which the code has been
' modified. The Dictionary has the VBComponent objects as key and the
' corresponding clsComp objects as item. The function uses the All Dictionary
' and removes all items where the components not indicates it has changed.
' ----------------------------------------------------------------------------
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    
    Set dct = dctAll
    For Each v In dct
        Set Comp = dct(v)
        If Not Comp.Changed Then
            dct.Remove v
        End If
    Next v
    Set Changed = dct
    
End Function

