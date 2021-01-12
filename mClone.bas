Attribute VB_Name = "mClone"
Option Explicit
' --------------------------------------------------------------
' Standard Module mRawsCloned Maintains the RawClones.dat file.
'
' ---------------------------------------------------------------------

Public Property Get CloneWorkbooks() As Dictionary
'    Set CloneWorkbooks =
End Property

Private Property Get DAT_FILE() As String
    DAT_FILE = mMe.CompManAddinPath & "\RawsCloned.dat"
End Property

