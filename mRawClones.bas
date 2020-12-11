Attribute VB_Name = "mRawClones"
Option Explicit
' --------------------------------------------------------------
' Standard Module mRawsCloned Maintains the RawClones.dat file.
'
' ---------------------------------------------------------------------

Public Property Get CloneWorkbooks() As Dictionary
'    Set CloneWorkbooks =
End Property

Private Property Get FileName() As String
    FileName = mCfg.CompManAddinPath & "\RawsCloned.dat"
End Property

