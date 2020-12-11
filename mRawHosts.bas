Attribute VB_Name = "mRawHosts"
Option Explicit
' --------------------------------------------------------------
' Standard Module mHosts Maintains the Raw  Hosts.dat file.
'
' --------------------------------------------------------------

Private Property Get FileName(Optional ByVal fl_path As String) As String
    FileName = fl_path & "\RawsHosted.dat"
End Property

Public Property Get RawHostWorkbooks(Optional ByVal hr_file As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with the name of any Workbook hosting a raw component.
' The Workbook name is the key and its full name is the item.
' ----------------------------------------------------------------------------
End Property




