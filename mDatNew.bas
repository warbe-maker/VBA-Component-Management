Attribute VB_Name = "mDatNew"
Option Explicit

Public Function IsRawHost(ByVal wbk_name As String) As Boolean
    IsRawHost = mHost.Exists(mBasic.BaseName(wbk_name))
End Function

Public Function IsCloneHost(ByVal wbk_name As String) As Boolean
    IsCloneHost = CloneWorkbooks.Exists(wbk_name)
End Function

