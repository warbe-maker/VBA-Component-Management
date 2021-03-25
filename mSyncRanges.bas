Attribute VB_Name = "mSyncRanges"
Option Explicit
' ---------------------------------------------------
' Standard Module mSyncRngsFrmt
'          Synchronize the formating of named ranges.
'
' ---------------------------------------------------
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRanges." & s
End Function

Public Function SyncFormating() As Boolean
' ----------------------------------------
' Synchronizes the formating of all named
' ranges. Returns True when at least one
' format has been synched.
' ----------------------------------------
    Const PROC = "SyncFormating"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim rng     As Range
    Dim nm      As Name
    Dim sSheet  As String
    Dim sName   As String
    Dim ws      As Worksheet
    Dim dct     As Dictionary
    
    Stats.Count sic_named_ranges_total, Sync.TargetNames.Count
    Set dct = Sync.TargetNames
    For Each v In dct
        If Not Sync.SourceNames.Exists(v) Then GoTo next_v
        Stats.Count sic_named_ranges
        sSheet = Replace(Split(dct(v), "!")(0), "=", vbNullString)
        sName = v
        Set ws = Sync.Target.Worksheets(sSheet)
        Set rng = ws.Range(sName)
        Debug.Print sName & " = " & rng.Address
next_v:
    Next v
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function

