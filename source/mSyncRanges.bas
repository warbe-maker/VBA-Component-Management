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
    Dim v           As Variant
    Dim rngSource   As Range
    Dim rngTarget   As Range
    Dim nm          As Name
    Dim sSheet      As String
    Dim sName       As String
    Dim ws          As Worksheet
    Dim dct         As Dictionary
    
    Stats.Count sic_named_ranges_total, Sync.TargetNames.Count
    Set dct = Sync.TargetNames
    For Each v In dct
        If Not Sync.SourceNames.Exists(v) Then GoTo next_v
        
        '~~ This is a range name which exists in the source and the target Workbook
        '~~ The range's formatting is synchronized
        Stats.Count sic_named_ranges
        sSheet = Replace(Split(dct(v), "!")(0), "=", vbNullString)
        sName = v
        Set ws = Sync.Target.Worksheets(sSheet)
        Set rngTarget = Sync.Target.Worksheets(sSheet).Range(sName)
        Set rngSource = Sync.Source.Worksheets(sSheet).Range(sName)
        Log.ServicedItem = rngTarget
        Debug.Print "Source range: " & rngSource.Name.Name, Tab(45), Sync.Source.Name & dct(v) & vbLf & _
                    "Target range: " & rngTarget.Name.Name, Tab(45), Sync.Target.Name & dct(v)
        SyncProperties rngSource, rngTarget
next_v:
    Next v
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function

Private Sub SyncProperties( _
                     ByRef sf_source As Range, _
                     ByRef sf_target As Range)
    
    With sf_target
        SyncProperty .AddComment, sf_source.AddComment, "AddComment"
        SyncProperty .AddIndent, sf_source.AddIndent, "AddIndent"
        SyncProperty .Borders.Color, sf_source.Borders.Color, "Borders.Color"
        SyncProperty .Borders.ColorIndex, sf_source.Borders.ColorIndex, "Borders.ColorIndex"
        SyncProperty .Borders.LineStyle, sf_source.Borders.LineStyle, "Borders.LineStyle"
        SyncProperty .Borders.ThemeColor, sf_source.Borders.ThemeColor, "Borders.ThemeColor"
        SyncProperty .Borders.TintAndShade, sf_source.Borders.TintAndShade, "Borders.TintAndShade"
        SyncProperty .Borders.Value, sf_source.Borders.Value, "Borders.Value"
        SyncProperty .Borders.Value, sf_source.Borders.Value, "Borders.Value"
        SyncProperty .ColumnWidth, sf_source.ColumnWidth, "ColumnWidth"
'        SyncProperty .Comment.xxx, sf_source.Comment.xxx, "Comment.xxx"
'        SyncProperty .Comment.xxx, sf_source.Comment.xxx, "Comment.xxx"
'        SyncProperty .Comment.xxx, sf_source.Comment.xxx, "Comment.xxx"
'        SyncProperty .Comment.xxx, sf_source.Comment.xxx, "Comment.xxx"
        SyncProperty .COMMENT.Visible, sf_source.COMMENT.Visible, "Comment.Visible"



'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
'        SyncProperty .xxx, sf_source.xxx, "xxx"
    End With

End Sub

Private Sub SyncProperty( _
                   ByRef v_target As Variant, _
                   ByRef v_source As Variant, _
                   ByVal s_property As String)
' ---------------------------------------------
' Synchronizes a single Shape or OLEObject
' property by skipping those not applicable.
' ---------------------------------------------
    On Error GoTo xt
    If v_target <> v_source Then
        On Error Resume Next ' The property may not be modifyable
        v_target = v_source
        If Err.Number = 0 _
        Then Log.Entry = "Property synched '" & s_property & "'" _
        Else Log.Entry = "Property synchronization failed (error " & Err.Number & ") '" & s_property & "'"
    End If

xt:
End Sub

