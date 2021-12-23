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
' format has been synced.
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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SyncNamedColumnsWidth( _
                    Optional ByRef ws_source As Worksheet = Nothing, _
                    Optional ByRef ws_target As Worksheet = Nothing)
    Const PROC = "SyncWidthNamedColumns"
    
    On Error GoTo eh
    Dim rngSource   As Range
    Dim rngTarget   As Range
    Dim i           As Long
    Dim ws          As Worksheet
    Dim sSheetName  As String
    Dim RangeName   As String
    
    If ws_source Is Nothing And ws_target Is Nothing Then
        For Each ws In Sync.Source.Worksheets
            If mSyncSheets.SheetExists(Wb:=Sync.Target _
                                     , sh1_name:=ws.Name _
                                     , sh1_code_name:=ws.CodeName _
                                     , sh2_name:=sSheetName _
                                      ) _
            Then
                Set ws_target = Sync.Target.Worksheets(sSheetName)
                Set ws_source = ws
                With ws
                    For i = 1 To .UsedRange.Columns.CountLarge
                        Set rngSource = ws_source.Columns.Item(i)
                        Set rngTarget = ws_target.Columns.Item(i)
                        On Error Resume Next
                        RangeName = rngSource.Name.Name
                        If Err.Number = 0 Then
                            '~~ This is a named column
                            On Error GoTo eh
                            rngTarget.EntireColumn.ColumnWidth = rngSource.ColumnWidth
                        End If
                    Next i
                End With
            End If
        Next ws
    Else
        With ws_source
            For i = 1 To .UsedRange.Columns.CountLarge
                Set rngSource = ws_source.Columns.Item(i)
                Set rngTarget = ws_target.Columns.Item(i)
                On Error Resume Next
                RangeName = rngSource.Name.Name
                If Err.Number = 0 Then
                    '~~ This is a named range (row)
                    On Error GoTo eh
                    rngTarget.EntireColumn.ColumnWidth = rngSource.ColumnWidth
                End If
            Next i
        End With
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncNamedRowsHeight( _
                    Optional ByRef ws_source As Worksheet = Nothing, _
                    Optional ByRef ws_target As Worksheet = Nothing)
' --------------------------------------------------------------------
'
' --------------------------------------------------------------------
    Const PROC = "SyncNamedRowsHeight"
    
    On Error GoTo eh
    Dim rngSource   As Range
    Dim rngTarget   As Range
    Dim i           As Long
    Dim ws          As Worksheet
    Dim sSheetName  As String
    Dim RangeName   As String
    
    If ws_source Is Nothing And ws_target Is Nothing Then
        For Each ws In Sync.Source.Worksheets
            If mSyncSheets.SheetExists(Wb:=Sync.Target _
                                     , sh1_name:=ws.Name _
                                     , sh1_code_name:=ws.CodeName _
                                     , sh2_name:=sSheetName _
                                      ) _
            Then
                Set ws_target = Sync.Target.Worksheets(sSheetName)
                Set ws_source = ws
                With ws
                    For i = 1 To .UsedRange.Rows.CountLarge
                        Set rngSource = ws_source.Rows.Item(i)
                        Set rngTarget = ws_target.Rows.Item(i)
                        On Error Resume Next
                        RangeName = rngSource.Name.Name
                        If Err.Number = 0 Then
                            '~~ This is a named row
                            On Error GoTo eh
                            rngTarget.EntireRow.RowHeight = rngSource.RowHeight
                        End If
                    Next i
                End With
            End If
        Next ws
    Else
        With ws_source
            For i = 1 To .UsedRange.Rows.CountLarge
                Set rngSource = ws_source.Rows.Item(i)
                Set rngTarget = ws_target.Rows.Item(i)
                On Error Resume Next
                RangeName = rngSource.Name.Name
                If Err.Number = 0 Then
                    '~~ This is a named range (row)
                    On Error GoTo eh
                    rngTarget.EntireRow.RowHeight = rngSource.RowHeight
                End If
            Next i
        End With
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub



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
        Then Log.Entry = "Property synced '" & s_property & "'" _
        Else Log.Entry = "Property synchronization failed (error " & Err.Number & ") '" & s_property & "'"
    End If

xt:
End Sub

