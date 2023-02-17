Attribute VB_Name = "mSyncRanges"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mSyncRanges: Provides all means/services for the synchroni-
'                              sation of named Range's formattings.
' Public services:
' - AllDone
' - CollectAllItems
' - CollectChanged
' - HasNames
' - SyncId
' - Sync
' - SyncAllNamedRanges
'
' ------------------------------------------------------------------------------
Private Const TITLE_SYNC_RANGES As String = "VB-Project Synchronization: Named Ranges"

'Public Sub CollectAllItems()
'' ------------------------------------------------------------------------------
'' Writes the Named Range potentially synchronizedto the wsSynch sheet.
'' ------------------------------------------------------------------------------
'    Const PROC = "CollectAllItems"
'
'    On Error GoTo eh
'    Dim dct         As New Dictionary
'    Dim nme         As Name
'    Dim v           As Variant
'
'    mBasic.BoP ErrSrc(PROC)
'
'    For Each nme In mSync.Source.Names
'        If IsValidUserRangeName(nme) Then
'            If Not dct.Exists(SyncId(nme)) _
'            Then mDct.DctAdd dct, SyncId(nme), nme, order_bykey, seq_ascending, sense_casesensitive
'        End If
'    Next nme
'    For Each nme In mSync.TargetWorkingCopy.Names
'        If IsValidUserRangeName(nme) Then
'            If Not dct.Exists(SyncId(nme)) _
'            Then mDct.DctAdd dct, SyncId(nme), nme, order_bykey, seq_ascending, sense_casesensitive
'        End If
'    Next nme
'
'    For Each v In dct
'        wsSync.RngItemAll(v) = True
'    Next v
'    Set dct = Nothing
'
'xt: mBasic.EoP ErrSrc(PROC)
'    Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'
'End Sub
'
'Public Sub SyncNamedColumnsWidth( _
'                    Optional ByRef wsh_source As Worksheet = Nothing, _
'                    Optional ByRef wsh_target As Worksheet = Nothing)
'    Const PROC = "SyncWidthNamedColumns"
'
'    On Error GoTo eh
'    Dim rngSource   As Range
'    Dim rngTarget   As Range
'    Dim i           As Long
'    Dim wsh          As Worksheet
'    Dim sSheetName  As String
'    Dim RangeName   As String
'
'    If wsh_source Is Nothing And wsh_target Is Nothing Then
'        For Each wsh In mSync.Source.Worksheets
'            If mSyncSheets.Exists(se_wbk:=sync_wbk_target _
'                                , se_wsh_name:=ws.Name _
'                                , se_wsh_code_name:=ws.CodeName _
'                                , se_wsh_name:=sSheetName _
'                                      ) _
'            Then
'                Set wsh_target = sync_wbk_target.Worksheets(sSheetName)
'                Set wsh_source = wsh
'                With wsh
'                    For i = 1 To .UsedRange.Columns.CountLarge
'                        Set rngSource = wsh_source.Columns.Item(i)
'                        Set rngTarget = wsh_target.Columns.Item(i)
'                        On Error Resume Next
'                        RangeName = rngSource.Name.Name
'                        If Err.Number = 0 Then
'                            '~~ This is a named column
'                            On Error GoTo eh
'                            rngTarget.EntireColumn.ColumnWidth = rngSource.ColumnWidth
'                        End If
'                    Next i
'                End With
'            End If
'        Next wsh
'    Else
'        With wsh_source
'            For i = 1 To .UsedRange.Columns.CountLarge
'                Set rngSource = wsh_source.Columns.Item(i)
'                Set rngTarget = wsh_target.Columns.Item(i)
'                On Error Resume Next
'                RangeName = rngSource.Name.Name
'                If Err.Number = 0 Then
'                    '~~ This is a named range (row)
'                    On Error GoTo eh
'                    rngTarget.EntireColumn.ColumnWidth = rngSource.ColumnWidth
'                End If
'            Next i
'        End With
'    End If
'
'xt: Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

'Public Sub SyncNamedRowsHeight(ByVal srh_rng_source As Range, _
'                               ByVal srh_rng_target As Range)
'' --------------------------------------------------------------------
''
'' --------------------------------------------------------------------
'    Const PROC = "SyncNamedRowsHeight"
'
'    On Error GoTo eh
'    Dim rngSource   As Range
'    Dim rngTarget   As Range
'    Dim i           As Long
'    Dim wsh          As Worksheet
'    Dim sSheetName  As String
'    Dim RangeName   As String
'
'    If wsh_source Is Nothing And wsh_target Is Nothing Then
'        For Each wsh In mSync.Source.Worksheets
'            If mSyncSheets.Exists(se_wbk:=sync_wbk_target _
'                                , se_wsh_name:=ws.Name _
'                                , se_wsh_code_name:=ws.CodeName _
'                                , se_wsh_name:=sSheetName _
'                                 ) _
'            Then
'                Set wsh_target = sync_wbk_target.Worksheets(sSheetName)
'                Set wsh_source = wsh
'                With wsh
'                    For i = 1 To .UsedRange.Rows.CountLarge
'                        Set rngSource = wsh_source.Rows.Item(i)
'                        Set rngTarget = wsh_target.Rows.Item(i)
'                        On Error Resume Next
'                        RangeName = rngSource.Name.Name
'                        If Err.Number = 0 Then
'                            '~~ This is a named row
'                            On Error GoTo eh
'                            rngTarget.EntireRow.RowHeight = rngSource.RowHeight
'                        End If
'                    Next i
'                End With
'            End If
'        Next wsh
'    Else
'        With wsh_source
'            For i = 1 To .UsedRange.Rows.CountLarge
'                Set rngSource = wsh_source.Rows.Item(i)
'                Set rngTarget = wsh_target.Rows.Item(i)
'                On Error Resume Next
'                RangeName = rngSource.Name.Name
'                If Err.Number = 0 Then
'                    '~~ This is a named range (row)
'                    On Error GoTo eh
'                    rngTarget.EntireRow.RowHeight = rngSource.RowHeight
'                End If
'            Next i
'        End With
'    End If
'
'xt: Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

'Public Function CollectChanged() As Dictionary
'' -----------------------------------------------------------------------------
'' Returns a collection of those named Ranges in the Sync-Source-Workbook of
'' which any property differs from the correponding range in the Sync-Target-
'' Workbook.
'' -----------------------------------------------------------------------------
'    Const PROC = "CollectChanged"
'
'    On Error GoTo eh
'    Dim dct         As New Dictionary
'    Dim nme         As Name
'    Dim TargetRange As Range
'    Dim SourceRange As Range
'    Dim SourceSheet As Worksheet
'    Dim v           As Variant
'    Dim sProgress   As String
'
'    sProgress = " "
'    mService.EstablishServiceLog mCompManClient.SRVC_SYNCHRONIZE
'
'    For Each nme In mSync.Source.Names
'        mSync.MonitorStep "Collecting Named Ranges which changed"
'        If IsValidUserRangeName(nme) Then
'            If CollectChangedStillRelevant(nme) Then
'                If Exists(nme, mSync.TargetWorkingCopy, TargetRange) Then
'                    Set SourceSheet = ReferredSheet(mSync.Source, nme)
'                    Set SourceRange = SourceSheet.Range(nme.Name)
'                    If PropertiesChanged(nme, SourceSheet, SourceRange, TargetRange) Then
'                    '~~ At least one of the range's properties differs bettween source and target
'                        mDct.DctAdd dct, SyncId(nme), nme, order_bykey, seq_ascending, sense_casesensitive
'                    End If
'                End If
'            End If
'        End If
'        sProgress = sProgress & "."
'    Next nme
'
'xt: Set CollectChanged = dct
'    Set dct = Nothing
'    Exit Function
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function
'
'Private Function CollectChangedStillRelevant(ByVal nme As Name) As Boolean
'    With wsSync
'        CollectChangedStillRelevant = _
'        IsSyncItem(nme) _
'        And Not .RngItemChangedDone(SyncId(nme)) _
'        And Not .RngItemChangedFailed(SyncId(nme))
'    End With
'End Function

' ------------------------------------------------------------------------------
' Standard Module mSyncRngsFrmt: Synchronize the formating of named ranges.
' ------------------------------------------------------------------------------
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRanges." & s
End Function

Private Function Exists(ByVal ex_nme As Name, _
                        ByVal ex_wbk As Workbook, _
               Optional ByRef ex_rng_result As Range) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Range Name (ex_name) exists in the Workbook
' (ex_wbk) disregarding the range it refers to and returns thhe Name object
' (ex_name_result).
' Note: Any non-relevant Range Name (like _xlfn.COUNTIFS for instance) are
' ignored, i.e. return FALSE.
' ------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim nme  As Name
    Dim wsh  As Worksheet
    
    For Each nme In ex_wbk.Names
        If nme.Name = ex_nme.Name And IsValidUserRangeName(nme) Then
            Exists = True
            Set wsh = ReferredSheet(ex_wbk, nme)
            Set ex_rng_result = wsh.Range(nme.Name)
            Exit Function
        End If
    Next nme
                        
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function SyncId(ByVal nme As Name) As String
    SyncId = "'" & nme.RefersTo & " (" & nme.Name & ")"
End Function

Private Function PropertiesChanged(ByVal sync_nme As Name, _
                                   ByVal sync_wsh As Worksheet, _
                                   ByVal sync_rng_source As Range, _
                                   ByVal sync_rng_target As Range) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when any of the Range's (sync_rng_target) properties differs
' from the corresponding Ranges's (sync_rng_target) properties.
' ------------------------------------------------------------------------------
    
    On Error Resume Next
    With sync_rng_target
        If PropertyChanged(sync_nme, sync_wsh, .AddComment, sync_rng_source.AddComment, "AddComment") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .AddIndent, sync_rng_source.AddIndent, "AddIndent") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.Color, sync_rng_source.Borders.Color, "Borders.Color") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.ColorIndex, sync_rng_source.Borders.ColorIndex, "Borders.ColorIndex") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.LineStyle, sync_rng_source.Borders.LineStyle, "Borders.LineStyle") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.ThemeColor, sync_rng_source.Borders.ThemeColor, "Borders.ThemeColor") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.TintAndShade, sync_rng_source.Borders.TintAndShade, "Borders.TintAndShade") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.Value, sync_rng_source.Borders.Value, "Borders.Value") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Borders.Value, sync_rng_source.Borders.Value, "Borders.Value") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .ColumnWidth, sync_rng_source.ColumnWidth, "ColumnWidth") Then PropertiesChanged = True
        If PropertyChanged(sync_nme, sync_wsh, .Comment.Visible, sync_rng_source.Comment.Visible, "Comment.Visible") Then PropertiesChanged = True
    End With
    On Error GoTo -1
    
End Function

Private Function PropertyChanged(ByVal pc_nme As Name, _
                                 ByVal pc_sheet As Worksheet, _
                                 ByVal pc_property_target As Variant, _
                                 ByVal pc_property_source As Variant, _
                                 ByVal pc_property_name As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the property (pc_property_target) differs from the property
' (pw_property_arget).
' ------------------------------------------------------------------------------
    Const PROC = "PropertyChanged"
    
    On Error GoTo eh
    PropertyChanged = pc_property_target <> pc_property_source
    If PropertyChanged Then
        mService.Log.ServicedItem("Named Range") = "Sheet(" & pc_sheet.Name & ").Range(" & pc_nme.Name & ").Property;" & pc_property_name
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ReferredSheet(ByVal rs_wbk As Workbook, _
                               ByVal rs_nme As Name) As Worksheet
' ------------------------------------------------------------------------
' Returns the Worksheet in the Workbook (rs_wbk) referred by the range name
' (rs_nme).
' ------------------------------------------------------------------------
    Dim wsh     As Worksheet
    Dim sSheet  As String
    
    sSheet = ReferredSheetName(rs_nme)
    For Each wsh In rs_wbk.Worksheets
        If wsh.Name = sSheet Then
            Set ReferredSheet = wsh
            Exit For
        End If
    Next wsh

End Function

Private Function ReferredSheetName(ByVal rs_nme As Name) As String
    ReferredSheetName = Replace(Replace(Split(rs_nme.RefersTo, "!")(0), "=", vbNullString), "'", vbNullString)
End Function

Public Sub Sync(ByRef sync_changed As Dictionary)
' ------------------------------------------------------------------------------
' Called by mSync.RunSync provided there are (still) any outstanding named
' range synchronizations are to be done. Displays them in a mode-less dialog
' for being confirmed either one by one or all together at a time.
' ------------------------------------------------------------------------------
    Const PROC = "Sync"
    
    On Error GoTo eh
    Dim AppRunArgs  As New Dictionary
    Dim cllButtons  As New Collection
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim SyncsDue    As Long
    Dim v           As Variant
                       
    SyncsDue = sync_changed.Count
    If SyncsDue = 0 Then GoTo xt
    
    '~~ There's at least one Reference still to be synchronized
    wsService.SyncDialogTitle = TITLE_SYNC_RANGES
    Set fSync = mMsg.MsgInstance(TITLE_SYNC_RANGES)
    With Msg.Section(1)
        .Label.Text = "Named Ranges properties had changed:"
        .Label.FontColor = rgbDarkGreen
        .Text.MonoSpaced = True
        For Each v In sync_changed
            .Text.Text = .Text.Text & vbLf & v
        Next v
    End With
    
    With Msg.Section(2)
        .Label.Text = "About synchronization of Named Ranges:"
        .Text.Text = "CompMan's 'Named Ranges Synchronization' service is essential for positioning shapes correctly - in case there are any to be synchronized."
    End With
    With Msg.Section(8).Label
        .Text = "See in README chapter CompMan's VB-Project-Synchronization service (GitHub README):"
        .FontColor = rgbBlue
        .OpenWhenClicked = mCompMan.README_URL & mCompMan.README_SYNC_CHAPTER
    End With
        
    '~~ Add an additional 'Synchronize All' button when there is more than one item to be synched
    Set cllButtons = mMsg.Buttons(cllButtons, SYNC_ALL_BTTN)
    mMsg.ButtonAppRun AppRunArgs, SYNC_ALL_BTTN _
                                , ThisWorkbook _
                                , "mSyncRanges.SyncAllNamedRanges"

    
    '~~ Display the mode-less dialog for the confirmation which Sheet synchronization to run
    mMsg.Dsply dsply_title:=TITLE_SYNC_RANGES _
             , dsply_msg:=Msg _
             , dsply_buttons:=cllButtons _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=AppRunArgs _
             , dsply_width_min:=45 _
             , dsply_pos:=wsService.SyncDialogTop & ";" & wsService.SyncDialogLeft
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Sub SyncAllNamedRanges(ByVal sync_wsh_name As String, _
'                              ByVal sync_rng_name As String)
'' ------------------------------------------------------------------------------
'' Synchronizes all properties of the target range (sync_rng_target) with the
'' corresponding property of the source range (sync_source_range).
'' Precondition: Worksheet name changes are synchronized!
'' ------------------------------------------------------------------------------
'    Dim rngSource   As Range
'    Dim rngTarget   As Range
'    Dim wshSource   As Worksheet
'    Dim wshTarget   As Worksheet
'
'    mService.EstablishServiceLog mCompManClient.SRVC_SYNCHRONIZE
'
'    Set wshSource = mSync.Source.Worksheets(sync_wsh_name)
'    Set wshTarget = mSync.TargetWorkingCopy.Worksheets(sync_wsh_name)
'    Set rngSource = wshSource.Range(sync_rng_name)
'    Set rngTarget = wshTarget.Range(sync_rng_name)
'
'    mService.Log.ServicedItem = rngTarget
'
'    On Error Resume Next
'    With rngTarget
'        SyncRangeProperty .AddIndent, rngSource.AddIndent, "AddIndent"
'        SyncRangeProperty .Borders.Color, rngSource.Borders.Color, "Borders.Color"
'        SyncRangeProperty .Borders.ColorIndex, rngSource.Borders.ColorIndex, "Borders.ColorIndex"
'        SyncRangeProperty .Borders.LineStyle, rngSource.Borders.LineStyle, "Borders.LineStyle"
'        SyncRangeProperty .Borders.ThemeColor, rngSource.Borders.ThemeColor, "Borders.ThemeColor"
'        SyncRangeProperty .BorderAround, rngSource.BorderAround, "BorderAround"
'        SyncRangeProperty .Borders.TintAndShade, rngSource.Borders.TintAndShade, "Borders.TintAndShade"
'        SyncRangeProperty .Borders.Value, rngSource.Borders.Value, "Borders.Value"
'        SyncRangeProperty .Borders.Weight, rngSource.Borders.Weight, "Borders.Weight"
'        SyncRangeProperty .ColumnWidth, rngSource.ColumnWidth, "ColumnWidth"
'        SyncRangeProperty .EntireColumn.ColumnWidth, rngSource.EntireColumn.ColumnWidth, "EntireColumn.ColumnWidth"
'        SyncRangeProperty .EntireRow.RowHeight, rngSource.EntireRow.RowHeight, "EntireRow.RowHeight"
'        SyncRangeProperty .Font.Background, rngSource.Font.Background, "Font.Background"
'        SyncRangeProperty .Font.Bold, rngSource.Font.Bold, "Font.Bold"
'        SyncRangeProperty .Font.Color, rngSource.Font.Color, "Font.Color"
'        SyncRangeProperty .Font.ColorIndex, rngSource.Font.ColorIndex, "Font.ColorIndex"
'        SyncRangeProperty .Font.FontStyle, rngSource.Font.FontStyle, "Font.FontStyle"
'        SyncRangeProperty .Font.Italic, rngSource.Font.Italic, "Font.Italic"
'        SyncRangeProperty .Font.Name, rngSource.Font.Name, "Font.Name"
'        SyncRangeProperty .Font.Size, rngSource.Font.Size, "Font.Size"
'        SyncRangeProperty .Font.Strikethrough, rngSource.Font.Strikethrough, "Font.Strikrough"
'        SyncRangeProperty .Font.Subscript, rngSource.Font.Subscript, "Font.Subscript"
'        SyncRangeProperty .Font.Superscript, rngSource.Font.Superscript, "Font.Superscript"
'        SyncRangeProperty .Font.ThemeColor, rngSource.Font.ThemeColor, "Font.ThemeColor"
'        SyncRangeProperty .Font.ThemeFont, rngSource.Font.ThemeFont, "Font.ThemeFont"
'        SyncRangeProperty .Font.TintAndShade, rngSource.Font.TintAndShade, "Font.TintandShade"
'        SyncRangeProperty .Font.Underline, rngSource.Font.Underline, "Font.Underline"
'    End With
'
'    wsSync.NmeItemChangedDone(SyncId(GetName(sync_rng_name))) = True
'
'    '~~ Re-display the synchronization dialog for still to be synchronized types/Items
'    mSync.MessageUnload TITLE_SYNC_RANGES
'    mSync.RunSync
'
'End Sub

Public Sub SyncRangeProperty(ByVal v_target As Variant, _
                             ByVal v_source As Variant, _
                             ByVal s_property As String)
' ------------------------------------------------------------------------------
' Synchronizes a the property value (v_target) with the corresponding property
' (v_source) ignoring it when not applicable.
' ------------------------------------------------------------------------------
    
    If v_target <> v_source Then
        On Error Resume Next
        v_target = v_source
        If Err.Number = 0 _
        Then Log.Entry = "Property synced '" & s_property & "'" _
        Else Log.Entry = "Property synchronization failed (error " & Err.Number & ") '" & s_property & "'"
        On Error GoTo -1
    End If

xt: On Error GoTo -1

End Sub

