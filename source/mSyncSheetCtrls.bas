Attribute VB_Name = "mSyncSheetCtrls"
Option Explicit
' -------------------------------------------------------------------
' Standard Module mSyncSheetControls
'          Services to synchronize new, obsolete, and properties of
'          sheet shapes and OLEObjects.
'
' -------------------------------------------------------------------
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncSheetCtrls" & s
End Function

Private Sub CopyShape( _
                ByRef sc_source As Worksheet, _
                ByRef sc_target As Worksheet, _
                ByVal sc_name As String)
' ---------------------------------------------
'
' ---------------------------------------------
    Dim SourceShape As Shape
    Dim TargetShape As Shape
    
    For Each SourceShape In sc_source.Shapes
        If SourceShape.Name <> sc_name Then GoTo next_shape
        SourceShape.Copy
        sc_target.Paste
        Set TargetShape = sc_target.Shapes(sc_target.Shapes.Count)
        With TargetShape
            .Name = sc_name
            .top = SourceShape.top
            .Left = SourceShape.Left
            .Width = SourceShape.Width
            .Height = SourceShape.Height
        End With
next_shape:
    Next SourceShape
End Sub

Private Sub CopyOOB( _
                ByRef sc_source As Worksheet, _
                ByRef sc_target As Worksheet, _
                ByVal sc_oob_name As String)
' ---------------------------------------------
'
' ---------------------------------------------
    Dim SourceOOB As OLEObject
    Dim TargetOOB As OLEObject
    
    For Each SourceOOB In sc_source.OLEObjects
        If SourceOOB.Name <> sc_oob_name Then GoTo next_shape
        SourceOOB.Copy
        sc_target.Paste
        Set TargetOOB = sc_target.OLEObjects(sc_target.OLEObjects.Count)
        With TargetOOB
            .Name = sc_oob_name
            .top = SourceOOB.top
            .Left = SourceOOB.Left
            .Width = SourceOOB.Width
            .Height = SourceOOB.Height
        End With
next_shape:
    Next SourceOOB
End Sub


Private Function SheetOOBExists( _
                          ByRef sync_wb As Workbook, _
                          ByVal sync_sheet_name As String, _
                          ByVal sync_sheet_code_name As String, _
                          ByVal sync_oob_name As String) As Boolean
' -----------------------------------------------------------------
' Returns TRUE when the OLEObject (sync_oob_name) exists in the
' Workbook (sync_wb) in a sheet with either the given Name
' (sync_sheet_name) or the provided CodeName (sync_sheet_code_name).
' Explanation: When this function is used to get the required info for
'              being confirmed, the concerned sheet may be one of which
'              the Name or the CodeName is about to be renamed - which
'              by then will not have taken place.
' ----------------------------------------------------------------------
    Const PROC = "SheetOOBExists"
    
    On Error GoTo eh
    Dim ws  As Worksheet
    Dim oob As OLEObject
    
    For Each ws In sync_wb.Worksheets
        If ws.Name <> sync_sheet_name And ws.CodeName <> sync_sheet_code_name Then GoTo next_sheet
        '~~ Either the Name or the CodeName is the same. Enough to clearly identify the sheet.
        For Each oob In ws.OLEObjects
            If oob.Name = sync_oob_name Then
                SheetOOBExists = True
                GoTo xt
            End If
        Next oob
next_sheet:
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function






Private Function SheetShapeExists( _
                            ByRef sync_wb As Workbook, _
                            ByVal sync_sheet_name As String, _
                            ByVal sync_sheet_code_name As String, _
                            ByVal sync_shape_name As String) As Boolean
' ----------------------------------------------------------------------
' Returns TRUE when the shape (sync_shape_name) exists in the Workbook
' (sync_wb) in a sheet with either the given Name (sync_sheet_name) or
' the provided CodeName (sync_sheet_code_name).
' Explanation: When this function is used to get the required info for
'              being confirmed, the concerned sheet may be one of which
'              the Name or the CodeName is about to be renamed - which
'              by then will not have taken place.
' ----------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim ws  As Worksheet
    Dim shp As Shape
    
    For Each ws In sync_wb.Worksheets
        If ws.Name <> sync_sheet_name And ws.CodeName <> sync_sheet_code_name Then GoTo next_sheet
        For Each shp In ws.Shapes
            If shp.Name = sync_shape_name Then
                SheetShapeExists = True
                GoTo xt
            End If
        Next shp
next_sheet:
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function

Public Sub SyncOOBsNew()
' -----------------------------------------------------------
' Copy new shapes from the sourec Workbook (Sync.Source) to
' the target Workbook (Sync.Target) and ajust the properties
' -----------------------------------------------------------
    Const PROC = "SyncOOBsNew"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    Dim sControl    As String
    Dim sSheet      As String
    Dim oob         As OLEObject
    
    With Sync
        For Each v In .SourceSheetOOBs
            sSheet = mSync.KeySheetName(v)
            sControl = .SourceSheetOOBs(v)
            If SheetOOBExists(sync_wb:=.Target _
                            , sync_sheet_name:=sSheet _
                            , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                            , sync_oob_name:=sControl _
                               ) _
            Then GoTo next_oob
            '~~ The source shape not yet exists in the target Workbook's corresponding sheet
            '~~ (idetified either by its Name or CodeName) and thus is regarde new and needs
            '~~ to be copied and its Properties adjusted.
            Set wsSource = .Source.Worksheets(sSheet)
            Set oob = wsSource.OLEObjects(sControl)
            Log.ServicedItem = oob
            If Sync.Mode = Confirm Then
                '~~ New OLEObjects coming with new sheets are not displayed for confirmation
                If Not .IsNewSheet(sSheet) Then
                     Stats.Count sic_shapes_new
                    .ConfInfo = "New!"
                End If
            Else
                '~~ When new OLEObjects are syncronized the Worksheet's Name/CodeName will have been syncronized before!
                If Not .IsNewSheet(sSheet) Then
                    Set wsTarget = .Target.Worksheets(sSheet)
                    CopyOOB sc_source:=wsSource _
                            , sc_target:=wsTarget _
                            , sc_oob_name:=sControl
                    Log.Entry = "Copied from source sheet"
                End If
            End If
next_oob:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub




Public Sub SyncShapesNew()
' -----------------------------------------------------------
' Copy new shapes from the sourec Workbook (Sync.Source) to
' the target Workbook (Sync.Target) and ajust the properties
' -----------------------------------------------------------
    Const PROC = "SyncShapesNew"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    Dim sControl    As String
    Dim sSheet      As String
    Dim shp         As Shape
    
    With Sync
        For Each v In .SourceSheetShapes
            sSheet = mSync.KeySheetName(v)
            sControl = mSync.KeyControlName(v)
            If SheetShapeExists(sync_wb:=.Target _
                              , sync_sheet_name:=sSheet _
                              , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                              , sync_shape_name:=sControl _
                               ) _
            Then GoTo next_shape
            '~~ The source shape not yet exists in the target Workbook's corresponding sheet
            '~~ (idetified either by its Name or CodeName) and thus is regarde new and needs
            '~~ to be copied and its Properties adjusted.
            Set wsSource = .Source.Worksheets(sSheet)
            Set shp = wsSource.Shapes(sControl)
            Log.ServicedItem = shp
            If .Mode = Confirm Then
                '~~ New shapes coming with new sheets are not displayed for confirmation
                If Not .IsNewSheet(sSheet) Then
                     Stats.Count sic_oobs_new
                    .ConfInfo = "New!"
                End If
            Else
                '~~ When new shapes are syncronized the Worksheet's Name/CodeName will have been syncronized before!
                If Not .IsNewSheet(sSheet) Then
                    Set wsTarget = .Target.Worksheets(sSheet)
                    CopyShape sc_source:=wsSource _
                            , sc_target:=wsTarget _
                            , sc_name:=sControl
                    Log.Entry = "Copied from source sheet"
                End If
            End If
next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub


Public Sub SyncOOBsObsolete()
' -----------------------------------------------------------
' Remove obsolete shapes in the target Workbook (Sync.Target)
' -----------------------------------------------------------
    Const PROC = "SyncOOBsObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsTarget    As Worksheet
    Dim sControl    As String
    Dim sSheet      As String
    Dim oob         As OLEObject
    
    With Sync
        For Each v In .TargetSheetOOBs
            sSheet = mSync.KeySheetName(v)
            sControl = .TargetSheetOOBs(v)
            If SheetOOBExists(sync_wb:=.Source _
                              , sync_sheet_name:=KeySheetName(v) _
                              , sync_sheet_code_name:=SheetCodeName(.Target, sSheet) _
                              , sync_oob_name:=sControl _
                               ) _
            Then GoTo next_oob
            Set wsTarget = .Target.Worksheets(sSheet)
            Set oob = wsTarget.OLEObjects(sControl)
            Log.ServicedItem = oob
            
            Stats.Count sic_oobs_obsolete
            '~~ The target name does not exist in the source and thus  has become obsolete
            If .Mode = Confirm Then
                .ConfInfo = "Obsolete!"
            Else
                wsTarget.OLEObjects(sControl).Delete
                Log.Entry = "Deteted!"
            End If
next_oob:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub






Public Sub SyncShapesObsolete()
' -----------------------------------------------------------
' Remove obsolete shapes in the target Workbook (Sync.Target)
' -----------------------------------------------------------
    Const PROC = "SyncShapesObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsTarget    As Worksheet
    Dim sControl    As String
    Dim sSheet      As String
    Dim shp         As Shape
    
    With Sync
        For Each v In .TargetSheetShapes
            sSheet = mSync.KeySheetName(v)
            sControl = mSync.KeyControlName(v)
            If SheetShapeExists(sync_wb:=.Source _
                              , sync_sheet_name:=KeySheetName(v) _
                              , sync_sheet_code_name:=SheetCodeName(.Target, sSheet) _
                              , sync_shape_name:=sControl _
                               ) _
            Then GoTo next_shape
            Set wsTarget = .Target.Worksheets(sSheet)
            Set shp = wsTarget.Shapes(sControl)
            Log.ServicedItem = shp
            
            Stats.Count sic_shapes_obsolete
            '~~ The target name does not exist in the source and thus  has become obsolete
            If .Mode = Confirm Then
                .ConfInfo = "Obsolete!"
            Else
                wsTarget.Shapes(sControl).Delete
                Log.Entry = "Deteted!"
            End If
next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub SyncShapesProperties()
' -----------------------------------------
' Syncronize all shape's properties between
' source and target Workbook.
' -----------------------------------------
    Const PROC = "SyncShapesProperties"
        
    On Error GoTo eh
    Dim v           As Variant
    Dim sControl    As String
    Dim sSheet      As String
    
    With Sync
        For Each v In .SourceSheetShapes
            sSheet = mSync.KeySheetName(v)
            sControl = mSync.KeyControlName(v)
            If Not SheetShapeExists(sync_wb:=.Target _
                                  , sync_sheet_name:=sSheet _
                                  , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                                  , sync_shape_name:=sControl _
                                   ) _
            Then GoTo next_shape
            SyncShapeProperties .Source.Worksheets(sSheet).Shapes.Item(sControl) _
                              , .Target.Worksheets(sSheet).Shapes.Item(sControl)

next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub SyncOOBsProperties()
' --------------------------------------------------------------
' Syncronize all OLEObjects' properties for synchronized sheets.
' Note: Is is essential that the sheets Name and CodeName had
'       been synchronized beforehand.
' --------------------------------------------------------------
    Const PROC = "SyncOOBsProperties"
        
    On Error GoTo eh
    Dim v           As Variant
    Dim sControl    As String
    Dim sSheet      As String
    Dim vSource   As Variant
    Dim vTarget   As Variant
    
    With Sync
        For Each v In .SourceSheetOOBs
            sSheet = mSync.KeySheetName(v)
            sControl = mSync.KeyControlName(v)
            If Not SheetOOBExists(sync_wb:=.Target _
                                  , sync_sheet_name:=sSheet _
                                  , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                                  , sync_oob_name:=sControl _
                                   ) _
            Then GoTo next_shape
            Set vSource = .Source.Worksheets(sSheet).OLEObjects(sControl)
            Set vTarget = .Target.Worksheets(sSheet).OLEObjects(sControl)
            SyncOOBProperties vSource _
                            , vTarget

next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncShapePropertiesTextEffect( _
                                    ByRef te_source As TextEffectFormat, _
                                    ByRef te_target As TextEffectFormat)
' -------------------------------------------------------------------------
'
' -------------------------------------------------------------------------
    On Error Resume Next
    With te_target
        SyncProperty .Alignment, te_target.Alignment, "TextEffectFormat.Alignment"
        SyncProperty .FontBold, te_source.FontBold, "TextEffectFormat.FontBold"
        SyncProperty .FontItalic, te_source.FontItalic, "TextEffectFormat.FontItalic"
        SyncProperty .FontName, te_source.FontName, "TextEffectFormat.FontName"
        SyncProperty .FontSize, te_source.FontSize, "TextEffectFormat.FontSize"
        SyncProperty .KernedPairs, te_source.KernedPairs, "TextEffectFormat.KernedPairs"
        SyncProperty .NormalizedHeight, te_source.NormalizedHeight, "TextEffectFormat.NormalizedHeight"
        SyncProperty .PresetShape, te_source.PresetShape, "TextEffectFormat.PresetShape"
        SyncProperty .PresetTextEffect, te_source.PresetTextEffect, "TextEffectFormat.PresetTextEffect"
        SyncProperty .RotatedChars, te_source.RotatedChars, "TextEffectFormat.RotatedChars"
        SyncProperty .text, te_source.text, "TextEffectFormat.Text"
        SyncProperty .Tracking, te_source.Tracking, "TextEffectFormat.Tracking"
    End With
    On Error GoTo -1
    
End Sub

Private Sub SyncShapePropertiesTextFrame( _
                                   ByRef tf_source As TextFrame, _
                                   ByRef tf_target As TextFrame)
' ----------------------------------------------------------------
' Syncronize the target shape's TextFrame properties with those
' from the source shape. Any non applicable property is just
' skipped, syncronized properties are logged.
' ----------------------------------------------------------------
    With tf_target
        On Error Resume Next
        SyncProperty .AutoMargins, tf_source.AutoMargins, "TextFrame.AutoMargins"
        SyncProperty .AutoSize, tf_source.AutoSize, "TextFrame.AutoSize"
        SyncProperty .HorizontalAlignment, tf_source.HorizontalAlignment, "TextFrame.HorizontalAlignment"
        SyncProperty .HorizontalOverflow, tf_source.HorizontalOverflow, "TextFrame.HorizontalOverflow"
        SyncProperty .MarginBottom, tf_source.MarginBottom, "TextFrame.MarginBottom"
        SyncProperty .MarginLeft, tf_source.MarginLeft, "TextFrame.MarginLeft"
        SyncProperty .MarginRight, tf_source.MarginRight, "TextFrame.MarginRight"
        SyncProperty .MarginTop, tf_source.MarginTop, "TextFrame.MarginTop"
        SyncProperty .Orientation, tf_source.Orientation, "TextFrame.Orientation"
        SyncProperty .ReadingOrder, tf_source.ReadingOrder, "TextFrame.ReadingOrder"
        SyncProperty .VerticalAlignment, tf_source.VerticalAlignment, "TextFrame.VerticalAlignment"
        SyncProperty .VerticalOverflow, tf_source.VerticalOverflow, "TextFrame.VerticalOverflow"
    End With
    On Error GoTo -1
xt: Exit Sub

End Sub

Private Sub SyncShapePropertiesTextFrame2( _
                               ByRef tf2_source As TextFrame2, _
                               ByRef tf2_target As TextFrame2)
' --------------------------------------------------------------
' Syncronize the target shape's TextFrame2 properties with those
' from the source shape. Any non applicable property is just
' skipped, syncronized properties are logged.
' --------------------------------------------------------------
    On Error Resume Next
    With tf2_target
        SyncProperty .AutoSize, tf2_source.AutoSize, "TextFrame2.AutoSize"
        SyncProperty .HorizontalAnchor, tf2_source.HorizontalAnchor, "TextFrame2.HorizontalAnchor"
        SyncProperty .MarginBottom, tf2_source.MarginBottom, "TextFrame2.MarginBottom"
        SyncProperty .MarginLeft, tf2_source.MarginLeft, "TextFrame2.MarginLeft"
        SyncProperty .MarginRight, tf2_source.MarginRight, "TextFrame2.MarginRight"
        SyncProperty .MarginTop, tf2_source.MarginTop, "TextFrame2.MarginTop"
        SyncProperty .NoTextRotation, tf2_source.NoTextRotation, "TextFrame2.NoTextRotation"
        SyncProperty .Orientation, tf2_source.Orientation, "TextFrame2.Orientation"
        SyncProperty .PathFormat, tf2_source.PathFormat, "TextFrame2.PathFormat"
        SyncProperty .VerticalAnchor, tf2_source.VerticalAnchor, "TextFrame2.VerticalAnchor"
        SyncProperty .WarpFormat, tf2_source.WarpFormat, "TextFrame2.WarpFormat"
        SyncProperty .WordArtformat, tf2_source.WordArtformat, "TextFrame2.WordArtFormat"
    End With
    On Error GoTo -1
    
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

Private Sub SyncShapeProperties( _
                          ByRef shp_source As Shape, _
                          ByRef shp_target As Shape)
' ----------------------------------------------------
' Syncronize the target shape's properties with those
' from the source shape. Any non applicable property
' is just skipped, syncronized properties are logged.
' ----------------------------------------------------
    Const PROC = "SynchProperties"
    
    On Error GoTo eh
    Log.ServicedItem = shp_target
    
    With shp_target
        On Error Resume Next
        SyncProperty .AlternativeText, shp_source.AlternativeText, "AlternativeText"
        SyncProperty .AutoShapeType, shp_source.AutoShapeType, "AutoShapeType"
        SyncProperty .BackgroundStyle, shp_source.BackgroundStyle, "BackgroundStyle"
        SyncProperty .BlackWhiteMode, shp_source.BlackWhiteMode, "BlackWhiteMode"
        SyncProperty .Decorative, shp_source.Decorative, "Decorative"
        SyncProperty .GraphicStyle, shp_source.GraphicStyle, "GraphicStyle"
        SyncProperty .Height, shp_source.Height, "Height"
        SyncProperty .Left, shp_source.Left, "Left"
        SyncProperty .LockAspectRatio, shp_source.LockAspectRatio, "LockAspectRatio"
        SyncProperty .Locked, shp_source.Locked, "Locked"
        SyncProperty .OnAction, shp_source.OnAction, "OnAction"
        SyncProperty .Placement, shp_source.Placement, "Placement"
        SyncProperty .Rotation, shp_source.Rotation, "Rotation"
        SyncProperty .ShapeStyle, shp_source.ShapeStyle, "ShapeStyle"
        SyncProperty .Title, shp_source.Title, "Title"
        SyncProperty .top, shp_source.top, "Top"
        SyncProperty .Visible, shp_source.Visible, "Visible"
        SyncProperty .Width, shp_source.Width, "Width"
    End With
    
    On Error Resume Next ' Some shapes may not have some properties
    SyncShapePropertiesTextFrame shp_source.TextFrame, shp_target.TextFrame
    SyncShapePropertiesTextFrame2 shp_source.TextFrame2, shp_target.TextFrame2
    SyncShapePropertiesTextEffect shp_source.TextEffect, shp_target.TextEffect
    
xt: On Error GoTo -1
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncOOBProperties( _
                        ByRef sop_source As Variant, _
                        ByRef sop_target As Variant)
' ----------------------------------------------------
' Syncronize the target shape's properties with those
' from the source shape. Any non applicable property
' is just skipped, syncronized properties are logged.
' ----------------------------------------------------
    Const PROC = "SyncOOBProperties"
    
    On Error GoTo eh
    Dim oobSource   As OLEObject
    Dim oobTarget   As OLEObject
'    Dim tbt As ToggleButton
'    Dim tbx As TextBox
'    Dim sbt As SpinButton
'    Dim scb As ScrollBar
'    Dim obt As OptionButton
'    Dim lbx As ListBox
'    Dim lbl As Label
'    Dim img As Image
'    Dim cbt As CommandButton
    
    Set oobTarget = sop_target
    Set oobSource = sop_source
    Log.ServicedItem = oobTarget
        
    Debug.Print oobTarget.OLEType
    
    With oobTarget
        On Error Resume Next
        SyncProperty .AutoLoad, oobSource.AutoLoad, "OLEObject.AutoLoad"
        SyncProperty .AutoUpdate, oobSource.AutoUpdate, "OLEObject.AutoUpdate"
        SyncProperty .Border.Color, oobSource.Border.Color, "OLEObject.Border.Color"
        SyncProperty .Border.ColorIndex, oobSource.Border.ColorIndex, "OLEObject.Border.ColorIndex"
        SyncProperty .Border.LineStyle, oobSource.Border.LineStyle, "OLEObject.Border.LineStyle"
        SyncProperty .Border.ThemeColor, oobSource.Border.ThemeColor, "OLEObject.Border.ThemeColor"
        SyncProperty .Border.TintAndShade, oobSource.Border.TintAndShade, "OLEObject.Border.TintAndShade"
        SyncProperty .Enabled, oobSource.Enabled, "OLEObject.Enabled"
        SyncProperty .Height, oobSource.Height, "OLEObject.Height"
        SyncProperty .Left, oobSource.Left, "OLEObject.Left"
        SyncProperty .LinkedCell, oobSource.LinkedCell, "OLEObject.LinkedCell"
        SyncProperty .Name, oobSource.Name, "OLEObject.Name"
        SyncProperty .OLEType, oobSource.OLEType, "OLEObject.OLEType"
        SyncProperty .Placement, oobSource.Placement, "OLEObject.Placement"
        SyncProperty .PrintObject, oobSource.PrintObject, "OLEObject.PrintObject"
        SyncProperty .Shadow, oobSource.Shadow, "OLEObject.Shadow"
        SyncProperty .SourceName, oobSource.SourceName, "OLEObject.SourceName"
        SyncProperty .top, oobSource.top, "OLEObject.Top"
    End With
xt: On Error GoTo -1
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

