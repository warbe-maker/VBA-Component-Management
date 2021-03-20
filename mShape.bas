Attribute VB_Name = "mShape"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mShape." & s
End Function

Public Sub SyncProperties(ByRef shp_source As Shape, _
                           ByRef shp_target As Shape)
' ----------------------------------------------------
' Syncronize the target shape's properties with those
' from the source shape. Any non applicable property
' is just skipped, syncronized properties are logged.
' ----------------------------------------------------
    Const PROC = "SynchProperties"
    
    On Error GoTo eh
    cLog.ServicedItem = shp_target
    
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
    SyncPropertiesTextFrame shp_source.TextFrame, shp_target.TextFrame
    SyncPropertiesTextFrame2 shp_source.TextFrame2, shp_target.TextFrame2
    SyncPropertiesTextEffect shp_source.TextEffect, shp_target.TextEffect
    
xt: On Error GoTo -1
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncPropertiesTextFrame( _
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

Private Sub SyncPropertiesTextFrame2( _
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

Private Sub SyncPropertiesTextEffect( _
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

Private Sub SyncProperty( _
                   ByRef v_target As Variant, _
                   ByRef v_source As Variant, _
                   ByVal s_property As String)
' ---------------------------------------------
' Synchronizes a single Shape property.
' ---------------------------------------------
    On Error GoTo xt
    If v_target <> v_source Then
        On Error Resume Next ' The property may not be modifyable
        v_target = v_source
        If Err.Number = 0 _
        Then cLog.Entry = "Property '" & s_property & "' synched" _
        Else cLog.Entry = "Difference in Property '" & s_property & "' could not be synched (error " & Err.Number & ")"
    End If
xt:
End Sub
