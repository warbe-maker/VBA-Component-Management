Attribute VB_Name = "mShape"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mShape." & s
End Function

Public Sub SynchShapeProperties(ByRef source As Shape, _
                                ByRef target As Shape)
' ------------------------------------------------------
' Syncronize the target shape's properties with those
' from the source shape. Any non applicable property
' is just skipped, syncronized properties are logged.
' ------------------------------------------------------
    Const PROC = "SynchShapeProperties"
    
    cLog.ServicedItem(TypeName(target)) = target.Name
    
    With target
        
        On Error GoTo p01
        If .AlternativeText <> source.AlternativeText Then
            On Error Resume Next
            .AlternativeText = source.AlternativeText
            If Err.Number = 0 Then
                cLog.Entry = "Property 'AlternativeText' synched"
            Else
                cLog.Entry = "Difference in Property 'AlternativeText' could not be synched (error " & Err.Number & ")"
            End If
        End If

p01:    On Error GoTo p02
        If .AutoShapeType <> source.AutoShapeType Then
            On Error Resume Next
            .AutoShapeType = source.AutoShapeType
            If Err.Number = 0 Then
                cLog.Entry = "Property 'AutoShapeType' synched"
            Else
                cLog.Entry = "Difference in Property 'AutoShapeType' could not be synched (error " & Err.Number & ")"
            End If
        End If

p02:    On Error GoTo p03
        If .BackgroundStyle <> source.BackgroundStyle Then
            On Error Resume Next
            .BackgroundStyle = source.BackgroundStyle
            If Err.Number = 0 Then
                cLog.Entry = "Property 'BackgroundStyle' synched"
            Else
                cLog.Entry = "Difference in Property 'BackgroundStyle' could not be synched (error " & Err.Number & ")"
            End If
        End If

p03:    On Error GoTo p04
        If .BlackWhiteMode <> source.BlackWhiteMode Then
            On Error Resume Next
            .BlackWhiteMode = source.BlackWhiteMode
            If Err.Number = 0 Then
                cLog.Entry = "Property 'BlackWhiteMode' synched"
            Else
                cLog.Entry = "Difference in Property 'BlackWhiteMode' could not be synched (error " & Err.Number & ")"
            End If
        End If

p04:    On Error GoTo p05
        If .Decorative <> source.Decorative Then
            On Error Resume Next
            .Decorative = source.Decorative
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Decorative' synched"
            Else
                cLog.Entry = "Difference in Property 'Decorative' could not be synched (error " & Err.Number & ")"
            End If
        End If

p05:    On Error GoTo p06
        If .GraphicStyle <> source.GraphicStyle Then
            On Error Resume Next
            .GraphicStyle = source.GraphicStyle
            If Err.Number = 0 Then
                cLog.Entry = "Property 'GraphicStyle' synched"
            Else
                cLog.Entry = "Difference in Property 'GraphicStyle' could not be synched (error " & Err.Number & ")"
            End If
        End If

p06:    On Error GoTo p07
        If .Height <> source.Height Then
            On Error Resume Next
            .Height = source.Height
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Height' synched"
            Else
                cLog.Entry = "Difference in Property 'Height' could not be synched (error " & Err.Number & ")"
            End If
        End If

p07:    On Error GoTo p08
        If .Left <> source.Left Then
            On Error Resume Next
            .Left = source.Left
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Left' synched"
            Else
                cLog.Entry = "Difference in Property 'Left' could not be synched (error " & Err.Number & ")"
            End If
        End If

p08:    On Error GoTo p09
        If .LockAspectRatio <> source.LockAspectRatio Then
            On Error Resume Next
            .LockAspectRatio = source.LockAspectRatio
            If Err.Number = 0 Then
                cLog.Entry = "Property 'LockAspectRatio' synched"
            Else
                cLog.Entry = "Difference in Property 'LockAspectRatio' could not be synched (error " & Err.Number & ")"
            End If
        End If

p09:    On Error GoTo p10
        If .Locked <> source.Locked Then
            On Error Resume Next
            .Locked = source.Locked
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Locked' synched"
            Else
                cLog.Entry = "Difference in Property 'Locked' could not be synched (error " & Err.Number & ")"
            End If
        End If

p10:    On Error GoTo p11
        If .OnAction <> source.OnAction Then
            On Error Resume Next
            .OnAction = source.OnAction
            If Err.Number = 0 Then
                cLog.Entry = "Property 'OnAction' synched"
            Else
                cLog.Entry = "Difference in Property 'OnAction' could not be synched (error " & Err.Number & ")"
            End If
        End If

p11:    On Error GoTo p12
        If .Placement <> source.Placement Then
            On Error Resume Next
            .Placement = source.Placement
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Placement' synched"
            Else
                cLog.Entry = "Difference in Property 'Placement' could not be synched (error " & Err.Number & ")"
            End If
        End If

p12:    On Error GoTo p13
        If .Rotation <> source.Rotation Then
            On Error Resume Next
            .Rotation = source.Rotation
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Rotation' synched"
            Else
                cLog.Entry = "Difference in Property 'Rotation' could not be synched (error " & Err.Number & ")"
            End If
        End If

p13:    On Error GoTo p14
        If .ShapeStyle <> source.ShapeStyle Then
            On Error Resume Next
            .ShapeStyle = source.ShapeStyle
            If Err.Number = 0 Then
                cLog.Entry = "Property 'ShapeStyle' synched"
            Else
                cLog.Entry = "Difference in Property 'ShapeStyle' could not be synched (error " & Err.Number & ")"
            End If
        End If

p14:    On Error GoTo p15
        If .Title <> source.Title Then
            On Error Resume Next
            .Title = source.Title
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Title' synched"
            Else
                cLog.Entry = "Difference in Property 'Title' could not be synched (error " & Err.Number & ")"
            End If
        End If

p15:    On Error GoTo p16
        If .top <> source.top Then
            On Error Resume Next
            .top = source.top
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Top' synched"
            Else
                cLog.Entry = "Difference in Property 'Top' could not be synched (error " & Err.Number & ")"
            End If
        End If

p16:    On Error GoTo p17
        If .Visible <> source.Visible Then
            On Error Resume Next
            .Visible = source.Visible
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Visible' synched"
            Else
                cLog.Entry = "Difference in Property 'Visible' could not be synched (error " & Err.Number & ")"
            End If
        End If

p17:    On Error GoTo p18
        If .Width <> source.Width Then
            On Error Resume Next
            .Width = source.Width
            If Err.Number = 0 Then
                cLog.Entry = "Property 'Width' synched"
            Else
                cLog.Entry = "Difference in Property 'Width' could not be synched (error " & Err.Number & ")"
            End If
        End If

p18:
        SyncShapePropertiesTextFrame source.TextFrame, target.TextFrame
        SyncShapePropertiesTextFrame2 source.TextFrame, target.TextFrame
        SyncShapePropertiesTextEffectFormat source.TextEffect, target.TextEffect
    
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SyncShapePropertiesTextFrame( _
                                   ByRef source As TextFrame, _
                                   ByRef target As TextFrame)
' -------------------------------------------------------------
'
' -------------------------------------------------------------
    
    On Error GoTo xt
    
    With target
        
        On Error GoTo p01
        If .AutoMargins <> source.AutoMargins Then
            On Error Resume Next
            .AutoMargins = source.AutoMargins
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.AutoMargins' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.AutoMargins' could not be synched (error " & Err.Number & ")"
            End If
        End If
        
p01:    On Error GoTo p02
        If .AutoSize <> source.AutoSize Then
            On Error Resume Next
            .AutoSize = source.AutoSize
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.AutoSize' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.AutoSize' could not be synched (error " & Err.Number & ")"
            End If
        End If

p02:    On Error GoTo p03
        If .HorizontalAlignment <> source.HorizontalAlignment Then
            On Error Resume Next
            .HorizontalAlignment = source.HorizontalAlignment
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.HorizontalAlignment' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.HorizontalAlignment' could not be synched (error " & Err.Number & ")"
            End If
        End If

p03:    On Error GoTo p04
        If .HorizontalOverflow <> source.HorizontalOverflow Then
            On Error Resume Next
            .HorizontalOverflow = source.HorizontalOverflow
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.HorizontalOverflow' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.HorizontalOverflow' could not be synched (error " & Err.Number & ")"
            End If
        End If

p04:    On Error GoTo p05
        If .MarginBottom <> source.MarginBottom Then
            On Error Resume Next
            .MarginBottom = source.MarginBottom
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.MarginBottom' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.MarginBottom' could not be synched (error " & Err.Number & ")"
            End If
        End If

p05:    On Error GoTo p06
        If .MarginLeft <> source.MarginLeft Then
            On Error Resume Next
            .MarginLeft = source.MarginLeft
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.MarginLeft' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.MarginLeft' could not be synched (error " & Err.Number & ")"
            End If
        End If

p06:    On Error GoTo p07
        If .MarginRight <> source.MarginRight Then
            On Error Resume Next
            .MarginRight = source.MarginRight
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.MarginRight' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.MarginRight' could not be synched (error " & Err.Number & ")"
            End If
        End If

p07:    On Error GoTo p08
        If .MarginTop <> source.MarginTop Then
            On Error Resume Next
            .MarginTop = source.MarginTop
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.MarginTop' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.MarginTop' could not be synched (error " & Err.Number & ")"
            End If
        End If

p08:    On Error GoTo p09
        If .Orientation <> source.Orientation Then
            On Error Resume Next
            .Orientation = source.Orientation
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.Orientation' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.Orientation' could not be synched (error " & Err.Number & ")"
            End If
        End If

p09:    On Error GoTo p10
        If .ReadingOrder <> source.ReadingOrder Then
            On Error Resume Next
            .ReadingOrder = source.ReadingOrder
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.ReadingOrder' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.ReadingOrder' could not be synched (error " & Err.Number & ")"
            End If
        End If

p10:    On Error GoTo p11
        If .VerticalAlignment <> source.VerticalAlignment Then
            On Error Resume Next
            .VerticalAlignment = source.VerticalAlignment
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.VerticalAlignment' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.VerticalAlignment' could not be synched (error " & Err.Number & ")"
            End If
        End If

p11:    On Error GoTo p12
        If .VerticalOverflow <> source.VerticalOverflow Then
            On Error Resume Next
            .VerticalOverflow = source.VerticalOverflow
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame.VerticalOverflow' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame.VerticalOverflow' could not be synched (error " & Err.Number & ")"
            End If
        End If

p12:
    End With

xt: Exit Sub

End Sub


Private Sub SyncShapePropertiesTextFrame2( _
                                    ByRef source As TextFrame2, _
                                    ByRef target As TextFrame2)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    
    On Error GoTo xt
    
    With target
        
        On Error GoTo p01
        If .AutoSize <> source.AutoSize Then
            On Error Resume Next
            .AutoSize = source.AutoSize
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.AutoSize' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.AutoSize' could not be synched (error " & Err.Number & ")"
            End If
        End If
        
p01:    On Error GoTo p02
        If .HorizontalAnchor <> source.HorizontalAnchor Then
            On Error Resume Next
            .HorizontalAnchor = source.HorizontalAnchor
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.HorizontalAnchor' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.HorizontalAnchor' could not be synched (error " & Err.Number & ")"
            End If
        End If

p02:    On Error GoTo p03
        If .MarginBottom <> source.MarginBottom Then
            On Error Resume Next
            .MarginBottom = source.MarginBottom
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.MarginBottom' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.MarginBottom' could not be synched (error " & Err.Number & ")"
            End If
        End If

p03:    On Error GoTo p04
        If .MarginLeft <> source.MarginLeft Then
            On Error Resume Next
            .MarginLeft = source.MarginLeft
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.MarginLeft' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.MarginLeft' could not be synched (error " & Err.Number & ")"
            End If
        End If

p04:    On Error GoTo p05
        If .MarginRight <> source.MarginRight Then
            On Error Resume Next
            .MarginRight = source.MarginRight
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.MarginRight' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.MarginRight' could not be synched (error " & Err.Number & ")"
            End If
        End If

p05:    On Error GoTo p06
        If .MarginTop <> source.MarginTop Then
            On Error Resume Next
            .MarginTop = source.MarginTop
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.MarginTop' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.MarginTop' could not be synched (error " & Err.Number & ")"
            End If
        End If

p06:    On Error GoTo p07
        If .NoTextRotation <> source.NoTextRotation Then
            On Error Resume Next
            .NoTextRotation = source.NoTextRotation
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.NoTextRotation' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.NoTextRotation' could not be synched (error " & Err.Number & ")"
            End If
        End If

p07:    On Error GoTo p08
        If .Orientation <> source.Orientation Then
            On Error Resume Next
            .Orientation = source.Orientation
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.Orientation' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.Orientation' could not be synched (error " & Err.Number & ")"
            End If
        End If

p08:    On Error GoTo p09
        If .PathFormat <> source.PathFormat Then
            On Error Resume Next
            .PathFormat = source.PathFormat
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.PathFormat' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.PathFormat' could not be synched (error " & Err.Number & ")"
            End If
        End If

p09:    On Error GoTo p10
        If .VerticalAnchor <> source.VerticalAnchor Then
            On Error Resume Next
            .VerticalAnchor = source.VerticalAnchor
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.VerticalAnchor' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.VerticalAnchor' could not be synched (error " & Err.Number & ")"
            End If
        End If

p10:    On Error GoTo p11
        If .WarpFormat <> source.WarpFormat Then
            On Error Resume Next
            .WarpFormat = source.WarpFormat
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.WarpFormat' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.WarpFormat' could not be synched (error " & Err.Number & ")"
            End If
        End If

p11:    On Error GoTo p12
        If .WordArtformat <> source.WordArtformat Then
            On Error Resume Next
            .WordArtformat = source.WordArtformat
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.WordArtFormat' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.WordArtFormat' could not be synched (error " & Err.Number & ")"
            End If
        End If

p12:    On Error GoTo p13
        If .WordArtformat <> source.WordArtformat Then
            On Error Resume Next
            .WordArtformat = source.WordArtformat
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextFrame2.WordWrap' synched"
            Else
                cLog.Entry = "Difference in Property 'TextFrame2.WordWrap' could not be synched (error " & Err.Number & ")"
            End If
        End If
p13:
    End With

xt: Exit Sub

End Sub

Private Sub SyncShapePropertiesTextEffectFormat( _
                                          ByRef source As TextEffectFormat, _
                                          ByRef target As TextEffectFormat)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------

    With target
    
        On Error GoTo p01
        If .Alignment <> source.Alignment Then
            On Error Resume Next
            .Alignment = source.Alignment
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.Alignment' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.Alignment' could not be synched (error " & Err.Number & ")"
            End If
        End If
    
p01:    On Error GoTo p02
        If .FontBold <> source.FontBold Then
            On Error Resume Next
            .FontBold = source.FontBold
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.FontBold' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.FontBold' could not be synched (error " & Err.Number & ")"
            End If
        End If

p02:    On Error GoTo p03
        If .FontItalic <> source.FontItalic Then
            On Error Resume Next
            .FontItalic = source.FontItalic
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.FontItalic' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.FontItalic' could not be synched (error " & Err.Number & ")"
            End If
        End If

p03:    On Error GoTo p04
        If .FontName <> source.FontName Then
            On Error Resume Next
            .FontName = source.FontName
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.FontName' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.FontName' could not be synched (error " & Err.Number & ")"
            End If
        End If

p04:    On Error GoTo p05
        If .FontSize <> source.FontSize Then
            On Error Resume Next
            .FontSize = source.FontSize
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.FontSize' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.FontSize' could not be synched (error " & Err.Number & ")"
            End If
        End If

p05:    On Error GoTo p06
        If .KernedPairs <> source.KernedPairs Then
            On Error Resume Next
            .KernedPairs = source.KernedPairs
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.KernedPairs' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.KernedPairs' could not be synched (error " & Err.Number & ")"
            End If
        End If

p06:    On Error GoTo p07
        If .NormalizedHeight <> source.NormalizedHeight Then
            On Error Resume Next
            .NormalizedHeight = source.NormalizedHeight
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.NormalizedHeight' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.NormalizedHeight' could not be synched (error " & Err.Number & ")"
            End If
        End If

p07:    On Error GoTo p08
        If .PresetShape <> source.PresetShape Then
            On Error Resume Next
            .PresetShape = source.PresetShape
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.PresetShape' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.PresetShape' could not be synched (error " & Err.Number & ")"
            End If
        End If

p08:    On Error GoTo p09
        If .PresetTextEffect <> source.PresetTextEffect Then
            On Error Resume Next
            .PresetTextEffect = source.PresetTextEffect
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.PresetTextEffect' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.PresetTextEffect' could not be synched (error " & Err.Number & ")"
            End If
        End If

p09:    On Error GoTo p10
        If .RotatedChars <> source.RotatedChars Then
            On Error Resume Next
            .RotatedChars = source.RotatedChars
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.RotatedChars' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.RotatedChars' could not be synched (error " & Err.Number & ")"
            End If
        End If

p10:    On Error GoTo p11
        If .text <> source.text Then
            On Error Resume Next
            .text = source.text
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.Text' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.Text' could not be synched (error " & Err.Number & ")"
            End If
        End If

p11:    On Error GoTo p12
        If .Tracking <> source.Tracking Then
            On Error Resume Next
            .Tracking = source.Tracking
            If Err.Number = 0 Then
                cLog.Entry = "Property 'TextEffectFormat.Tracking' synched"
            Else
                cLog.Entry = "Difference in Property 'TextEffectFormat.Tracking' could not be synched (error " & Err.Number & ")"
            End If
        End If

p12:
    End With

End Sub

