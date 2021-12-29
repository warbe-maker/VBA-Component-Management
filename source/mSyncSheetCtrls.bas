Attribute VB_Name = "mSyncSheetCtrls"
Option Explicit
' -------------------------------------------------------------------
' Standard Module mSyncSheetControls
'          Services to synchronize new, obsolete, and properties of
'          sheet shapes and OLEObjects.
'
' -------------------------------------------------------------------

Public Enum enCntrlPropertiesLet
    '~~ -----------------------------------------------------
    '~~ All properties applicable and 'writeable'
    '~~ for any kind of Form-Control (Shape, OLEObject, etc.)
    '~~ -----------------------------------------------------
    enAccelerator
    enAlignment
    enAltHTML
    enAlternativeText
    enAutoLoad
    enAutoShapeType
    enAutoSize
    enAutoUpdate
    enBackColor
    enBackStyle
    enBackgroundStyle
    enBlackWhiteMode
    enBorderColor
    enBorderColorIndex
    enBorderLineStyle
    enBorderParent
    enBorderThemeColor
    enBorderTintAndShade
    enBorderWeight
    enBordersSuppress
    enCaption
    enDecorative
    enEnabled
    enFontBold
    enFontColor
    enFontColorIndex
    enFontCreator
    enFontFontStyle
    enFontItalic
    enFontName
    enFontParent
    enFontSize
    enFontStrikethrough
    enFontSubscript
    enFontSuperscript
    enFontThemeColor
    enFontThemeFont
    enFontTintAndShade
    enFontUnderline
    enForeColor
    enGraphicStyle
    enGroupName
    enHeight
    enLeft
    enLinkedCell
    enListFillRange
    enLockAspectRatio
    enLocked
    enMouseIcon
    enMousePointer
    enMultiSelect
    enName
    enOnAction
    enPicture
    enPicturePosition
    enPlacement
    enPrintObject
    enRotation
    enShadow
    enShapeStyle
    enSourceName
    enSpecialEffect
    enTakeFocusOnClick
    enTextAlign
    enTextEffectFormatAlignment
    enTextEffectFormatFontBold
    enTextEffectFormatFontItalic
    enTextEffectFormatFontName
    enTextEffectFormatFontSize
    enTextEffectFormatKernedPairs
    enTextEffectFormatNormalizedHeight
    enTextEffectFormatPresetShape
    enTextEffectFormatPresetTextEffect
    enTextEffectFormatRotatedChars
    enTextEffectFormattext
    enTextEffectFormatTracking
    enTextFrame2AutoSize
    enTextFrame2HorizontalAnchor
    enTextFrame2MarginBottom
    enTextFrame2MarginLeft
    enTextFrame2MarginRight
    enTextFrame2MarginTop
    enTextFrame2NoTextRotation
    enTextFrame2Orientation
    enTextFrame2PathFormat
    enTextFrame2VerticalAnchor
    enTextFrame2WarpForma
    enTextFrame2WordArtformat
    enTextFrameAutoMargins
    enTextFrameAutoSize
    enTextFrameHorizontalAlignment
    enTextFrameHorizontalOverflow
    enTextFrameMarginBottom
    enTextFrameMarginLeft
    enTextFrameMarginRight
    enTextFrameMarginTop
    enTextFrameOrientation
    enTextFrameReadingOrder
    enTextFrameVerticalAlignment
    enTextFrameVerticalOverflow
    enTitle
    enTop
    enTripleState
    enVisible
    enWidth
    enWordWrap
    en_Font_Reserved
End Enum

Public Function CntrlPropertyName(ByVal en As enCntrlPropertiesLet) As String
' ------------------------------------------------------------------------
' Returns the Property-Name for the provided enumerated property (en).
' ------------------------------------------------------------------------
    
    Select Case en
        Case enAccelerator:                         CntrlPropertyName = "Accelerator"
        Case enAlignment:                           CntrlPropertyName = "Alignment"
        Case enAltHTML:                             CntrlPropertyName = "AltHTML"
        Case enAlternativeText:                     CntrlPropertyName = "AlternativeText"
        Case enAutoLoad:                            CntrlPropertyName = "AutoLoad"
        Case enAutoShapeType:                       CntrlPropertyName = "AutoShapeType"
        Case enAutoSize:                            CntrlPropertyName = "AutoSize"
        Case enAutoUpdate:                          CntrlPropertyName = "AutoUpdate"
        Case enBackColor:                           CntrlPropertyName = "BackColor"
        Case enBackgroundStyle:                     CntrlPropertyName = "BackgroundStyle"
        Case enBackStyle:                           CntrlPropertyName = "BackStyle"
        Case enBlackWhiteMode:                      CntrlPropertyName = "BlackWhitemode"
        Case enBorderColor:                         CntrlPropertyName = "Border.Color"
        Case enBorderColorIndex:                    CntrlPropertyName = "Border.ColorIndex"
        Case enBorderLineStyle:                     CntrlPropertyName = "Border.LineStyle"
        Case enBorderParent:                        CntrlPropertyName = "Border.Parent"
        Case enBorderThemeColor:                    CntrlPropertyName = "Border.ThemeColor"
        Case enBorderTintAndShade:                  CntrlPropertyName = "Border.TintAndShade"
        Case enBorderWeight:                        CntrlPropertyName = "Weight"
        Case enCaption:                             CntrlPropertyName = "Caption"
        Case enDecorative:                          CntrlPropertyName = "Decorative"
        Case enEnabled:                             CntrlPropertyName = "Enabled"
        Case enFontBold:                            CntrlPropertyName = "Font.Bold"
        Case enFontColor:                           CntrlPropertyName = "Font.Color"
        Case enFontColorIndex:                      CntrlPropertyName = "Font.ColorIndex"
        Case enFontCreator:                         CntrlPropertyName = "Font.Creator"
        Case enFontFontStyle:                       CntrlPropertyName = "Font.FontStyle"
        Case enFontItalic:                          CntrlPropertyName = "Font.Italic"
        Case enFontName:                            CntrlPropertyName = "Font.Name"
        Case enFontParent:                          CntrlPropertyName = "Font.Parent"
        Case enFontSize:                            CntrlPropertyName = "Font.Size"
        Case enFontStrikethrough:                   CntrlPropertyName = "Font.Strikethrough"
        Case enFontSubscript:                       CntrlPropertyName = "Font.Subscript"
        Case enFontSuperscript:                     CntrlPropertyName = "Font.Superscript"
        Case enFontThemeColor:                      CntrlPropertyName = "Font.ThemeColor"
        Case enFontThemeFont:                       CntrlPropertyName = "Font.ThemeFont"
        Case enFontTintAndShade:                    CntrlPropertyName = "Font.TintAndShade"
        Case enFontUnderline:                       CntrlPropertyName = "Font.Underline"
        Case enForeColor:                           CntrlPropertyName = "ForeColor"
        Case enGraphicStyle:                        CntrlPropertyName = "GraphicStyle"
        Case enHeight:                              CntrlPropertyName = "Height"
        Case enLeft:                                CntrlPropertyName = "Left"
        Case enLinkedCell:                          CntrlPropertyName = "LinkedCell"
        Case enListFillRange:                       CntrlPropertyName = "ListFillRange"
        Case enLockAspectRatio:                     CntrlPropertyName = "Ratio"
        Case enLocked:                              CntrlPropertyName = "Locked"
        Case enMouseIcon:                           CntrlPropertyName = "MouseIcon"
        Case enMousePointer:                        CntrlPropertyName = "MousPointer"
        Case enName:                                CntrlPropertyName = "Name"
        Case enPicture:                             CntrlPropertyName = "Picture"
        Case enPicturePosition:                     CntrlPropertyName = "PicturePosition"
        Case enPlacement:                           CntrlPropertyName = "Placement"
        Case enPrintObject:                         CntrlPropertyName = "PrintObject"
        Case enRotation:                            CntrlPropertyName = "Rotation"
        Case enShadow:                              CntrlPropertyName = "Shadow"
        Case enShapeStyle:                          CntrlPropertyName = "ShapeStyle"
        Case enTakeFocusOnClick:                    CntrlPropertyName = "TakeFocusOnClick"
        Case enTextEffectFormatAlignment:           CntrlPropertyName = "TextEffectFormat.Alignment"
        Case enTextEffectFormatFontBold:            CntrlPropertyName = "TextEffectFormat.FontBold"
        Case enTextEffectFormatFontItalic:          CntrlPropertyName = "TextEffectFormat.FontItalic"
        Case enTextEffectFormatFontName:            CntrlPropertyName = "TextEffectFormat.FontName"
        Case enTextEffectFormatFontSize:            CntrlPropertyName = "TextEffectFormat.FontSize"
        Case enTextEffectFormatKernedPairs:         CntrlPropertyName = "TextEffectFormat.KernedPairs"
        Case enTextEffectFormatNormalizedHeight:    CntrlPropertyName = "TextEffectFormat.NormalizedHeight"
        Case enTextEffectFormatPresetShape:         CntrlPropertyName = "TextEffectFormat.PresetShape"
        Case enTextEffectFormatPresetTextEffect:    CntrlPropertyName = "TextEffectFormat.PresetTextEffect"
        Case enTextEffectFormatRotatedChars:        CntrlPropertyName = "TextEffectFormat.RotatedChars"
        Case enTextEffectFormattext:                CntrlPropertyName = "TextEffectFormat.text"
        Case enTextEffectFormatTracking:            CntrlPropertyName = "TextEffectFormat.Tracking"
        Case enTextFrame2AutoSize:                  CntrlPropertyName = "TextFrame2.AutoSize"
        Case enTextFrame2HorizontalAnchor:          CntrlPropertyName = "TextFrame2.HorizontalAnchor"
        Case enTextFrame2MarginBottom:              CntrlPropertyName = "TextFrame2.MarginBottom"
        Case enTextFrame2MarginLeft:                CntrlPropertyName = "TextFrame2.MarginLeft"
        Case enTextFrame2MarginRight:               CntrlPropertyName = "TextFrame2.MarginRight"
        Case enTextFrame2MarginTop:                 CntrlPropertyName = "TextFrame2.MarginTop"
        Case enTextFrame2NoTextRotation:            CntrlPropertyName = "TextFrame2.NoTextRotation"
        Case enTextFrame2Orientation:               CntrlPropertyName = "TextFrame2.Orientation"
        Case enTextFrame2PathFormat:                CntrlPropertyName = "TextFrame2.PathFormat"
        Case enTextFrame2VerticalAnchor:            CntrlPropertyName = "TextFrame2.VerticalAnchor"
        Case enTextFrame2WarpForma:                 CntrlPropertyName = "TextFrame2.WarpFormat"
        Case enTextFrame2WordArtformat:             CntrlPropertyName = "TextFrame2.WordArtformat"
        Case enTextFrameAutoMargins:                CntrlPropertyName = "TextFrame.AutoMargins"
        Case enTextFrameAutoSize:                   CntrlPropertyName = "TextFrame.AutoSize"
        Case enTextFrameHorizontalAlignment:        CntrlPropertyName = "TextFrame.HorizontalAlignment"
        Case enTextFrameHorizontalOverflow:         CntrlPropertyName = "TextFrame.HorizontalOverflow"
        Case enTextFrameMarginBottom:               CntrlPropertyName = "TextFrame.MarginBottom"
        Case enTextFrameMarginLeft:                 CntrlPropertyName = "TextFrame.MarginLeft"
        Case enTextFrameMarginRight:                CntrlPropertyName = "TextFrame.MarginRight"
        Case enTextFrameMarginTop:                  CntrlPropertyName = "TextFrame.MarginTop"
        Case enTextFrameOrientation:                CntrlPropertyName = "TextFrame.Orientation"
        Case enTextFrameReadingOrder:               CntrlPropertyName = "TextFrame.ReadingOrder"
        Case enTextFrameVerticalAlignment:          CntrlPropertyName = "TextFrame.VerticalAlignment"
        Case enTextFrameVerticalOverflow:           CntrlPropertyName = "TextFrame.VerticalOverflow"
        Case enTitle:                               CntrlPropertyName = "Title"
        Case enTop:                                 CntrlPropertyName = "Top"
        Case enVisible:                             CntrlPropertyName = "Visible"
        Case enWidth:                               CntrlPropertyName = "Width"
        Case enWordWrap:                            CntrlPropertyName = "WordWrap"
        Case Else
                                              CntrlPropertyName = "? for " & en
    End Select
    
End Function

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
            .Top = SourceShape.Top
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
            .Top = SourceOOB.Top
            .Left = SourceOOB.Left
            .Width = SourceOOB.Width
            .Height = SourceOOB.Height
        End With
next_shape:
    Next SourceOOB
End Sub

Public Sub SyncOOBsNew()
' -------------------------------------------------------------
' Copy all new controls from the sourec Workbook (Sync.Source)
' to the target Workbook (Sync.Target) and ajust the properties
' -------------------------------------------------------------
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
            If mSheetControls.ControlExists(sync_wb:=.Target _
                                          , sync_sheet_name:=sSheet _
                                          , sync_sheet_code_name:=mSyncSheets.SheetCodeName(.Source, sSheet) _
                                          , sync_sheet_control_name:=sControl _
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
                     Stats.Count sic_sheet_controls_new
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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncControlsNew()
' -----------------------------------------------------------
' Copy new shapes from the sourec Workbook (Sync.Source) to
' the target Workbook (Sync.Target) and ajust the properties
' -----------------------------------------------------------
    Const PROC = "SyncControlsNew"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    Dim sControl    As String
    Dim sSheet      As String
    Dim shp         As Shape
    
    With Sync
        For Each v In .SourceSheetControls
            sSheet = mSync.KeySheetName(v)
            sControl = .SourceSheetControls(v)
            If mSheetControls.ControlExists(sync_wb:=.Target _
                                          , sync_sheet_name:=sSheet _
                                          , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                                          , sync_sheet_control_name:=sControl _
                                           ) _
            Then GoTo next_shape
            '~~ The source sheet control not yet exists in the target Workbook's corresponding sheet
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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
            If mSheetControls.ControlExists(sync_wb:=.Source _
                                          , sync_sheet_name:=KeySheetName(v) _
                                          , sync_sheet_code_name:=SheetCodeName(.Target, sSheet) _
                                          , sync_sheet_control_name:=sControl _
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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncControlsObsolete()
' -----------------------------------------------------------
' Remove obsolete shapes in the target Workbook (Sync.Target)
' -----------------------------------------------------------
    Const PROC = "SyncControlsObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsTarget    As Worksheet
    Dim sControl    As String
    Dim sSheet      As String
    Dim shp         As Shape
    
    With Sync
        For Each v In .TargetSheetControls
            Application.ScreenUpdating = False
            sSheet = mSync.KeySheetName(v)
            sControl = mSync.KeyControlName(v)
            If mSheetControls.ControlExists(sync_wb:=.Source _
                                          , sync_sheet_name:=KeySheetName(v) _
                                          , sync_sheet_code_name:=SheetCodeName(.Target, sSheet) _
                                          , sync_sheet_control_name:=sControl _
                                           ) _
            Then GoTo next_shape
            Set wsTarget = .Target.Worksheets(sSheet)
            Set shp = wsTarget.Shapes(sControl)
            Log.ServicedItem = shp
            
            Stats.Count sic_sheet_controls_obsolete
            '~~ The target name does not exist in the source and thus  has become obsolete
            If .Mode = Confirm Then
                .ConfInfo = "Obsolete!"
            Else
                wsTarget.Shapes(sControl).Delete
                Log.Entry = "Deteted!"
            End If
next_shape:
            Application.ScreenUpdating = True
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncControlsProperties()
' -----------------------------------------
' Syncronize all shape's properties between
' source and target Workbook.
' -----------------------------------------
    Const PROC = "SyncControlsProperties"
        
    On Error GoTo eh
    Dim v           As Variant
    Dim sControl    As String
    Dim sSheet      As String
    
    With Sync
        For Each v In .SourceSheetControls
            sSheet = mSync.KeySheetName(v)
            sControl = mSync.KeyControlName(v)
            If Not mSheetControls.ControlExists(sync_wb:=.Target _
                                              , sync_sheet_name:=sSheet _
                                              , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                                              , sync_sheet_control_name:=sControl _
                                               ) _
            Then GoTo next_shape
            Log.ServicedItem = sControl

            SyncCntrlProperties .Source.Worksheets(sSheet).Shapes.Item(sControl) _
                              , .Target.Worksheets(sSheet).Shapes.Item(sControl)

next_shape:
        Next v
    End With

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
            If Not mSheetControls.ControlExists(sync_wb:=.Target _
                                              , sync_sheet_name:=sSheet _
                                              , sync_sheet_code_name:=SheetCodeName(.Source, sSheet) _
                                              , sync_sheet_control_name:=sControl _
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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SyncControlPropertiesTextEffect( _
                                      ByRef te_source As TextEffectFormat, _
                                      ByRef te_target As TextEffectFormat)
' --------------------------------------------------------------------------
'
' --------------------------------------------------------------------------
    Dim en As enCntrlPropertiesLet
    
    On Error Resume Next
    With te_target
        en = enTextEffectFormatAlignment:           SyncProperty en, .Alignment, te_target.Alignment
        en = enTextEffectFormatFontBold:            SyncProperty en, .FontBold, te_source.FontBold
        en = enTextEffectFormatFontItalic:          SyncProperty en, .FontItalic, te_source.FontItalic
        en = enTextEffectFormatFontName:            SyncProperty en, .FontName, te_source.FontName
        en = enTextEffectFormatFontSize:            SyncProperty en, .FontSize, te_source.FontSize
        en = enTextEffectFormatKernedPairs:         SyncProperty en, .KernedPairs, te_source.KernedPairs
        en = enTextEffectFormatNormalizedHeight:    SyncProperty en, .NormalizedHeight, te_source.NormalizedHeight
        en = enTextEffectFormatPresetShape:         SyncProperty en, .PresetShape, te_source.PresetShape
        en = enTextEffectFormatPresetTextEffect:    SyncProperty en, .PresetTextEffect, te_source.PresetTextEffect
        en = enTextEffectFormatRotatedChars:        SyncProperty en, .RotatedChars, te_source.RotatedChars
        en = enTextEffectFormattext:                SyncProperty en, .Text, te_source.Text
        en = enTextEffectFormatTracking:            SyncProperty en, .Tracking, te_source.Tracking
    End With
    On Error GoTo -1
    
End Sub

Private Sub SyncControlPropertiesTextFrame( _
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

Private Sub SyncControlPropertiesTextFrame2( _
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
                   ByVal en As enCntrlPropertiesLet, _
                   ByRef v_target As Variant, _
                   ByRef v_source As Variant)
' ----------------------------------------------------
' Synchronizes a Shape or OLEObject property.
' A not applicable property is skipped.
' ----------------------------------------------------
    On Error GoTo xt
    If v_target <> v_source Then
        On Error Resume Next ' The property may not be modifyable
        v_target = v_source
        Select Case Err.Number
            Case 0:     Log.Entry = "Property synced: " & CntrlPropertyName(en) & "(from '" & CStr(v_target) & "' to '" & CStr(v_source) & "')"
            Case Else
                Log.Entry = "Property synchronization failed (error " & Err.Number & ") '" & CntrlPropertyName(en) & "'"
        End Select
    End If

xt:
End Sub

Private Sub SyncCntrlProperties( _
                          ByRef cntrl_source As Variant, _
                          ByRef cntrl_target As Variant)
' --------------------------------------------------------
' Syncronize the target shape's properties with those
' from the source shape. Any non applicable property
' is just skipped, syncronized properties are logged.
' --------------------------------------------------------
    Const PROC = "SynchProperties"
    
    On Error GoTo eh
    Dim en  As enCntrlPropertiesLet
    
    Log.ServicedItem = cntrl_target
    
    With cntrl_target
        en = enAccelerator:         SyncProperty en, .Accelerator, cntrl_source.Accelerator
        en = enAlignment:           SyncProperty en, .Alignment, cntrl_source.Alignment
        en = enAltHTML:             SyncProperty en, .AltHTML, cntrl_source.AltHTML
        en = enAlternativeText:     SyncProperty en, .AlternativeText, cntrl_source.AlternativeText
        en = enAutoLoad:            SyncProperty en, .AutoLoad, cntrl_source.AutoLoad
        en = enAutoShapeType:       SyncProperty en, .AutoShapeType, cntrl_source.AutoShapeType
        en = enAutoSize:            SyncProperty en, .AutoSize, cntrl_source.AutoSize
        en = enAutoUpdate:          SyncProperty en, .AutoUpdate, cntrl_source.AutoUpdate
        en = enBackColor:           SyncProperty en, .BackColor, cntrl_source.BackColor
        en = enBackStyle:           SyncProperty en, .BackStyle, cntrl_source.BackStyle
        en = enBackgroundStyle:     SyncProperty en, .BackgroundStyle, cntrl_source.BackgroundStyle
        en = enBlackWhiteMode:      SyncProperty en, .BlackWhiteMode, cntrl_source.BlackWhiteMode
        en = enBorderColor:         SyncProperty en, .BorderColor, cntrl_source.BorderColor
        en = enBorderColorIndex:    SyncProperty en, .BorderColorIndex, cntrl_source.BorderColorIndex
        en = enBorderLineStyle:     SyncProperty en, .BorderLineStyle, cntrl_source.BorderLineStyle
        en = enBorderParent:        SyncProperty en, .BorderParent, cntrl_source.BorderParent
        en = enBorderThemeColor:    SyncProperty en, .BorderThemeColor, cntrl_source.BorderThemeColor
        en = enBorderTintAndShade:  SyncProperty en, .BorderTintAndShade, cntrl_source.BorderTintAndShade
        en = enBorderWeight:        SyncProperty en, .BorderWeight, cntrl_source.BorderWeight
        en = enBordersSuppress:     SyncProperty en, .BordersSuppress, cntrl_source.BordersSuppress
        en = enCaption:             SyncProperty en, .Caption, cntrl_source.Caption
        en = enDecorative:          SyncProperty en, .Decorative, cntrl_source.Decorative
        en = enEnabled:             SyncProperty en, .Enabled, cntrl_source.Enabled
        en = enFontBold:            SyncProperty en, .FontBold, cntrl_source.FontBold
        en = enFontColor:           SyncProperty en, .FontColor, cntrl_source.FontColor
        en = enFontColorIndex:      SyncProperty en, .FontColorIndex, cntrl_source.FontColorIndex
        en = enFontCreator:         SyncProperty en, .FontCreator, cntrl_source.FontCreator
        en = enFontFontStyle:       SyncProperty en, .FontStyle, cntrl_source.FontStyle
        en = enFontItalic:          SyncProperty en, .FontItalic, cntrl_source.FontItalic
        en = enFontName:            SyncProperty en, .FontName, cntrl_source.FontName
        en = enFontParent:          SyncProperty en, .FontParent, cntrl_source.FontParent
        en = enFontSize:            SyncProperty en, .FontSize, cntrl_source.FontSize
        en = enFontStrikethrough:   SyncProperty en, .FONTSTRIKETHROUGH, cntrl_source.FONTSTRIKETHROUGH
        en = enFontSubscript:       SyncProperty en, .FontSubscript, cntrl_source.FontSubscript
        en = enFontSuperscript:     SyncProperty en, .FontSuperscript, cntrl_source.FontSuperscript
        en = enFontThemeColor:      SyncProperty en, .FontThemeColor, cntrl_source.FontThemeColor
        en = enFontThemeFont:       SyncProperty en, .FontThemeFont, cntrl_source.FontThemeFont
        en = enFontTintAndShade:    SyncProperty en, .FontFontTintAndShade, cntrl_source.FontFontTintAndShade
        en = enFontUnderline:       SyncProperty en, .FontUnderline, cntrl_source.FontUnderline
        en = enForeColor:           SyncProperty en, .ForeColor, cntrl_source.ForeColor
        en = enGraphicStyle:        SyncProperty en, .GraphicStyle, cntrl_source.GraphicStyle
        en = enGroupName:           SyncProperty en, .GroupName, cntrl_source.GroupName
        en = enHeight:              SyncProperty en, .Height, cntrl_source.Height
        en = enLeft:                SyncProperty en, .Left, cntrl_source.Left
        en = enLinkedCell:          SyncProperty en, .LinkedCell, cntrl_source.LinkedCell
        en = enListFillRange:       SyncProperty en, .ListFillRange, cntrl_source.ListFillRange
        en = enLockAspectRatio:     SyncProperty en, .LockAspectRatio, cntrl_source.LockAspectRatio
        en = enLocked:              SyncProperty en, .Locked, cntrl_source.Locked
        en = enMouseIcon:           SyncProperty en, .MouseIcon, cntrl_source.MouseIcon
        en = enMousePointer:        SyncProperty en, .MousePointer, cntrl_source.MousePointer
        en = enMultiSelect:         SyncProperty en, .MultiSelect, cntrl_source.MultiSelect
        en = enName:                SyncProperty en, .Name, cntrl_source.Name
        en = enOnAction:            SyncProperty en, .OnAction, cntrl_source.OnAction
        en = enPicture:             SyncProperty en, .Picture, cntrl_source.Picture
        en = enPicturePosition:     SyncProperty en, .PicturePosition, cntrl_source.PicturePosition
        en = enPlacement:           SyncProperty en, .Placement, cntrl_source.Placement
        en = enPrintObject:         SyncProperty en, .PrintObject, cntrl_source.PrintObject
        en = enRotation:            SyncProperty en, .Rotation, cntrl_source.Rotation
        en = enShadow:              SyncProperty en, .Shadow, cntrl_source.Shadow
        en = enShapeStyle:          SyncProperty en, .ShapeStyle, cntrl_source.ShapeStyle
        en = enSourceName:          SyncProperty en, .SourceName, cntrl_source.SourceName
        en = enSpecialEffect:       SyncProperty en, .SpecialEffect, cntrl_source.SpecialEffect
        en = enTakeFocusOnClick:    SyncProperty en, .TakeFocusOnClick, cntrl_source.TakeFocusOnClick
        en = enTextAlign:           SyncProperty en, .TextAlign, cntrl_source.TextAlign
        en = enTitle:               SyncProperty en, .Title, cntrl_source.Title
        en = enTop:                 SyncProperty en, .Top, cntrl_source.Top
        en = enVisible:             SyncProperty en, .Visible, cntrl_source.Visible
        en = enWidth:               SyncProperty en, .Width, cntrl_source.Width
        en = enWordWrap:            SyncProperty en, .WordWrap, cntrl_source.WordWrap
    End With
    
    On Error Resume Next ' Some shapes may not have some properties
    SyncControlPropertiesTextFrame cntrl_source.TextFrame, cntrl_target.TextFrame
    SyncControlPropertiesTextFrame2 cntrl_source.TextFrame2, cntrl_target.TextFrame2
    SyncControlPropertiesTextEffect cntrl_source.TextEffect, cntrl_target.TextEffect
    
xt: On Error GoTo -1
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
        SyncProperty .Top, oobSource.Top, "OLEObject.Top"
    End With

xt: On Error GoTo -1
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

