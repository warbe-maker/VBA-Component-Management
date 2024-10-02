Attribute VB_Name = "mSyncShapePrprtys"
Option Explicit
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
Private lPropertyMaxLen  As Long

Private Enum enProperties
    ' -----------------------------------------------------------------
    ' Enumeration of all read/write properties of Shapes and OLEObjects
    ' -----------------------------------------------------------------
    enOOBAAAAAAAA
    enOOBAutoLoad
    enOOBBorderColor
    enOOBBorderColorIndex
    enOOBBorderLineStyle
    enOOBHeight
    enOOBLeft
    enOOBLinkedCell
    enOOBName
    enOOBObjectBackColor
    enOOBObjectBackStyle
    enOOBObjectCaption
    enOOBObjectEnabled
    enOOBObjectFontBold
    enOOBObjectFontItalic
    enOOBObjectFontName
    enOOBObjectFontSize
    enOOBObjectForeColor
    enOOBObjectLocked
    enOOBObjectMouseIcon
    enOOBObjectMousePointer
    enOOBObjectPicture
    enOOBObjectPictureposition
    enOOBObjectTakeFocusOnClick
    enOOBObjectWordWrap
    enOOBOLEType
    enOOBPlacement
    enOOBPrintObject
    enOOBShadow
    enOOBShapeRangeAlternativeText
    enOOBShapeRangeAutoShapeType
    enOOBShapeRangeBackgroundStyle
    enOOBShapeRangeBlackWhiteMode
    enOOBShapeRangeDecorative
    enOOBShapeRangeGraphicStyle
    enOOBShapeRangeHeight
    enOOBShapeRangeLeft
    enOOBShapeRangeLineBackColor
    enOOBShapeRangeLineBeginArrowheadLength
    enOOBShapeRangeLineBeginArrowheadStyle
    enOOBShapeRangeLineBeginArrowheadWidth
    enOOBShapeRangeLineDashStyle
    enOOBShapeRangeLineEndArrowheadLength
    enOOBShapeRangeLineEndArrowheadStyle
    enOOBShapeRangeLineEndArrowheadWidth
    enOOBShapeRangeLineForeColorBrightness
    enOOBShapeRangeLineForeColorObjectThemeColor
    enOOBShapeRangeLineForeColorRGB
    enOOBShapeRangeLineForeColorSchemeColor
    enOOBShapeRangeLineForeColorTintAndShade
    enOOBShapeRangeLineForeColorType
    enOOBShapeRangeLineInsetPen
    enOOBShapeRangeLinePattern
    enOOBShapeRangeLineStyle
    enOOBShapeRangeLineTransparency
    enOOBShapeRangeLineVisible
    enOOBShapeRangeLineWeight
    enOOBShapeRangeLockAspectRatio
    enOOBShapeRangeName
    enOOBShapeRangePictureFormatBrightness
    enOOBShapeRangePictureFormatColorType
    enOOBShapeRangePictureFormatContrast
    enOOBShapeRangePictureFormatTransparencyColor
    enOOBShapeRangePictureFormatTransparentBackground
    enOOBShapeRangeRotation
    enOOBShapeRangeShadowBlur
    enOOBShapeRangeShadowForeColorBrightness
    enOOBShapeRangeShadowForeColorObjectThemeColor
    enOOBShapeRangeShadowForeColorRGB
    enOOBShapeRangeShadowForeColorSchemeColor
    enOOBShapeRangeShadowForeColorTintAndShade
    enOOBShapeRangeShadowObscured
    enOOBShapeRangeShadowOffsetX
    enOOBShapeRangeShadowOffsetY
    enOOBShapeRangeShadowRotateWithShape
    enOOBShapeRangeShadowSize
    enOOBShapeRangeShadowStyle
    enOOBShapeRangeShadowTransparency
    enOOBShapeRangeShadowType
    enOOBShapeRangeShadowVisible
    enOOBShapeRangeShapeStyle
    enOOBShapeRangeSoftEdgeRadius
    enOOBShapeRangeSoftEdgeType
    enOOBShapeRangeTextEffectFontBold
    enOOBShapeRangeTextEffectFontItalic
    enOOBShapeRangeTextEffectFontName
    enOOBShapeRangeTextEffectFontSize
    enOOBShapeRangeTextEffectKernedPairs
    enOOBShapeRangeTextEffectNormalizedHeight
    enOOBShapeRangeTextEffectPresetShape
    enOOBShapeRangeTextEffectPresetTextEffect
    enOOBShapeRangeTextEffectRotatedChars
    enOOBShapeRangeTextEffectText
    enOOBShapeRangeTextFrame2AutoSize
    enOOBShapeRangeTextFrame2HorizontalAnchor
    enOOBShapeRangeTextFrame2MarginBottom
    enOOBShapeRangeTextFrame2MarginLeft
    enOOBShapeRangeTextFrame2MarginRight
    enOOBShapeRangeTextFrame2MarginTop
    enOOBShapeRangeTextFrame2NoTextRotation
    enOOBShapeRangeTextFrame2Orientation
    enOOBShapeRangeTextFrame2PathFormat
    enOOBShapeRangeTextFrame2VerticalAnchor
    enOOBShapeRangeTextFrame2WarpFormat
    enOOBShapeRangeTextFrame2WordArtformat
    enOOBShapeRangeTextFrameAutoMargins
    enOOBShapeRangeTextFrameAutoSize
    enOOBShapeRangeTextFrameCharactersCaption
    enOOBShapeRangeTextFrameCharactersTextFontBackground
    enOOBShapeRangeTextFrameCharactersTextFontBold
    enOOBShapeRangeTextFrameCharactersTextFontColor
    enOOBShapeRangeTextFrameCharactersTextFontColorIndex
    enOOBShapeRangeTextFrameCharactersTextFontFontStyle
    enOOBShapeRangeTextFrameCharactersTextFontItalic
    enOOBShapeRangeTextFrameCharactersTextFontName
    enOOBShapeRangeTextFrameCharactersTextFontSize
    enOOBShapeRangeTextFrameCharactersTextFontStrikethrough
    enOOBShapeRangeTextFrameCharactersTextFontSubscript
    enOOBShapeRangeTextFrameCharactersTextFontSuperscript
    enOOBShapeRangeTextFrameCharactersTextFontThemeColor
    enOOBShapeRangeTextFrameCharactersTextFontThemeFont
    enOOBShapeRangeTextFrameCharactersTextFontUnderline
    enOOBShapeRangeTextFrameHorizontalAlignment
    enOOBShapeRangeTextFrameHorizontalOverflow
    enOOBShapeRangeTextFrameMarginBottom
    enOOBShapeRangeTextFrameMarginLeft
    enOOBShapeRangeTextFrameMarginRight
    enOOBShapeRangeTextFrameMarginTop
    enOOBShapeRangeTextFrameOrientation
    enOOBShapeRangeTextFrameReadingOrder
    enOOBShapeRangeTextFrameVerticalAlignment
    enOOBShapeRangeTextFrameVerticalOverflow
    enOOBShapeRangeTitle
    enOOBShapeRangeTop
    enOOBShapeRangeVisible
    enOOBShapeRangeWidth
    enOOBSourceName
    enOOBTop
    enOOBVisible
    enOOBWidth
    enOOBZZZZZZZZZZZZZ
    enShapeAAAAAAAAAAA
    enShapeAccelerator
    enShapeAlignment
    enShapeAlternativeText
    enShapeAltHTML
    enShapeAutoLoad
    enShapeAutoShapeType
    enShapeAutoSize
    enShapeAutoUpdate
    enShapeBackColor
    enShapeBackgroundStyle
    enShapeBackStyle
    enShapeBlackWhiteMode
    enShapeBorderColor
    enShapeBorderColorIndex
    enShapeBorderLineStyle
    enShapeBorderParent
    enShapeBordersSuppress
    enShapeBorderThemeColor
    enShapeBorderTintAndShade
    enShapeBorderWeight
    enShapeCaption
    enShapeControlFormatEnabled
    enShapeControlFormatDropDownLines
    enShapeControlFormatLargeChange
    enShapeControlFormatLinkedCell
    enShapeControlFormatListFillRange
    enShapeControlFormatListIndex
    enShapeControlFormatLockedText
    enShapeControlFormatMax
    enShapeControlFormatMin
    enShapeControlFormatMultiSelect
    enShapeControlFormatPrintObject
    enShapeControlFormatSmallChange
    enShapeDecorative
    enShapeEnabled
    enShapeForeColorBrightness
    enShapeForeColorObjectThemeColor
    enShapeForeColorRGB
    enShapeForeColorSchmeColor
    enShapeForeColorTintAndShade
    enShapeGraphicStyle
    enShapeGroupName
    enShapeHeight
    enShapeLeft
    enShapeLinkedCell
    enShapeListFillRange
    enShapeLockAspectRatio
    enShapeLocked
    enShapeMouseIcon
    enShapeMousePointer
    enShapeMultiSelect
    enShapeName
    enShapeOnAction
    enShapePicture
    enShapePicturePosition
    enShapePlacement
    enShapePrintObject
    enShapeRotation
    enShapeShadowBlur
    enShapeShadowForeColorObjectThemeColor
'    enShapeShadowForeColorRGB
    enShapeShadowForeColorSchmeColor
    enShapeShadowForeColorTintAndShade
    enShapeShadowObscured
    enShapeShadowRotateWithShape
    enShapeShadowSize
    enShapeShadowStyle
    enShapeShadowTransparency
    enShapeShadowType
    enShapeShadowVisible
    enShapeShapeStyle
    enShapeSourceName
    enShapeSpecialEffect
    enShapeTextEffectFontBold
    enShapeTextEffectFontItalic
    enShapeTextEffectFontName
    enShapeTextEffectFontSize
    enShapeTextEffectKernedPairs
    enShapeTextEffectNormalizedHeight
    enShapeTextEffectPresetShape
    enShapeTextEffectPresetTextEffect
    enShapeTextEffectRotatedChars
    enShapeTextEffectText
    enShapeTextFrame2AutoSize
    enShapeTextFrame2HorizontalAnchor
    enShapeTextFrame2MarginBottom
    enShapeTextFrame2MarginLeft
    enShapeTextFrame2MarginRight
    enShapeTextFrame2MarginTop
    enShapeTextFrame2NoTextRotation
    enShapeTextFrame2Orientation
    enShapeTextFrame2PathFormat
    enShapeTextFrame2VerticalAnchor
    enShapeTextFrame2WarpFormat
    enShapeTextFrame2WordArtformat
    enShapeTextFrameAutoMargins
    enShapeTextFrameAutoSize
    enShapeTextFrameCharactersCaption
    enShapeTextFrameCharactersText
    enShapeTextFrameHorizontalAlignment
    enShapeTextFrameHorizontalOverflow
    enShapeTextFrameMarginBottom
    enShapeTextFrameMarginLeft
    enShapeTextFrameMarginRight
    enShapeTextFrameMarginTop
    enShapeTextFrameOrientation
    enShapeTextFrameReadingOrder
    enShapeTextFrameVerticalAlignment
    enShapeTextFrameVerticalOverflow
    enShapeTitle
    enShapeTop
    enShapeTripleState
    enShapeVisible
    enShapeWidth
    enShapeWordWrap
    enShapeZZZZZZZZZZZZZZ
    enShape_Font_Reserved
End Enum

Private oobSource           As OLEObject    ' module global, set by SyncProperties
Private oobTarget           As OLEObject    ' module global, set by SyncProperties
Private shpTarget           As Shape        ' module global, set by SyncProperties
Private shpSource           As Shape        ' module global, set by SyncProperties
Private wshTarget           As Worksheet    ' module global, set by SyncProperties
Private wshSource           As Worksheet    ' module global, set by SyncProperties
Private dctDueSyncAssert    As Dictionary   ' Properties synched but still pending assertion
Private enProperty          As enProperties

Private Property Get enPropertiesOOBFirst() As enProperties:     enPropertiesOOBFirst = enOOBAAAAAAAA + 1:           End Property

Private Property Get enPropertiesOOBLast() As enProperties:      enPropertiesOOBLast = enOOBZZZZZZZZZZZZZ - 1:       End Property

Private Property Get enPropertiesShapeFirst() As enProperties:   enPropertiesShapeFirst = enShapeAAAAAAAAAAA + 1:    End Property

Private Property Get enPropertiesShapeLast() As enProperties:    enPropertiesShapeLast = enShapeZZZZZZZZZZZZZZ - 1:  End Property

Private Property Let OLEObjectSource(ByVal oob As OLEObject):    Set oobSource = oob:    End Property

Private Property Let OLEObjectTarget(ByVal oob As OLEObject):    Set oobTarget = oob:    End Property

Private Property Get PropertyMaxLen(Optional ByVal en_from As enProperties = -1, _
                                   Optional ByVal en_to As enProperties = -1) As Long
    If en_from = -1 Then en_from = enPropertiesOOBFirst
    If en_to = -1 Then en_to = enPropertiesShapeLast
    For enProperty = en_from To en_to
        lPropertyMaxLen = mBasic.Max(lPropertyMaxLen, Len(mSyncShapePrprtys.PropertyName(enProperty)))
    Next enProperty
    PropertyMaxLen = lPropertyMaxLen
End Property

Public Property Let ShapeSource(ByVal shp As Shape):            Set shpSource = shp:    End Property

Public Property Let ShapeTarget(ByVal shp As Shape):            Set shpTarget = shp:    End Property

Private Property Let SheetSource(ByVal wsh As Worksheet):       Set wshSource = wsh:    End Property

Private Property Let SheetTarget(ByVal wsh As Worksheet):       Set wshTarget = wsh:    End Property

Public Sub CollectChanged(ByVal c_shp_source As Shape)
' ------------------------------------------------------------------------------
' Collects each changed property as a due synchronization.
' ------------------------------------------------------------------------------
    Const PROC = "CollectChanged"
    
    On Error GoTo eh
    Dim vTarget     As Variant
    Dim vSource     As Variant
    Dim sId         As String
    
    mBasic.BoP ErrSrc(PROC)
    If c_shp_source.Type = msoOLEControlObject Then
        '~~ Synchronize all properties of the oobTarget with the oobSource
        Set oobSource = shpSource.OLEFormat.Object
        
        If Not mSyncShapes.CorrespondingOOB(shpSource, wshTarget, oobTarget) Is Nothing Then
            For enProperty = mSyncShapePrprtys.enPropertiesOOBFirst To mSyncShapePrprtys.enPropertiesOOBLast
                '~~ Loop through all OOB property enumerations
                sId = SyncId(enProperty, c_shp_source)
                Services.LogItem = PropertyServiced(shpSource, wshTarget)
                On Error Resume Next
                vTarget = mSyncShapePrprtys.PropertyValue(enProperty, oobTarget)
                On Error Resume Next
                vSource = mSyncShapePrprtys.PropertyValue(enProperty, oobSource)
                If Err.Number <> 0 Then GoTo np1
                If vTarget <> vSource Then
                    mSync.DueSyncLet , PropertyName(enProperty), enSyncActionChangeShapePrprty, , sId
                End If
np1:        Next enProperty
        Else
            Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "No corresponding target OOB found for the source Shape " & mSyncShapes.ShapeNames(shpSource)
        End If
    Else ' any other type of shape
        '~~ Synchronize the properties of the shpTarget with the shpSource
        If Not mSyncShapes.CorrespondingShape(shpSource, wshTarget, shpTarget) Is Nothing Then
            For enProperty = mSyncShapePrprtys.enPropertiesShapeFirst To mSyncShapePrprtys.enPropertiesShapeLast
                sId = SyncId(enProperty, c_shp_source)
                '~~ Loop through all Shape property enumerations
                On Error Resume Next
                vTarget = mSyncShapePrprtys.PropertyValue(enProperty, shpTarget)
                If Err.Number <> 0 Then GoTo np2
                On Error Resume Next
                vSource = mSyncShapePrprtys.PropertyValue(enProperty, shpSource)
                If Err.Number <> 0 Then GoTo np2
                If vTarget <> vSource Then
                    mSync.DueSyncLet , PropertyName(enProperty), enSyncActionChangeShapePrprty, , sId
                End If
np2:        Next enProperty
        Else
            Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "No corresponding target Shape found for the source Shape " & mSyncShapes.ShapeNames(shpSource)
        End If
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncShapePrprtys." & s
End Function

                                                           
Private Function PropertyChange(ByVal pc_target As Variant, _
                                ByVal pc_source As Variant) As String
    PropertyChange = " from >" & PropertyValueString(pc_target) & "< to >" & PropertyValueString(pc_source) & "<"
End Function

Private Function PropertyName(ByVal enProperty As enProperties) As String
' ------------------------------------------------------------------------
' Returns the name of the Property (enProperty). Used for logging.
' ------------------------------------------------------------------------
    
    Select Case enProperty
        Case enOOBAutoLoad:                                             PropertyName = "OOB.AutoLoad"
        Case enOOBBorderColor:                                          PropertyName = "OOB.Border.Color"
        Case enOOBBorderColorIndex:                                     PropertyName = "OOB.Border.ColorIndex"
        Case enOOBBorderLineStyle:                                      PropertyName = "OOB.Border.LineStyle"
        Case enOOBObjectBackColor:                                      PropertyName = "OOB.Object.BackColor"
        Case enOOBObjectBackStyle:                                      PropertyName = "OOB.Object.BackStyle"
        Case enOOBObjectCaption:                                        PropertyName = "OOB.Object.Caption"
        Case enOOBObjectEnabled:                                        PropertyName = "OOB.Object.Enabled"
        Case enOOBObjectFontBold:                                       PropertyName = "OOB.Object.Font.Bold"
        Case enOOBObjectFontItalic:                                     PropertyName = "OOB.Object.Font.Italic"
        Case enOOBObjectFontName:                                       PropertyName = "OOB.Object.Font.Name"
        Case enOOBObjectFontSize:                                       PropertyName = "OOB.Object.Font.Size"
        Case enOOBObjectForeColor:                                      PropertyName = "OOB.Object.ForeColor"
        Case enOOBHeight:                                               PropertyName = "OOB.Height"
        Case enOOBLeft:                                                 PropertyName = "OOB.Left"
        Case enOOBObjectLocked:                                         PropertyName = "OOB.Object.Locked"
        Case enOOBObjectMouseIcon:                                      PropertyName = "OOB.Object.MouseIcon"
        Case enOOBObjectMousePointer:                                   PropertyName = "OOB.Object.MousePointer"
        Case enOOBObjectPicture:                                        PropertyName = "OOB.Object.Picture"
        Case enOOBObjectPictureposition:                                PropertyName = "OOB.Object.Pictureposition"
        Case enOOBPlacement:                                            PropertyName = "OOB.Placement"
        Case enOOBPrintObject:                                          PropertyName = "OOB.PrintObject"
        Case enOOBShadow:                                               PropertyName = "OOB.Shadow"
        Case enOOBObjectTakeFocusOnClick:                               PropertyName = "OOB.Object.TakeFocusOnClick"
        Case enOOBTop:                                                  PropertyName = "OOB.Top"
        Case enOOBVisible:                                              PropertyName = "OOB.Visible"
        Case enOOBWidth:                                                PropertyName = "OOB.Width"
        Case enOOBObjectWordWrap:                                       PropertyName = "OOB.Object.WordWrap"
        Case enOOBLinkedCell:                                           PropertyName = "OOB.LinkedCell"
        Case enOOBName:                                                 PropertyName = "OOB.Name"
        Case enOOBObjectCaption:                                        PropertyName = "OOB.Object.Caption"
        Case enOOBObjectForeColor:                                      PropertyName = "OOB.Object.ForeColor"
        Case enOOBOLEType:                                              PropertyName = "OOB.OLEType"
        Case enOOBShapeRangeAlternativeText:                            PropertyName = "OOB.ShapeRange.AlternativeText"
        Case enOOBShapeRangeAutoShapeType:                              PropertyName = "OOB.ShapeRange.AutoShapeType"
        Case enOOBShapeRangeBackgroundStyle:                            PropertyName = "OOB.ShapeRange.BackgroundStyle"
        Case enOOBShapeRangeBlackWhiteMode:                             PropertyName = "OOB.ShapeRange.BlackWhiteMode"
        Case enOOBShapeRangeDecorative:                                 PropertyName = "OOB.ShapeRange.Decorative"
        Case enOOBShapeRangeGraphicStyle:                               PropertyName = "OOB.ShapeRange.GraphicStyle"
        Case enOOBShapeRangeHeight:                                     PropertyName = "OOB.ShapeRange.Height"
        Case enOOBShapeRangeLeft:                                       PropertyName = "OOB.ShapeRange.Left"
        Case enOOBShapeRangeLineBackColor:                              PropertyName = "OOB.ShapeRange.Line.BackColor"
        Case enOOBShapeRangeLineBeginArrowheadLength:                   PropertyName = "OOB.ShapeRange.Line.BeginArrowheadLength"
        Case enOOBShapeRangeLineBeginArrowheadStyle:                    PropertyName = "OOB.ShapeRange.Line.BeginArrowheadStyle"
        Case enOOBShapeRangeLineBeginArrowheadWidth:                    PropertyName = "OOB.ShapeRange.Line.BeginArrowheadWidth"
        Case enOOBShapeRangeLineDashStyle:                              PropertyName = "OOB.ShapeRange.Line.DashStyle"
        Case enOOBShapeRangeLineEndArrowheadLength:                     PropertyName = "OOB.ShapeRange.Line.EndArrowheadLength"
        Case enOOBShapeRangeLineEndArrowheadStyle:                      PropertyName = "OOB.ShapeRange.Line.EndArrowheadStyle"
        Case enOOBShapeRangeLineEndArrowheadWidth:                      PropertyName = "OOB.ShapeRange.Line.EndArrowheadWidth"
        Case enOOBShapeRangeLineForeColorBrightness:                    PropertyName = "OOB.ShapeRange.Line.ForeColor.Brightness"
        Case enOOBShapeRangeLineForeColorObjectThemeColor:              PropertyName = "OOB.ShapeRange.Line.ForeColor.ObjectThemeColor"
        Case enOOBShapeRangeLineForeColorRGB:                           PropertyName = "OOB.ShapeRange.Line.ForeColor.RGB"
        Case enOOBShapeRangeLineForeColorSchemeColor:                   PropertyName = "OOB.ShapeRange.Line.ForeColor.ColorSchemeColor"
        Case enOOBShapeRangeLineForeColorTintAndShade:                  PropertyName = "OOB.ShapeRange.Line.ForeColor.TintAndShade"
        Case enOOBShapeRangeLineForeColorType:                          PropertyName = "OOB.ShapeRange.Line.ForeColor.Type"
        Case enOOBShapeRangeLineInsetPen:                               PropertyName = "OOB.ShapeRange.Line.InsetPen"
        Case enOOBShapeRangeLinePattern:                                PropertyName = "OOB.ShapeRange.Line.Pattern"
        Case enOOBShapeRangeLineStyle:                                  PropertyName = "OOB.ShapeRange.Line.Style"
        Case enOOBShapeRangeLineTransparency:                           PropertyName = "OOB.ShapeRange.Line.Transparency"
        Case enOOBShapeRangeLineVisible:                                PropertyName = "OOB.ShapeRange.Line.Visible"
        Case enOOBShapeRangeLineWeight:                                 PropertyName = "OOB.ShapeRange.Line.Weight"
        Case enOOBShapeRangeLockAspectRatio:                            PropertyName = "OOB.ShapeRange.LockAspectRatio"
        Case enOOBShapeRangeName:                                       PropertyName = "OOB.ShapeRange.Name"
        Case enOOBShapeRangePictureFormatBrightness:                    PropertyName = "OOB.ShapeRange.PictureFormat.Brightness"
        Case enOOBShapeRangePictureFormatColorType:                     PropertyName = "OOB.ShapeRange.PictureFormat.ColorType"
        Case enOOBShapeRangePictureFormatContrast:                      PropertyName = "OOB.ShapeRange.PictureFormat.Contrast"
        Case enOOBShapeRangePictureFormatTransparencyColor:             PropertyName = "OOB.ShapeRange.PictureFormat.TransparencyColor"
        Case enOOBShapeRangePictureFormatTransparentBackground:         PropertyName = "OOB.ShapeRange.PictureFormat.TransparentBackground"
        Case enOOBShapeRangeRotation:                                   PropertyName = "OOB.ShapeRange.Rotation"
        Case enOOBShapeRangeShadowBlur:                                 PropertyName = "OOB.ShapeRange.Shadow.Blur"
        Case enOOBShapeRangeShadowForeColorBrightness:                  PropertyName = "OOB.ShapeRange.Shadow.ForeColor.Brightness"
        Case enOOBShapeRangeShadowForeColorObjectThemeColor:            PropertyName = "OOB.ShapeRange.Shadow.ForeColor.ObjectThemeColor"
        Case enOOBShapeRangeShadowForeColorRGB:                         PropertyName = "OOB.ShapeRange.Shadow.ForeColor.RGB"
        Case enOOBShapeRangeShadowForeColorSchemeColor:                 PropertyName = "OOB.ShapeRange.Shadow.ForeColor.SchemeColor"
        Case enOOBShapeRangeShadowForeColorTintAndShade:                PropertyName = "OOB.ShapeRange.Shadow.ForeColor.TintAndShade"
        Case enOOBShapeRangeShadowOffsetX:                              PropertyName = "OOB.ShapeRange.Shadow.OffsetX"
        Case enOOBShapeRangeShadowOffsetY:                              PropertyName = "OOB.ShapeRange.Shadow.OffsetY"
        Case enOOBShapeRangeShadowRotateWithShape:                      PropertyName = "OOB.ShapeRange.Shadow.RotateWithShape"
        Case enOOBShapeRangeShadowSize:                                 PropertyName = "OOB.ShapeRange.Shadow.Size"
        Case enOOBShapeRangeShadowStyle:                                PropertyName = "OOB.ShapeRange.Shadow.Style"
        Case enOOBShapeRangeShadowTransparency:                         PropertyName = "OOB.ShapeRange.Shadow.Transparency"
        Case enOOBShapeRangeShadowType:                                 PropertyName = "OOB.ShapeRange.Shadow.Type"
        Case enOOBShapeRangeShadowObscured:                             PropertyName = "OOB.ShapeRange.Shadow.Obscured"
        Case enOOBShapeRangeShadowVisible:                              PropertyName = "OOB.ShapeRange.Shadow.Visible"
        Case enOOBShapeRangeShapeStyle:                                 PropertyName = "OOB.ShapeRange.Shadow.Style"
        Case enOOBShapeRangeSoftEdgeRadius:                             PropertyName = "OOB.ShapeRange.SoftEdge.Radius"
        Case enOOBShapeRangeSoftEdgeType:                               PropertyName = "OOB.ShapeRange.SoftEdge.Type"
        Case enOOBShapeRangeTextEffectFontBold:                         PropertyName = "OOB.ShapeRange.TextEffect.FontBold"
        Case enOOBShapeRangeTextEffectFontItalic:                       PropertyName = "OOB.ShapeRange.TextEffect.FontItalic"
        Case enOOBShapeRangeTextEffectFontName:                         PropertyName = "OOB.ShapeRange.TextEffect.FontName"
        Case enOOBShapeRangeTextEffectFontSize:                         PropertyName = "OOB.ShapeRange.TextEffect.FontSize"
        Case enOOBShapeRangeTextEffectKernedPairs:                      PropertyName = "OOB.ShapeRange.TextEffect.KernedPairs"
        Case enOOBShapeRangeTextEffectNormalizedHeight:                 PropertyName = "OOB.ShapeRange.TextEffect.NormalizedHeight"
        Case enOOBShapeRangeTextEffectPresetShape:                      PropertyName = "OOB.ShapeRange.TextEffect.PresetShape"
        Case enOOBShapeRangeTextEffectPresetTextEffect:                 PropertyName = "OOB.ShapeRange.TextEffect.PresetTextEffect"
        Case enOOBShapeRangeTextEffectRotatedChars:                     PropertyName = "OOB.ShapeRange.TextEffect.RotatedChars"
        Case enOOBShapeRangeTextEffectText:                             PropertyName = "OOB.ShapeRange.TextEffect.Text"
        Case enOOBShapeRangeTextFrame2AutoSize:                         PropertyName = "OOB.ShapeRange.TextFrame2.AutoSize"
        Case enOOBShapeRangeTextFrame2HorizontalAnchor:                 PropertyName = "OOB.ShapeRange.TextFrame2.HorizontalAnchor"
        Case enOOBShapeRangeTextFrame2MarginBottom:                     PropertyName = "OOB.ShapeRange.TextFrame2.MarginBottom"
        Case enOOBShapeRangeTextFrame2MarginLeft:                       PropertyName = "OOB.ShapeRange.TextFrame2.MarginLeft"
        Case enOOBShapeRangeTextFrame2MarginRight:                      PropertyName = "OOB.ShapeRange.TextFrame2.MarginRight"
        Case enOOBShapeRangeTextFrame2MarginTop:                        PropertyName = "OOB.ShapeRange.TextFrame2.MarginTop"
        Case enOOBShapeRangeTextFrame2NoTextRotation:                   PropertyName = "OOB.ShapeRange.TextFrame2.NoTextRotation"
        Case enOOBShapeRangeTextFrame2Orientation:                      PropertyName = "OOB.ShapeRange.TextFrame2.Orientation"
        Case enOOBShapeRangeTextFrame2PathFormat:                       PropertyName = "OOB.ShapeRange.TextFrame2.PathFormat"
        Case enOOBShapeRangeTextFrame2VerticalAnchor:                   PropertyName = "OOB.ShapeRange.TextFrame2.VerticalAnchor"
        Case enOOBShapeRangeTextFrame2WarpFormat:                       PropertyName = "OOB.ShapeRange.TextFrame2.WarpFormat"
        Case enOOBShapeRangeTextFrame2WordArtformat:                    PropertyName = "OOB.ShapeRange.TextFrame2.WordArtformat"
        Case enOOBShapeRangeTextFrameAutoMargins:                       PropertyName = "OOB.ShapeRange.TextFrame.AutoMargins"
        Case enOOBShapeRangeTextFrameAutoSize:                          PropertyName = "OOB.ShapeRange.TextFrame.AutoSize"
        Case enOOBShapeRangeTextFrameCharactersCaption:                 PropertyName = "OOB.ShapeRange.TextFrame.Characters.Caption"
        Case enOOBShapeRangeTextFrameCharactersTextFontBackground:      PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text.FontBackground"
        Case enOOBShapeRangeTextFrameCharactersTextFontBold:            PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text.FontBold"
        Case enOOBShapeRangeTextFrameCharactersTextFontColor:           PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontColorIndex:      PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontFontStyle:       PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontItalic:          PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontName:            PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontSize:            PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontStrikethrough:   PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontSubscript:       PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontSuperscript:     PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontThemeColor:      PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontThemeFont:       PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameCharactersTextFontUnderline:       PropertyName = "OOB.ShapeRange.TextFrame.Characters.Text."
        Case enOOBShapeRangeTextFrameHorizontalAlignment:               PropertyName = "OOB.ShapeRange.TextFrame.HorizontalAlignment"
        Case enOOBShapeRangeTextFrameHorizontalOverflow:                PropertyName = "OOB.ShapeRange.TextFrame.HorizontalOverflow"
        Case enOOBShapeRangeTextFrameMarginBottom:                      PropertyName = "OOB.ShapeRange.TextFrame.MarginBottom"
        Case enOOBShapeRangeTextFrameMarginLeft:                        PropertyName = "OOB.ShapeRange.TextFrame.MarginLeft"
        Case enOOBShapeRangeTextFrameMarginRight:                       PropertyName = "OOB.ShapeRange.TextFrame.MarginRight"
        Case enOOBShapeRangeTextFrameMarginTop:                         PropertyName = "OOB.ShapeRange.TextFrame.MarginTop"
        Case enOOBShapeRangeTextFrameOrientation:                       PropertyName = "OOB.ShapeRange.TextFrame.Orientation"
        Case enOOBShapeRangeTextFrameReadingOrder:                      PropertyName = "OOB.ShapeRange.TextFrame.ReadingOrder"
        Case enOOBShapeRangeTextFrameVerticalAlignment:                 PropertyName = "OOB.ShapeRange.TextFrame.VerticalAlignment"
        Case enOOBShapeRangeTextFrameVerticalOverflow:                  PropertyName = "OOB.ShapeRange.TextFrame.VerticalOverflow"
        Case enOOBShapeRangeTitle:                                      PropertyName = "OOB.ShapeRange.Title"
        Case enOOBShapeRangeTop:                                        PropertyName = "OOB.ShapeRange.Top"
        Case enOOBShapeRangeVisible:                                    PropertyName = "OOB.ShapeRange.Visible"
        Case enOOBShapeRangeWidth:                                      PropertyName = "OOB.ShapeRange.Width"
        Case enOOBSourceName:                                           PropertyName = "OOB.SourceName"
        Case enOOBTop:                                                  PropertyName = "OOB.Top"
        
        Case enShapeAccelerator:                                PropertyName = "Shape.Accelerator"
        Case enShapeAlignment:                                  PropertyName = "Shape.Alignment"
        Case enShapeAlternativeText:                            PropertyName = "Shape.AlternativeText"
        Case enShapeAltHTML:                                    PropertyName = "Shape.AltHTML"
        Case enShapeAutoLoad:                                   PropertyName = "Shape.AutoLoad"
        Case enShapeAutoShapeType:                              PropertyName = "Shape.AutoShapeType"
        Case enShapeAutoSize:                                   PropertyName = "Shape.AutoSize"
        Case enShapeAutoUpdate:                                 PropertyName = "Shape.AutoUpdate"
        Case enShapeBackColor:                                  PropertyName = "Shape.BackColor"
        Case enShapeBackgroundStyle:                            PropertyName = "Shape.BackgroundStyle"
        Case enShapeBackStyle:                                  PropertyName = "Shape.BackStyle"
        Case enShapeBlackWhiteMode:                             PropertyName = "Shape.BlackWhitemode"
        Case enShapeBorderColor:                                PropertyName = "Shape.Border.Color"
        Case enShapeBorderColorIndex:                           PropertyName = "Shape.Border.ColorIndex"
        Case enShapeBorderLineStyle:                            PropertyName = "Shape.Border.LineStyle"
        Case enShapeBorderParent:                               PropertyName = "Shape.Border.Parent"
        Case enShapeBordersSuppress:                            PropertyName = "Shape.BordersSuppress"
        Case enShapeBorderThemeColor:                           PropertyName = "Shape.Border.ThemeColor"
        Case enShapeBorderTintAndShade:                         PropertyName = "Shape.Border.TintAndShade"
        Case enShapeBorderWeight:                               PropertyName = "Shape.Weight"
        Case enShapeCaption:                                    PropertyName = "Shape.Caption"
        Case enShapeControlFormatEnabled:                       PropertyName = "Shape.ControlFormat.Enabled"
        Case enShapeControlFormatDropDownLines:                 PropertyName = "Shape.ControlFormat.DropDownLines"
        Case enShapeControlFormatLargeChange:                   PropertyName = "Shape.ControlFormat.LargeChange"
        Case enShapeControlFormatLinkedCell:                    PropertyName = "Shape.ControlFormat.LinkedCell"
        Case enShapeControlFormatListFillRange:                 PropertyName = "Shape.ControlFormat.ListFillRange"
        Case enShapeControlFormatListIndex:                     PropertyName = "Shape.ControlFormat.ListIndex"
        Case enShapeControlFormatLockedText:                    PropertyName = "Shape.ControlFormat.LockedText"
        Case enShapeControlFormatMax:                           PropertyName = "Shape.ControlFormat.Max"
        Case enShapeControlFormatMin:                           PropertyName = "Shape.ControlFormat.Min"
        Case enShapeControlFormatMultiSelect:                   PropertyName = "Shape.ControlFormat.MultiSelect"
        Case enShapeControlFormatPrintObject:                   PropertyName = "Shape.ControlFormat.PrintObject"
        Case enShapeControlFormatSmallChange:                   PropertyName = "Shape.ControlFormat.SmallChange"
        Case enShapeDecorative:                                 PropertyName = "Shape.Decorative"
        Case enShapeEnabled:                                    PropertyName = "Shape.Enabled"
        Case enShapeForeColorBrightness:                        PropertyName = "Shape.ForeColor.Brightness"
        Case enShapeForeColorObjectThemeColor:                  PropertyName = "Shape.ForeColor.ObjectThemeColor"
        Case enShapeForeColorRGB:                               PropertyName = "Shape.ForeColor.RGB"
        Case enShapeForeColorSchmeColor:                        PropertyName = "Shape.ForeColor.SchemeColor"
        Case enShapeForeColorTintAndShade:                      PropertyName = "Shape.ForeColor.TintAndShade"
        Case enShapeGraphicStyle:                               PropertyName = "Shape.GraphicStyle"
        Case enShapeGroupName:                                  PropertyName = "Shape.GroupName"
        Case enShapeHeight:                                     PropertyName = "Shape.Height"
        Case enShapeLeft:                                       PropertyName = "Shape.Left"
        Case enShapeLinkedCell:                                 PropertyName = "Shape.LinkedCell"
        Case enShapeListFillRange:                              PropertyName = "Shape.ListFillRange"
        Case enShapeLockAspectRatio:                            PropertyName = "Shape.Ratio"
        Case enShapeLocked:                                     PropertyName = "Shape.Locked"
        Case enShapeMouseIcon:                                  PropertyName = "Shape.MouseIcon"
        Case enShapeMousePointer:                               PropertyName = "Shape.MousPointer"
        Case enShapeMultiSelect:                                PropertyName = "Shape.MultiSelect"
        Case enShapeName:                                       PropertyName = "Shape.Name"
        Case enShapeOnAction:                                   PropertyName = "Shape.OnAction"
        Case enShapePicture:                                    PropertyName = "Shape.Picture"
        Case enShapePicturePosition:                            PropertyName = "Shape.PicturePosition"
        Case enShapePlacement:                                  PropertyName = "Shape.Placement"
        Case enShapePrintObject:                                PropertyName = "Shape.PrintObject"
        Case enShapeRotation:                                   PropertyName = "Shape.Rotation"
        Case enShapeShadowBlur:                                 PropertyName = "Shape.Shadow.Blur"
        Case enShapeShadowForeColorObjectThemeColor:            PropertyName = "Shape.Shadow.ForeColor.ObjectThemeColor"
'        Case enShapeShadowForeColorRGB:                         PropertyName = "Shape.Shadow.ForeColor.RGB"
        Case enShapeShadowForeColorSchmeColor:                  PropertyName = "Shape.Shadow.ForeColor.SchemeColor"
        Case enShapeShadowForeColorTintAndShade:                PropertyName = "Shape.Shadow.ForeColor.TintAndShade"
        Case enShapeShadowObscured:                             PropertyName = "Shape.Shadow.Obscured"
        Case enShapeShadowRotateWithShape:                      PropertyName = "Shape.Shadow.RotateWithShape"
        Case enShapeShadowSize:                                 PropertyName = "Shape.Shadow.Size"
        Case enShapeShadowStyle:                                PropertyName = "Shape.Shadow.Style"
        Case enShapeShadowTransparency:                         PropertyName = "Shape.Shadow.Transparency"
        Case enShapeShadowType:                                 PropertyName = "Shape.Shadow.Type"
        Case enShapeShadowVisible:                              PropertyName = "Shape.Shadow.Visible"
        Case enShapeShapeStyle:                                 PropertyName = "Shape.ShapeStyle"
        Case enShapeSourceName:                                 PropertyName = "Shape.SourceName"
        Case enShapeSpecialEffect:                              PropertyName = "Shape.SpecialEffect"
        Case enShapeTextEffectFontBold:                         PropertyName = "Shape.TextEffect.FontBold"
        Case enShapeTextEffectFontItalic:                       PropertyName = "Shape.TextEffect.FontItalic"
        Case enShapeTextEffectFontName:                         PropertyName = "Shape.TextEffect.FontName"
        Case enShapeTextEffectFontSize:                         PropertyName = "Shape.TextEffect.FontSize"
        Case enShapeTextEffectKernedPairs:                      PropertyName = "Shape.TextEffect.KernedPairs"
        Case enShapeTextEffectNormalizedHeight:                 PropertyName = "Shape.TextEffect.NormalizedHeight"
        Case enShapeTextEffectPresetShape:                      PropertyName = "Shape.TextEffect.PresetShape"
        Case enShapeTextEffectPresetTextEffect:                 PropertyName = "Shape.TextEffect.PresetTextEffect"
        Case enShapeTextEffectRotatedChars:                     PropertyName = "Shape.TextEffect.RotatedChars"
        Case enShapeTextEffectText:                             PropertyName = "Shape.TextEffect.text"
        Case enShapeTextFrame2AutoSize:                         PropertyName = "Shape.TextFrame2.AutoSize"
        Case enShapeTextFrame2HorizontalAnchor:                 PropertyName = "Shape.TextFrame2.HorizontalAnchor"
        Case enShapeTextFrame2MarginBottom:                     PropertyName = "Shape.TextFrame2.MarginBottom"
        Case enShapeTextFrame2MarginLeft:                       PropertyName = "Shape.TextFrame2.MarginLeft"
        Case enShapeTextFrame2MarginRight:                      PropertyName = "Shape.TextFrame2.MarginRight"
        Case enShapeTextFrame2MarginTop:                        PropertyName = "Shape.TextFrame2.MarginTop"
        Case enShapeTextFrame2NoTextRotation:                   PropertyName = "Shape.TextFrame2.NoTextRotation"
        Case enShapeTextFrame2Orientation:                      PropertyName = "Shape.TextFrame2.Orientation"
        Case enShapeTextFrame2PathFormat:                       PropertyName = "Shape.TextFrame2.PathFormat"
        Case enShapeTextFrame2VerticalAnchor:                   PropertyName = "Shape.TextFrame2.VerticalAnchor"
        Case enShapeTextFrame2WarpFormat:                       PropertyName = "Shape.TextFrame2.WarpFormat"
        Case enShapeTextFrame2WordArtformat:                    PropertyName = "Shape.TextFrame2.WordArtformat"
        Case enShapeTextFrameAutoMargins:                       PropertyName = "Shape.TextFrame.AutoMargins"
        Case enShapeTextFrameAutoSize:                          PropertyName = "Shape.TextFrame.AutoSize"
        Case enShapeTextFrameCharactersCaption:                 PropertyName = "Shape.TextFrame.Characters.Caption"
        Case enShapeTextFrameCharactersText:                    PropertyName = "Shape.TextFrame.Characters.Text"
        Case enShapeTextFrameHorizontalAlignment:               PropertyName = "Shape.TextFrame.HorizontalAlignment"
        Case enShapeTextFrameHorizontalOverflow:                PropertyName = "Shape.TextFrame.HorizontalOverflow"
        Case enShapeTextFrameMarginBottom:                      PropertyName = "Shape.TextFrame.MarginBottom"
        Case enShapeTextFrameMarginLeft:                        PropertyName = "Shape.TextFrame.MarginLeft"
        Case enShapeTextFrameMarginRight:                       PropertyName = "Shape.TextFrame.MarginRight"
        Case enShapeTextFrameMarginTop:                         PropertyName = "Shape.TextFrame.MarginTop"
        Case enShapeTextFrameOrientation:                       PropertyName = "Shape.TextFrame.Orientation"
        Case enShapeTextFrameReadingOrder:                      PropertyName = "Shape.TextFrame.ReadingOrder"
        Case enShapeTextFrameVerticalAlignment:                 PropertyName = "Shape.TextFrame.VerticalAlignment"
        Case enShapeTextFrameVerticalOverflow:                  PropertyName = "Shape.TextFrame.VerticalOverflow"
        Case enShapeTitle:                                      PropertyName = "Shape.Title"
        Case enShapeTop:                                        PropertyName = "Shape.Top"
        Case enShapeTripleState:                                PropertyName = "Shape.TripleState"
        Case enShapeVisible:                                    PropertyName = "Shape.Visible"
        Case enShapeWidth:                                      PropertyName = "Shape.Width"
        Case enShapeWordWrap:                                   PropertyName = "Shape.WordWrap"
        Case enShape_Font_Reserved:                             PropertyName = "Shape.Font_Reserved"
        Case Else:                                              PropertyName = "for enProperty " & enProperty & " is no Name specified (missing Case after enProperty " & enProperty - 1 & ")"
    End Select
    
End Function

Private Function PropertyServiced(ByVal spv_shp_source As Shape, _
                                  ByVal spv_wsh_target As Worksheet) As String
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    PropertyServiced = mSyncShapes.TypeString(spv_shp_source) & ";" & spv_wsh_target.Name & "." & ShapeNames(spv_shp_source)
End Function

Private Function PropertyValue(ByVal sp_enProperty As enProperties, _
                               ByVal sp_obj As Variant) As String
' ------------------------------------------------------------------------------
' Synchronizes the Properties vf the shape (shpSource) with the corresponding
' shape in the corresponding target sheet in the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Dim oob As OLEObject
    Dim shp As Shape
    
    If TypeName(sp_obj) = "Shape" Then
        Set shp = sp_obj
    ElseIf TypeName(sp_obj) = "OOBObject" Then
        Set oob = sp_obj
    End If
    If TypeName(sp_obj) = "Shape" Then
        If shp.Type = msoOLEControlObject Then Set oob = shp.OLEFormat.Object
    End If
    If Not oob Is Nothing Then
        Select Case sp_enProperty
            Case enOOBAutoLoad:                                     PropertyValue = PropertyValueString(oob.AutoLoad) '  oobSource.Object.AutoLoad
            Case enOOBBorderColor:                                  PropertyValue = oob.Border.Color
            Case enOOBBorderColorIndex:                             PropertyValue = oob.Border.ColorIndex '  oobSource.Border.ColorIndex
            Case enOOBBorderLineStyle:                              PropertyValue = oob.Border.LineStyle '  oobSource.Border.LineStyle
            Case enOOBHeight:                                       PropertyValue = oob.Height
            Case enOOBLeft:                                         PropertyValue = oob.Left
            Case enOOBLinkedCell:                                   PropertyValue = oob.LinkedCell '  oobSource.LinkedCell
            Case enOOBName:                                         PropertyValue = oob.Name '  oobSource.Name
            Case enOOBObjectBackColor:                              PropertyValue = oob.Object.BackColor
            Case enOOBObjectBackStyle:                              PropertyValue = oob.Object.BackStyle
            Case enOOBObjectCaption:                                PropertyValue = oob.Object.Caption '  oobSource.Object.Caption
            Case enOOBObjectEnabled:                                PropertyValue = PropertyValueString(oob.Object.Enabled)
            Case enOOBObjectFontBold:                               PropertyValue = PropertyValueString(oob.Object.Font.Bold)
            Case enOOBObjectFontItalic:                             PropertyValue = PropertyValueString(oob.Object.Font.Italic)
            Case enOOBObjectFontName:                               PropertyValue = oob.Object.Font.Name
            Case enOOBObjectFontSize:                               PropertyValue = oob.Object.Font.Size
            Case enOOBObjectForeColor:                              PropertyValue = oob.Object.ForeColor
            Case enOOBObjectLocked:                                 PropertyValue = PropertyValueString(oob.Object.Locked)
            Case enOOBObjectMouseIcon:                              PropertyValue = PropertyValueString(oob.Object.MouseIcon)
            Case enOOBObjectMousePointer:                           PropertyValue = oob.Object.MousePointer
            Case enOOBObjectPicture:                                PropertyValue = oob.Object.Picture
            Case enOOBObjectPictureposition:                        PropertyValue = oob.Object.PicturePosition
            Case enOOBPlacement:                                    PropertyValue = oob.Placement '  oobSource.Object.Placement
            Case enOOBPrintObject:                                  PropertyValue = PropertyValueString(oob.PrintObject)
            Case enOOBShadow:                                       PropertyValue = PropertyValueString(oob.Shadow)
            Case enOOBObjectTakeFocusOnClick:                       PropertyValue = PropertyValueString(oob.Object.TakeFocusOnClick)
            Case enOOBTop:                                          PropertyValue = oob.Top
            Case enOOBVisible:                                      PropertyValue = PropertyValueString(oob.Visible)
            Case enOOBWidth:                                        PropertyValue = oob.Width
            Case enOOBObjectWordWrap:                               PropertyValue = PropertyValueString(oob.Object.WordWrap)
            Case enOOBOLEType:                                      PropertyValue = oob.OLEType '  oobSource.OLEType
            Case enOOBShapeRangeAlternativeText:                    PropertyValue = oob.ShapeRange.AlternativeText '  oobSource.ShapeRange.AlternativeText
            Case enOOBShapeRangeAutoShapeType:                      PropertyValue = oob.ShapeRange.AutoShapeType '  oobSource.ShapeRange.AutoShapeType
            Case enOOBShapeRangeBackgroundStyle:                    PropertyValue = oob.ShapeRange.BackgroundStyle '  oobSource.ShapeRange.BackgroundStyle
            Case enOOBShapeRangeBlackWhiteMode:                     PropertyValue = oob.ShapeRange.BlackWhiteMode '  oobSource.ShapeRange.BlackWhiteMode
            Case enOOBShapeRangeDecorative:                         PropertyValue = oob.ShapeRange.Decorative '  oobSource.ShapeRange.Decorative
            Case enOOBShapeRangeGraphicStyle:                       PropertyValue = oob.ShapeRange.GraphicStyle '  oobSource.ShapeRange.GraphicStyle
            Case enOOBShapeRangeHeight:                             PropertyValue = oob.ShapeRange.Height '  oobSource.ShapeRange.Height
            Case enOOBShapeRangeLeft:                               PropertyValue = oob.ShapeRange.Left '  oobSource.ShapeRange.Left
            Case enOOBShapeRangeLineBackColor:                      PropertyValue = oob.ShapeRange.Line.BackColor '  oobSource.ShapeRange.Line.BackColor
            Case enOOBShapeRangeLineBeginArrowheadLength:           PropertyValue = oob.ShapeRange.Line.BeginArrowheadLength '  oobSource.ShapeRange.Line.BeginArrowheadLength
            Case enOOBShapeRangeLineBeginArrowheadStyle:            PropertyValue = oob.ShapeRange.Line.BeginArrowheadStyle '  oobSource.ShapeRange.Line.BeginArrowheadStyle
            Case enOOBShapeRangeLineBeginArrowheadWidth:            PropertyValue = oob.ShapeRange.Line.BeginArrowheadWidth '  oobSource.ShapeRange.Line.BeginArrowheadWidth
            Case enOOBShapeRangeLineDashStyle:                      PropertyValue = oob.ShapeRange.Line.DashStyle '  oobSource.ShapeRange.Line.DashStyle
            Case enOOBShapeRangeLineEndArrowheadLength:             PropertyValue = oob.ShapeRange.Line.EndArrowheadLength '  oobSource.ShapeRange.Line.EndArrowheadLength
            Case enOOBShapeRangeLineEndArrowheadStyle:              PropertyValue = oob.ShapeRange.Line.EndArrowheadStyle '  oobSource.ShapeRange.Line.EndArrowheadStyle
            Case enOOBShapeRangeLineEndArrowheadWidth:              PropertyValue = oob.ShapeRange.Line.EndArrowheadWidth '  oobSource.ShapeRange.Line.EndArrowheadWidth
            Case enOOBShapeRangeLineForeColorBrightness:            PropertyValue = oob.ShapeRange.Line.ForeColor.Brightness '  oobSource.ShapeRange.Line.ForeColor.Brightness
            Case enOOBShapeRangeLineForeColorObjectThemeColor:      PropertyValue = oob.ShapeRange.Line.ForeColor.ObjectThemeColor '  oobSource.ShapeRange.Line.ForeColor.ObjectThemeColor
            Case enOOBShapeRangeLineForeColorRGB:                   PropertyValue = oob.ShapeRange.Line.ForeColor.RGB '  oobSource.ShapeRange.Line.ForeColor.RGB
            Case enOOBShapeRangeLineForeColorSchemeColor:           PropertyValue = oob.ShapeRange.Line.ForeColor.SchemeColor '  oobSource.ShapeRange.Line.ForeColor.SchemeColor
            Case enOOBShapeRangeLineForeColorTintAndShade:          PropertyValue = oob.ShapeRange.Line.ForeColor.TintAndShade '  oobSource.ShapeRange.Line.ForeColor.TintAndShade
            Case enOOBShapeRangeLineForeColorType:                  PropertyValue = oob.ShapeRange.Line.ForeColor.Type '  oobSource.ShapeRange.Line.ForeColor.Type
            Case enOOBShapeRangeLineInsetPen:                       PropertyValue = oob.ShapeRange.Line.InsetPen '  oobSource.ShapeRange.Line.InsetPen
            Case enOOBShapeRangeLinePattern:                        PropertyValue = oob.ShapeRange.Line.Pattern '  oobSource.ShapeRange.Line.Pattern
            Case enOOBShapeRangeLineStyle:                          PropertyValue = oob.ShapeRange.Line.Style '  oobSource.ShapeRange.Line.Style
            Case enOOBShapeRangeLineTransparency:                   PropertyValue = oob.ShapeRange.Line.Transparency '  oobSource.ShapeRange.Line.Transparency
            Case enOOBShapeRangeLineVisible:                        PropertyValue = oob.ShapeRange.Line.Visible '  oobSource.ShapeRange.Line.Visible
            Case enOOBShapeRangeLineWeight:                         PropertyValue = oob.ShapeRange.Line.Weight '  oobSource.ShapeRange.Line.Weight
            Case enOOBShapeRangeLockAspectRatio:                    PropertyValue = oob.ShapeRange.LockAspectRatio '  oobSource.ShapeRange.LockAspectRatio
            Case enOOBShapeRangeName:                               PropertyValue = oob.ShapeRange.Name '  oobSource.ShapeRange.Name
            Case enOOBShapeRangePictureFormatBrightness:            PropertyValue = oob.ShapeRange.PictureFormat.Brightness '  oobSource.ShapeRange.Picture.Brightness
            Case enOOBShapeRangePictureFormatColorType:             PropertyValue = oob.ShapeRange.PictureFormat.ColorType '  oobSource.ShapeRange.PictureFormat.ColorType
            Case enOOBShapeRangePictureFormatContrast:              PropertyValue = oob.ShapeRange.PictureFormat.Contrast '  oobSource.ShapeRange.PictureFormat.Contrast
            Case enOOBShapeRangePictureFormatTransparencyColor:     PropertyValue = oob.ShapeRange.PictureFormat.TransparencyColor '  oobSource.ShapeRange.PictureFormat.TransparencyColor
            Case enOOBShapeRangePictureFormatTransparentBackground: PropertyValue = oob.ShapeRange.PictureFormat.TransparentBackground '  oobSource.ShapeRange.PictureFormat.TransparentBackground
            Case enOOBShapeRangeRotation:                           PropertyValue = oob.ShapeRange.Rotation '  oobSource.ShapeRange.Rotation
            Case enOOBShapeRangeShadowBlur:                         PropertyValue = oob.ShapeRange.Shadow.Blur '  oobSource.ShapeRange.Shadow.Blur
            Case enOOBShapeRangeShadowForeColorBrightness:          PropertyValue = oob.ShapeRange.Shadow.ForeColor.Brightness '  oobSource.ShapeRange.Shadow.ForeColor.Brightness
            Case enOOBShapeRangeShadowForeColorObjectThemeColor:    PropertyValue = oob.ShapeRange.Shadow.ForeColor.ObjectThemeColor '  oobSource.ShapeRange.Shadow.ForeColor.ObjectThemeColor
            Case enOOBShapeRangeShadowForeColorRGB:                 PropertyValue = oob.ShapeRange.Shadow.ForeColor.RGB '  oobSource.ShapeRange.Shadow.ForeColor.RGB
            Case enOOBShapeRangeShadowForeColorSchemeColor:         PropertyValue = oob.ShapeRange.Shadow.ForeColor.SchemeColor '  oobSource.ShapeRange.Shadow.ForeColor.SchemeColor
            Case enOOBShapeRangeShadowForeColorTintAndShade:        PropertyValue = oob.ShapeRange.Shadow.ForeColor.TintAndShade '  oobSource.ShapeRange.Shadow.ForeColor.TintAndShade
            Case enOOBShapeRangeShadowObscured:                     PropertyValue = oob.ShapeRange.Shadow.Obscured '  oobSource.ShapeRange.Shadow.Obscured
            Case enOOBShapeRangeShadowOffsetX:                      PropertyValue = oob.ShapeRange.Shadow.OffsetX '  oobSource.ShapeRange.Shadow.OffsetX
            Case enOOBShapeRangeShadowOffsetY:                      PropertyValue = oob.ShapeRange.Shadow.OffsetY '  oobSource.ShapeRange.Shadow.OffsetY
            Case enOOBShapeRangeShadowRotateWithShape:              PropertyValue = oob.ShapeRange.Shadow.RotateWithShape '  oobSource.ShapeRange.Shadow.RotateWithShape
            Case enOOBShapeRangeShadowSize:                         PropertyValue = oob.ShapeRange.Shadow.Size '  oobSource.ShapeRange.Shadow.Size
            Case enOOBShapeRangeShadowStyle:                        PropertyValue = oob.ShapeRange.Shadow.Style '  oobSource.ShapeRange.Shadow.Style
            Case enOOBShapeRangeShadowTransparency:                 PropertyValue = oob.ShapeRange.Shadow.Transparency '  oobSource.ShapeRange.Shadow.Transparency
            Case enOOBShapeRangeShadowType:                         PropertyValue = oob.ShapeRange.Shadow.Type '  oobSource.ShapeRange.Shadow.Type
            Case enOOBShapeRangeShadowVisible:                      PropertyValue = oob.ShapeRange.Shadow.Visible '  oobSource.ShapeRange.Shadow.Visible
            Case enOOBShapeRangeShapeStyle:                         PropertyValue = oob.ShapeRange.ShapeStyle '  oobSource.ShapeRange.ShapeStyle
            Case enOOBShapeRangeSoftEdgeRadius:                     PropertyValue = oob.ShapeRange.SoftEdge.Radius '  oobSource.ShapeRange.SoftEdge.Radius
            Case enOOBShapeRangeSoftEdgeType:                       PropertyValue = oob.ShapeRange.SoftEdge.Type '  oobSource.ShapeRange.SoftEdge.Type
            Case enOOBShapeRangeTextEffectFontBold:                 PropertyValue = oob.ShapeRange.TextEffect.FontBold '  oobSource.ShapeRange.TextEffect.FontBold
            Case enOOBShapeRangeTextEffectFontItalic:               PropertyValue = oob.ShapeRange.TextEffect.FontItalic '  oobSource.ShapeRange.TextEffect.FontItalic
            Case enOOBShapeRangeTextEffectFontName:                 PropertyValue = oob.ShapeRange.TextEffect.FontName '  oobSource.ShapeRange.TextEffect.FontName
            Case enOOBShapeRangeTextEffectFontSize:                 PropertyValue = oob.ShapeRange.TextEffect.FontSize '  oobSource.ShapeRange.TextEffect.FontSize
            Case enOOBShapeRangeTextEffectKernedPairs:              PropertyValue = oob.ShapeRange.TextEffect.KernedPairs '  oobSource.ShapeRange.TextEffect.KernedPairs
            Case enOOBShapeRangeTextEffectNormalizedHeight:         PropertyValue = oob.ShapeRange.TextEffect.NormalizedHeight '  oobSource.ShapeRange.TextEffect.NormalizedHeight
            Case enOOBShapeRangeTextEffectPresetShape:              PropertyValue = oob.ShapeRange.TextEffect.PresetShape '  oobSource.ShapeRange.TextEffect.PresetShape
            Case enOOBShapeRangeTextEffectPresetTextEffect:         PropertyValue = oob.ShapeRange.TextEffect.PresetTextEffect '  oobSource.ShapeRange.TextEffect.PresetTextEffect
            Case enOOBShapeRangeTextEffectRotatedChars:             PropertyValue = oob.ShapeRange.TextEffect.RotatedChars '  oobSource.ShapeRange.TextEffect.RotatedChars
            Case enOOBShapeRangeTextEffectText:                     PropertyValue = oob.ShapeRange.TextEffect.Text '  oobSource.ShapeRange.TextEffect.Text
            Case enOOBShapeRangeTextFrame2AutoSize:                 PropertyValue = oob.ShapeRange.TextFrame2.AutoSize '  oobSource.ShapeRange.TextFrame2.AutoSize
            Case enOOBShapeRangeTextFrame2HorizontalAnchor:         PropertyValue = oob.ShapeRange.TextFrame2.HorizontalAnchor '  oobSource.ShapeRange.TextFrame2.HorizontalAnchor
            Case enOOBShapeRangeTextFrame2MarginBottom:             PropertyValue = oob.ShapeRange.TextFrame2.MarginBottom '  oobSource.ShapeRange.TextFrame2.MarginBottom
            Case enOOBShapeRangeTextFrame2MarginLeft:               PropertyValue = oob.ShapeRange.TextFrame2.MarginRight '  oobSource.ShapeRange.TextFrame2.MarginRight
            Case enOOBShapeRangeTextFrame2MarginRight:              PropertyValue = oob.ShapeRange.TextFrame2.MarginLeft '  oobSource.ShapeRange.TextFrame2.MarginLeft
            Case enOOBShapeRangeTextFrame2MarginTop:                PropertyValue = oob.ShapeRange.TextFrame2.MarginTop '  oobSource.ShapeRange.TextFrame2.MarginTop
            Case enOOBShapeRangeTextFrame2NoTextRotation:           PropertyValue = oob.ShapeRange.TextFrame2.NoTextRotation '  oobSource.ShapeRange.TextFrame2.NoTextRotation
            Case enOOBShapeRangeTextFrame2Orientation:              PropertyValue = oob.ShapeRange.TextFrame2.Orientation '  oobSource.ShapeRange.TextFrame2.Orientation
            Case enOOBShapeRangeTextFrame2PathFormat:               PropertyValue = oob.ShapeRange.TextFrame2.PathFormat '  oobSource.ShapeRange.TextFrame2.PathFormat
            Case enOOBShapeRangeTextFrame2WarpFormat:               PropertyValue = oob.ShapeRange.TextFrame2.WarpFormat '  oobSource.ShapeRange.TextFrame2.WarpFormat
            Case enOOBShapeRangeTextFrame2WordArtformat:            PropertyValue = oob.ShapeRange.TextFrame2.WordArtformat '  oobSource.ShapeRange.TextFrame2.WordArtformat
            Case enOOBShapeRangeTextFrameAutoMargins:               PropertyValue = oob.ShapeRange.TextFrame.AutoMargins '  oobSource.ShapeRange.TextFrame.AutoMargins
            Case enOOBShapeRangeTextFrameAutoSize:                  PropertyValue = oob.ShapeRange.TextFrame.AutoSize '  oobSource.ShapeRange.TextFrame.AutoSize
            Case enOOBShapeRangeTextFrameCharactersCaption:         PropertyValue = oob.ShapeRange.TextFrame.Characters.Caption '  oobSource.ShapeRange.TextFrame.Characters.Caption
            Case enOOBShapeRangeTextFrameHorizontalAlignment:       PropertyValue = oob.ShapeRange.TextFrame.HorizontalAlignment '  oobSource.ShapeRange.TextFrame.HorizontalAlignment
            Case enOOBShapeRangeTextFrameHorizontalOverflow:        PropertyValue = oob.ShapeRange.TextFrame.HorizontalOverflow '  oobSource.ShapeRange.TextFrame.HorizontalOverflow
            Case enOOBShapeRangeTextFrameMarginBottom:              PropertyValue = oob.ShapeRange.TextFrame.MarginBottom '  oobSource.ShapeRange.TextFrame.MarginBottom
            Case enOOBShapeRangeTextFrameMarginLeft:                PropertyValue = oob.ShapeRange.TextFrame.MarginRight '  oobSource.ShapeRange.TextFrame.MarginRight
            Case enOOBShapeRangeTextFrameMarginRight:               PropertyValue = oob.ShapeRange.TextFrame.MarginLeft '  oobSource.ShapeRange.TextFrame.MarginLeft
            Case enOOBShapeRangeTextFrameMarginTop:                 PropertyValue = oob.ShapeRange.TextFrame.MarginTop '  oobSource.ShapeRange.TextFrame.MarginTop
            Case enOOBShapeRangeTextFrameOrientation:               PropertyValue = oob.ShapeRange.TextFrame.Orientation '  oobSource.ShapeRange.TextFrame.Orientation
            Case enOOBShapeRangeTextFrameReadingOrder:              PropertyValue = oob.ShapeRange.TextFrame.ReadingOrder '  oobSource.ShapeRange.TextFrame.ReadingOrder
            Case enOOBShapeRangeTextFrameVerticalAlignment:         PropertyValue = oob.ShapeRange.TextFrame.VerticalAlignment '  oobSource.ShapeRange.TextFrame.VerticalAlignment
            Case enOOBShapeRangeTextFrameVerticalOverflow:          PropertyValue = oob.ShapeRange.TextFrame.VerticalOverflow '  oobSource.ShapeRange.TextFrame.VerticalOverflow
            Case enOOBShapeRangeTitle:                              PropertyValue = oob.ShapeRange.Title '  oobSource.ShapeRange.TextFrame.VerticalOverflow
            Case enOOBShapeRangeTop:                                PropertyValue = oob.ShapeRange.Top '  oobSource.ShapeRange.Top
            Case enOOBShapeRangeVisible:                            PropertyValue = oob.ShapeRange.Visible '  oobSource.ShapeRange.Visible
            Case enOOBShapeRangeWidth:                              PropertyValue = oob.ShapeRange.Width '  oobSource.ShapeRange.Width
            Case enOOBSourceName:                                   PropertyValue = oob.SourceName '  oobSource.SourceName
        End Select
    Else
        Select Case sp_enProperty
            Case enShapeAccelerator:                                PropertyValue = shp.Accelerator '  shpSource.Accelerator
            Case enShapeAlignment:                                  PropertyValue = shp.Alignment '  shpSource.Alignment
            Case enShapeAlternativeText:                            PropertyValue = shp.AlternativeText '  shpSource.AlternativeText
            Case enShapeAltHTML:                                    PropertyValue = shp.AltHTML '  shpSource.AltHTML
            Case enShapeAutoLoad:                                   PropertyValue = shp.AutoLoad '  shpSource.AutoLoad
            Case enShapeAutoShapeType:                              PropertyValue = shp.AutoShapeType '  shpSource.AutoShapeType
            Case enShapeAutoSize:                                   PropertyValue = shp.AutoSize '  shpSource.AutoSize
            Case enShapeAutoUpdate:                                 PropertyValue = shp.AutoUpdate '  shpSource.AutoUpdate
            Case enShapeBackColor:                                  PropertyValue = shp.BackColor '  shpSource.BackColor
            Case enShapeBackgroundStyle:                            PropertyValue = shp.BackgroundStyle '  shpSource.BackgroundStyle
            Case enShapeBackStyle:                                  PropertyValue = shp.BackStyle '  shpSource.BackStyle
            Case enShapeBlackWhiteMode:                             PropertyValue = shp.BlackWhiteMode '  shpSource.BlackWhiteMode
            Case enShapeBorderColor:                                PropertyValue = shp.BorderColor '  shpSource.BorderColor
            Case enShapeBorderColorIndex:                           PropertyValue = shp.BorderColorIndex '  shpSource.BorderColorIndex
            Case enShapeBorderLineStyle:                            PropertyValue = shp.BorderLineStyle '  shpSource.BorderLineStyle
            Case enShapeBorderParent:                               PropertyValue = shp.BorderParent '  shpSource.BorderParent
            Case enShapeBordersSuppress:                            PropertyValue = shp.BordersSuppress '  shpSource.BordersSuppress
            Case enShapeBorderThemeColor:                           PropertyValue = shp.BorderThemeColor '  shpSource.BorderThemeColor
            Case enShapeBorderTintAndShade:                         PropertyValue = shp.BorderTintAndShade '  shpSource.BorderTintAndShade
            Case enShapeBorderWeight:                               PropertyValue = shp.BorderWeight '  shpSource.BorderWeight
            Case enShapeCaption:                                    PropertyValue = shp.Caption '  shpSource.Caption
            Case enShapeControlFormatEnabled:                       PropertyValue = shp.ControlFormat.Enabled
            Case enShapeControlFormatDropDownLines:                 PropertyValue = shp.ControlFormat.DropDownLines
            Case enShapeControlFormatLargeChange:                   PropertyValue = shp.ControlFormat.LargeChange
            Case enShapeControlFormatLinkedCell:                    PropertyValue = shp.ControlFormat.LinkedCell
            Case enShapeControlFormatListFillRange:                 PropertyValue = shp.ControlFormat.ListFillRange
            Case enShapeControlFormatListIndex:                     PropertyValue = shp.ControlFormat.ListIndex
            Case enShapeControlFormatLockedText:                    PropertyValue = shp.ControlFormat.LockedText
            Case enShapeControlFormatMax:                           PropertyValue = shp.ControlFormat.Max
            Case enShapeControlFormatMin:                           PropertyValue = shp.ControlFormat.Min
            Case enShapeControlFormatMultiSelect:                   PropertyValue = shp.ControlFormat.MultiSelect
            Case enShapeControlFormatPrintObject:                   PropertyValue = shp.ControlFormat.PrintObject
            Case enShapeControlFormatSmallChange:                   PropertyValue = shp.ControlFormat.SmallChange
            Case enShapeDecorative:                                 PropertyValue = shp.Decorative '  shpSource.Decorative
            Case enShapeEnabled:                                    PropertyValue = shp.Enabled '  shpSource.Enabled
            Case enShapeForeColorBrightness:                        PropertyValue = shp.ForeColor.Brightness
            Case enShapeForeColorObjectThemeColor:                  PropertyValue = shp.ForeColor.ObjectThemeColor
            Case enShapeForeColorRGB:                               PropertyValue = shp.ForeColor.RGB
            Case enShapeForeColorSchmeColor:                        PropertyValue = shp.ForeColor.SchemeColor
            Case enShapeForeColorTintAndShade:                      PropertyValue = shp.ForeColor.TintAndShade
            Case enShapeGraphicStyle:                               PropertyValue = shp.GraphicStyle '  shpSource.GraphicStyle
            Case enShapeGroupName:                                  PropertyValue = shp.GroupName '  shpSource.GroupName
            Case enShapeHeight:                                     PropertyValue = Round(shp.Height, 0)
            Case enShapeLeft:                                       PropertyValue = shp.Left '  shpSource.Left
            Case enShapeLinkedCell:                                 PropertyValue = shp.LinkedCell '  shpSource.LinkedCell
            Case enShapeListFillRange:                              PropertyValue = shp.ListFillRange '  shpSource.ListFillRange
            Case enShapeLockAspectRatio:                            PropertyValue = shp.LockAspectRatio '  shpSource.LockAspectRatio
            Case enShapeLocked:                                     PropertyValue = shp.Locked '  shpSource.Locked
            Case enShapeMouseIcon:                                  PropertyValue = shp.MouseIcon '  shpSource.MouseIcon
            Case enShapeMousePointer:                               PropertyValue = shp.MousePointer '  shpSource.MousePointer
            Case enShapeMultiSelect:                                PropertyValue = shp.MultiSelect '  shpSource.MultiSelect
            Case enShapeName:                                       PropertyValue = shp.Name '  shpSource.Name
            Case enShapeOnAction:                                   PropertyValue = shp.OnAction '  shpSource.OnAction
            Case enShapePicture:                                    PropertyValue = shp.Picture '  shpSource.Picture
            Case enShapePicturePosition:                            PropertyValue = shp.PicturePosition '  shpSource.PicturePosition
            Case enShapePlacement:                                  PropertyValue = shp.Placement '  shpSource.Placement
            Case enShapePrintObject:                                PropertyValue = shp.PrintObject '  shpSource.PrintObject
            Case enShapeRotation:                                   PropertyValue = shp.Rotation '  shpSource.Rotation
            Case enShapeShadowBlur:                                 PropertyValue = shp.Shadow.Blur '  shpSource.Shadow.Blur
            Case enShapeShadowForeColorObjectThemeColor:            PropertyValue = shp.Shadow.ForeColor.ObjectThemeColor
'            Case enShapeShadowForeColorRGB:                         PropertyValue = shp.Shadow.ForeColor.RGB
            Case enShapeShadowForeColorSchmeColor:                  PropertyValue = shp.Shadow.ForeColor.SchemeColor
            Case enShapeShadowForeColorTintAndShade:                PropertyValue = shp.Shadow.ForeColor.TintAndShade
            Case enShapeShadowObscured:                             PropertyValue = shp.Shadow.Obscured '  shpSource.Shadow.Obscured
            Case enShapeShadowRotateWithShape:                      PropertyValue = shp.Shadow.RotateWithShape '  shpSource.Shadow.RotateWithShape
            Case enShapeShadowSize:                                 PropertyValue = shp.Shadow.Size '  shpSource.Shadow.Size
            Case enShapeShadowStyle:                                PropertyValue = shp.Shadow.Style '  shpSource.Shadow.Style
            Case enShapeShadowTransparency:                         PropertyValue = shp.Shadow.Transparency '  shpSource.Shadow.Transparency
            Case enShapeShadowType:                                 PropertyValue = shp.Shadow.Type '  shpSource.Shadow.Type
            Case enShapeShadowVisible:                              PropertyValue = shp.Shadow.Visible '  shpSource.Shadow.Visible
            Case enShapeShapeStyle:                                 PropertyValue = shp.ShapeStyle '  shpSource.ShapeStyle
            Case enShapeSourceName:                                 PropertyValue = shp.SourceName '  shpSource.SourceName
            Case enShapeSpecialEffect:                              PropertyValue = shp.SpecialEffect '  shpSource.SpecialEffect
            Case enShapeTextEffectFontBold:                         PropertyValue = shp.TextEffect.FontBold '  shpSource.TextEffect.FontBold
            Case enShapeTextEffectFontItalic:                       PropertyValue = shp.TextEffect.FontItalic '  shpSource.TextEffect.FontItalic
            Case enShapeTextEffectFontName:                         PropertyValue = shp.TextEffect.FontName '  shpSource.TextEffect.FontName
            Case enShapeTextEffectFontSize:                         PropertyValue = shp.TextEffect.FontSize '  shpSource.TextEffect.FontSize
            Case enShapeTextEffectKernedPairs:                      PropertyValue = shp.TextEffect.KernedPairs '  shpSource.TextEffect.KernedPairs
            Case enShapeTextEffectNormalizedHeight:                 PropertyValue = shp.TextEffect.NormalizedHeight '  shpSource.TextEffect.NormalizedHeight
            Case enShapeTextEffectPresetShape:                      PropertyValue = shp.TextEffect.PresetShape '  shpSource.TextEffect.PresetShape
            Case enShapeTextEffectPresetTextEffect:                 PropertyValue = shp.TextEffect.PresetTextEffect '  shpSource.TextEffect.PresetTextEffect
            Case enShapeTextEffectRotatedChars:                     PropertyValue = shp.TextEffect.RotatedChars '  shpSource.TextEffect.RotatedChars
            Case enShapeTextEffectText:                             PropertyValue = shp.TextEffect.Text '  shpSource.TextEffect.Text
            Case enShapeTextFrame2AutoSize:                         PropertyValue = shp.TextFrame2.AutoSize '  shpSource.TextFrame2.AutoSize
            Case enShapeTextFrame2HorizontalAnchor:                 PropertyValue = shp.TextFrame2.HorizontalAnchor '  shpSource.TextFrame2.HorizontalAnchor
            Case enShapeTextFrame2MarginBottom:                     PropertyValue = shp.TextFrame2.MarginBottom '  shpSource.TextFrame2.MarginBottom
            Case enShapeTextFrame2MarginLeft:                       PropertyValue = shp.TextFrame2.MarginRight '  shpSource.TextFrame2.MarginRight
            Case enShapeTextFrame2MarginRight:                      PropertyValue = shp.TextFrame2.MarginLeft '  shpSource.TextFrame2.MarginLeft
            Case enShapeTextFrame2MarginTop:                        PropertyValue = shp.TextFrame2.MarginTop '  shpSource.TextFrame2.MarginTop
            Case enShapeTextFrame2NoTextRotation:                   PropertyValue = shp.TextFrame2.NoTextRotation '  shpSource.TextFrame2.NoTextRotation
            Case enShapeTextFrame2Orientation:                      PropertyValue = shp.TextFrame2.Orientation '  shpSource.TextFrame2.Orientation
            Case enShapeTextFrame2PathFormat:                       PropertyValue = shp.TextFrame2.PathFormat '  shpSource.TextFrame2.PathFormat
            Case enShapeTextFrame2WarpFormat:                       PropertyValue = shp.TextFrame2.WarpFormat '  shpSource.TextFrame2.WarpFormat
            Case enShapeTextFrame2WordArtformat:                    PropertyValue = shp.TextFrame2.WordArtformat '  shpSource.TextFrame2.WordArtformat
            Case enShapeTextFrameAutoMargins:                       PropertyValue = shp.TextFrame.AutoMargins '  shpSource.TextFrame.AutoMargins
            Case enShapeTextFrameAutoSize:                          PropertyValue = shp.TextFrame.AutoSize '  shpSource.TextFrame.AutoSize
            Case enShapeTextFrameCharactersCaption:                 PropertyValue = shp.TextFrame.Characters.Caption '  shpSource.TextFrame.Characters.Caption
            Case enShapeTextFrameCharactersText:                    PropertyValue = shp.TextFrame.Characters.Text
            Case enShapeTextFrameHorizontalAlignment:               PropertyValue = shp.TextFrame.HorizontalAlignment '  shpSource.TextFrame.HorizontalAlignment
            Case enShapeTextFrameHorizontalOverflow:                PropertyValue = shp.TextFrame.HorizontalOverflow '  shpSource.TextFrame.HorizontalOverflow
            Case enShapeTextFrameMarginBottom:                      PropertyValue = shp.TextFrame.MarginBottom '  shpSource.TextFrame.MarginBottom
            Case enShapeTextFrameMarginLeft:                        PropertyValue = shp.TextFrame.MarginRight '  shpSource.TextFrame.MarginRight
            Case enShapeTextFrameMarginRight:                       PropertyValue = shp.TextFrame.MarginLeft '  shpSource.TextFrame.MarginLeft
            Case enShapeTextFrameMarginTop:                         PropertyValue = shp.TextFrame.MarginTop '  shpSource.TextFrame.MarginTop
            Case enShapeTextFrameOrientation:                       PropertyValue = shp.TextFrame.Orientation '  shpSource.TextFrame.Orientation
            Case enShapeTextFrameReadingOrder:                      PropertyValue = shp.TextFrame.ReadingOrder '  shpSource.TextFrame.ReadingOrder
            Case enShapeTextFrameVerticalAlignment:                 PropertyValue = shp.TextFrame.VerticalAlignment '  shpSource.TextFrame.VerticalAlignment
            Case enShapeTextFrameVerticalOverflow:                  PropertyValue = shp.TextFrame.VerticalOverflow '  shpSource.TextFrame.VerticalOverflow
            Case enShapeTitle:                                      PropertyValue = shp.Title '  shpSource.Title
            Case enShapeTop:                                        PropertyValue = shp.Top '  shpSource.Top
            Case enShapeVisible:                                    PropertyValue = shp.Visible '  shpSource.Visible
            Case enShapeWidth:                                      PropertyValue = shp.Width '  shpSource.Width
            Case enShapeWordWrap:                                   PropertyValue = shp.WordWrap '  shpSource.WordWrap
        End Select
    End If

xt: Exit Function

End Function

Private Function PropertyValueString(ByVal v As Variant) As String
    Select Case TypeName(v)
        Case "Boolean"
            If v Then PropertyValueString = "True" Else PropertyValueString = "False"
        Case "Range"
            PropertyValueString = v.Address
        Case Else
            PropertyValueString = v
    End Select
End Function

Private Sub SyncAssertion(ByVal sa_obj_source As Variant, _
                          ByVal sa_en_property As enProperties, _
                          ByVal sa_target_value As Variant, _
                          ByVal sa_source_value As Variant, _
                          ByRef sa_dct_synched As Dictionary)
' ------------------------------------------------------------------------------
' Check whether the Property (sa_en_property) has been synchronized
' (sa_target_value = sa_source_value checked by SyncPropertyAsserted) and add a
' corresponding entry to the Dichtionary (sa_dct_synched) with the concerned
' shape's name as item.
' ------------------------------------------------------------------------------

    If SyncPropertyAsserted(sa_en_property) Then
        If Not sa_dct_synched.Exists(PropertyName(sa_en_property)) Then
            mDct.DctAdd sa_dct_synched, PropertyName(sa_en_property), sa_obj_source.Parent.Name & "." & mSyncShapes.ShapeNames(sa_obj_source), , seq_ascending
        End If
        LogServiced.ColsItems "change", "Shape-Property", PropertyName(sa_en_property), "changed", PropertyChange(sa_target_value, sa_source_value)
    Else
        LogServiced.ColsItems "change", "Shape-Property", PropertyName(sa_en_property), "failed", " "
    End If

End Sub

Private Sub SynchabilityCheckOOB(ByVal sc_property As String, _
                                 ByRef sc_dvt_read_write As Dictionary)
' ------------------------------------------------------------------------------
' Check r/w for the current property and if yes add a corresponding entry to the
' Dictionary (sc_dvt_read_write).
' ------------------------------------------------------------------------------
    Dim oobTargetBkp    As OLEObject
    
    Set oobTargetBkp = oobTarget
    Set oobTarget = oobSource
    On Error Resume Next
    mSyncShapePrprtys.SyncProperty enProperty
    If Err.Number = 0 Then
        If Not sc_dvt_read_write.Exists(sc_property) Then
            '~~ read/write proved
            mDct.DctAdd sc_dvt_read_write, sc_property, ": synchronizability proved   (example " & wshSource.Name & "." & ShapeNames(oobSource) & ")", order_bykey, seq_ascending, sense_casesensitive
        End If
    End If
    Set oobTarget = oobTargetBkp

End Sub

Private Sub SynchabilityCheckShape(ByVal sc_en_property As enProperties, _
                                   ByVal sc_property As String, _
                                   ByRef sc_dvt_read_write As Dictionary)
' ------------------------------------------------------------------------------
' Check r/w for the current property and if yes add a corresponding entry to the
' Dictionary (sc_dvt_read_write).
' ------------------------------------------------------------------------------
    Dim shpTargetBkp    As Shape
    
    Set shpTargetBkp = shpTarget
    Set shpTarget = shpSource
    On Error Resume Next
    mSyncShapePrprtys.SyncProperty sc_en_property
    If Err.Number = 0 Then
        If Not sc_dvt_read_write.Exists(sc_property) Then
            '~~ read/write proved
            mDct.DctAdd sc_dvt_read_write, sc_property, ": synchronizability proved     (example " & wshSource.Name & "." & ShapeNames(shpSource) & ")", order_bykey, seq_ascending, sense_casesensitive
        End If
    End If
    Set shpTarget = shpTargetBkp

End Sub

Private Sub SynchedCheckOOB(ByVal sc_en_property As enProperties, _
                            ByVal sc_property As String, _
                            ByVal sc_property_max_len As Long, _
                            ByRef sc_dct_synched As Dictionary)
' ------------------------------------------------------------------------------
' Check whether the current property (sc_en_property) has been synchronized and
' add a corresponding entry to the Dictionary (sc_dct_synched).
' ------------------------------------------------------------------------------
    Const PROC = "SynchedCheckOOB"
    
    Dim vTarget As Variant
    Dim vSource As Variant
    
    vTarget = mSyncShapePrprtys.PropertyValue(sc_en_property, oobTarget)
    vSource = mSyncShapePrprtys.PropertyValue(sc_en_property, oobSource)
    If mSyncShapePrprtys.PropertyValue(sc_en_property, oobTarget) = mSyncShapePrprtys.PropertyValue(sc_en_property, oobSource) Then
        If Not sc_dct_synched.Exists(sc_property) Then
            '~~ read/write proved
            mDct.DctAdd sc_dct_synched, sc_property, ": synchronizability proved     (example " & wshSource.Name & "." & ShapeNames(oobSource) & ")", order_bykey, seq_ascending, sense_casesensitive
        Else
            Debug.Print ErrSrc(PROC) & ": " & mBasic.Align(sc_property, sc_property_max_len, , , ".") & " not changed " & mSyncShapePrprtys.PropertyChange(vTarget, vSource)
            Stop
        End If
    End If

End Sub

Private Function SyncId(ByVal s_en_property As enProperties, _
                        ByVal s_shp As Shape)
    SyncId = mSyncShapes.SyncId(s_shp) & "." & PropertyName(s_en_property)
End Function

Public Sub SyncProperties(ByRef sp_dct_sync_asserted As Dictionary)
' ------------------------------------------------------------------------------
' Synchronizes all the target Shape's (sp_shp_target) - the target OOB's
' respectively - applicable (read/write) properties with the source shape
' (sp_shp_source).
' ------------------------------------------------------------------------------
    Const PROC = "SyncProperties"

    On Error GoTo eh
    Dim vTarget     As Variant
    Dim vSource     As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set dctDueSyncAssert = New Dictionary
    Set wshSource = shpSource.Parent
    Set wshTarget = shpTarget.Parent
    
    If shpSource.Type = msoOLEControlObject Then
        '~~ Synchronize all properties of the oobTarget with the oobSource
        Set oobSource = shpSource.OLEFormat.Object
        
        If Not mSyncShapes.CorrespondingOOB(shpSource, wshTarget, oobTarget) Is Nothing Then
            For enProperty = mSyncShapePrprtys.enPropertiesOOBFirst To mSyncShapePrprtys.enPropertiesOOBLast
                '~~ Loop through all OOB property enumerations
                Services.LogItem = PropertyServiced(shpSource, wshTarget)
                vTarget = mSyncShapePrprtys.PropertyValue(enProperty, oobTarget)
                vSource = mSyncShapePrprtys.PropertyValue(enProperty, oobSource)
                If vTarget <> vSource Then
                    mSyncShapePrprtys.SyncProperty enProperty
                    SyncAssertion shpSource, enProperty, vSource, vTarget, sp_dct_sync_asserted
                End If
            Next enProperty
        Else
            Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "No corresponding target OOB found for the source Shape " & mSyncShapes.ShapeNames(shpSource)
        End If
    Else ' any other type of shape
        '~~ Synchronize the properties of the shpTarget with the shpSource
        If Not mSyncShapes.CorrespondingShape(shpSource, wshTarget, shpTarget) Is Nothing Then
            For enProperty = mSyncShapePrprtys.enPropertiesShapeFirst To mSyncShapePrprtys.enPropertiesShapeLast
                '~~ Loop through all Shape property enumerations
                Services.LogItem = PropertyServiced(shpSource, wshTarget)
                On Error Resume Next
                vTarget = mSyncShapePrprtys.PropertyValue(enProperty, shpTarget)
                If Err.Number <> 0 Then GoTo np
                On Error Resume Next
                vSource = mSyncShapePrprtys.PropertyValue(enProperty, shpSource)
                If Err.Number <> 0 Then GoTo np
                If vTarget <> vSource Then
                    mSyncShapePrprtys.SyncProperty enProperty
                    SyncAssertion shpSource, enProperty, vTarget, vSource, sp_dct_sync_asserted
                End If
np:         Next enProperty
        Else
            Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "No corresponding target Shape found for the source Shape " & mSyncShapes.ShapeNames(shpSource)
        End If
    End If
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SyncProperty(ByVal sp_en_property As enProperties)
' ------------------------------------------------------------------------------
' Synchronizes the Properties vf the shape (shpSource) with the corresponding
' shape in the corresponding target sheet in the Sync-Target-Workbook.
' Note: oobTarget And oobSource are global for this module and set by
'       SyncProperties.
' ------------------------------------------------------------------------------
    
    If sp_en_property = 202 Then Stop
    With oobTarget
        Select Case sp_en_property
            Case enOOBAutoLoad:                                     oobTarget.AutoLoad = oobSource.AutoLoad
            Case enOOBBorderColor:                                  oobTarget.Border.Color = oobSource.Border.Color
            Case enOOBBorderColorIndex:                             oobTarget.Border.ColorIndex = oobSource.Border.ColorIndex
            Case enOOBBorderLineStyle:                              oobTarget.Border.LineStyle = oobSource.Border.LineStyle
            Case enOOBHeight:                                       oobTarget.Height = oobSource.Height
            Case enOOBLeft:                                         oobTarget.Left = oobSource.Left
            Case enOOBLinkedCell:                                   oobTarget.LinkedCell = oobSource.LinkedCell
            Case enOOBName:                                         oobTarget.Name = oobSource.Name
            Case enOOBObjectCaption:                                oobTarget.Object.Caption = oobSource.Object.Caption
            Case enOOBObjectEnabled:                                oobTarget.Object.Enabled = oobSource.Object.Enabled
            Case enOOBObjectForeColor:                              oobTarget.Object.ForeColor = oobSource.Object.ForeColor
            Case enOOBObjectBackColor:                              oobTarget.Object.BackColor = oobSource.Object.BackColor
            Case enOOBObjectBackStyle:                              oobTarget.Object.BackStyle = oobSource.Object.BackStyle
            Case enOOBObjectFontBold:                               oobTarget.Object.Font.Bold = oobSource.Object.Font.Bold
            Case enOOBObjectFontItalic:                             oobTarget.Object.Font.Italic = oobSource.Object.Font.Italic
            Case enOOBObjectFontName:                               oobTarget.Object.Font.Name = oobSource.Object.Font.Name
            Case enOOBObjectFontSize:                               oobTarget.Object.Font.Size = oobSource.Object.Font.Size
            Case enOOBObjectLocked:                                 oobTarget.Object.Locked = oobSource.Object.Locked
            Case enOOBObjectMouseIcon:                              oobTarget.Object.MouseIcon = oobSource.Object.MouseIcon
            Case enOOBObjectMousePointer:                           oobTarget.Object.MousePointer = oobSource.Object.MousePointer
            Case enOOBObjectPicture:                                oobTarget.Object.Picture = oobSource.Object.Picture
            Case enOOBObjectPictureposition:                        oobTarget.Object.PicturePosition = oobSource.Object.PicturePosition
            Case enOOBObjectTakeFocusOnClick:                       oobTarget.Object.TakeFocusOnClick = oobSource.Object.TakeFocusOnClick
            Case enOOBObjectWordWrap:                               oobTarget.Object.WordWrap = oobSource.Object.WordWrap
            Case enOOBOLEType:                                      oobTarget.OLEType = oobSource.OLEType
            Case enOOBPlacement:                                    oobTarget.Placement = oobSource.Placement
            Case enOOBPrintObject:                                  oobTarget.PrintObject = oobSource.PrintObject
            Case enOOBShadow:                                       oobTarget.Shadow = oobSource.Shadow
            Case enOOBShapeRangeAlternativeText:                    oobTarget.ShapeRange.AlternativeText = oobSource.ShapeRange.AlternativeText
            Case enOOBShapeRangeAutoShapeType:                      oobTarget.ShapeRange.AutoShapeType = oobSource.ShapeRange.AutoShapeType
            Case enOOBShapeRangeBackgroundStyle:                    oobTarget.ShapeRange.BackgroundStyle = oobSource.ShapeRange.BackgroundStyle
            Case enOOBShapeRangeBlackWhiteMode:                     oobTarget.ShapeRange.BlackWhiteMode = oobSource.ShapeRange.BlackWhiteMode
            Case enOOBShapeRangeDecorative:                         oobTarget.ShapeRange.Decorative = oobSource.ShapeRange.Decorative
            Case enOOBShapeRangeGraphicStyle:                       oobTarget.ShapeRange.GraphicStyle = oobSource.ShapeRange.GraphicStyle
            Case enOOBShapeRangeHeight:                             oobTarget.ShapeRange.Height = oobSource.ShapeRange.Height
            Case enOOBShapeRangeLeft:                               oobTarget.ShapeRange.Left = oobSource.ShapeRange.Left
            Case enOOBShapeRangeLineBackColor:                      oobTarget.ShapeRange.Line.BackColor = oobSource.ShapeRange.Line.BackColor
            Case enOOBShapeRangeLineBeginArrowheadLength:           oobTarget.ShapeRange.Line.BeginArrowheadLength = oobSource.ShapeRange.Line.BeginArrowheadLength
            Case enOOBShapeRangeLineBeginArrowheadStyle:            oobTarget.ShapeRange.Line.BeginArrowheadStyle = oobSource.ShapeRange.Line.BeginArrowheadStyle
            Case enOOBShapeRangeLineBeginArrowheadWidth:            oobTarget.ShapeRange.Line.BeginArrowheadWidth = oobSource.ShapeRange.Line.BeginArrowheadWidth
            Case enOOBShapeRangeLineDashStyle:                      oobTarget.ShapeRange.Line.DashStyle = oobSource.ShapeRange.Line.DashStyle
            Case enOOBShapeRangeLineEndArrowheadLength:             oobTarget.ShapeRange.Line.EndArrowheadLength = oobSource.ShapeRange.Line.EndArrowheadLength
            Case enOOBShapeRangeLineEndArrowheadStyle:              oobTarget.ShapeRange.Line.EndArrowheadStyle = oobSource.ShapeRange.Line.EndArrowheadStyle
            Case enOOBShapeRangeLineEndArrowheadWidth:              oobTarget.ShapeRange.Line.EndArrowheadWidth = oobSource.ShapeRange.Line.EndArrowheadWidth
            Case enOOBShapeRangeLineForeColorBrightness:            oobTarget.ShapeRange.Line.ForeColor.Brightness = oobSource.ShapeRange.Line.ForeColor.Brightness
            Case enOOBShapeRangeLineForeColorObjectThemeColor:      oobTarget.ShapeRange.Line.ForeColor.ObjectThemeColor = oobSource.ShapeRange.Line.ForeColor.ObjectThemeColor
            Case enOOBShapeRangeLineForeColorRGB:                   oobTarget.ShapeRange.Line.ForeColor.RGB = oobSource.ShapeRange.Line.ForeColor.RGB
            Case enOOBShapeRangeLineForeColorSchemeColor:           oobTarget.ShapeRange.Line.ForeColor.SchemeColor = oobSource.ShapeRange.Line.ForeColor.SchemeColor
            Case enOOBShapeRangeLineForeColorTintAndShade:          oobTarget.ShapeRange.Line.ForeColor.TintAndShade = oobSource.ShapeRange.Line.ForeColor.TintAndShade
            Case enOOBShapeRangeLineInsetPen:                       oobTarget.ShapeRange.Line.InsetPen = oobSource.ShapeRange.Line.InsetPen
            Case enOOBShapeRangeLinePattern:                        oobTarget.ShapeRange.Line.Pattern = oobSource.ShapeRange.Line.Pattern
            Case enOOBShapeRangeLineStyle:                          oobTarget.ShapeRange.Line.Style = oobSource.ShapeRange.Line.Style
            Case enOOBShapeRangeLineTransparency:                   oobTarget.ShapeRange.Line.Transparency = oobSource.ShapeRange.Line.Transparency
            Case enOOBShapeRangeLineVisible:                        oobTarget.ShapeRange.Line.Visible = oobSource.ShapeRange.Line.Visible
            Case enOOBShapeRangeLineWeight:                         oobTarget.ShapeRange.Line.Weight = oobSource.ShapeRange.Line.Weight
            Case enOOBShapeRangeLockAspectRatio:                    oobTarget.ShapeRange.LockAspectRatio = oobSource.ShapeRange.LockAspectRatio
            Case enOOBShapeRangeName:                               oobTarget.ShapeRange.Name = oobSource.ShapeRange.Name
            Case enOOBShapeRangePictureFormatBrightness:            oobTarget.ShapeRange.PictureFormat.Brightness = oobSource.ShapeRange.Picture.Brightness
            Case enOOBShapeRangePictureFormatColorType:             oobTarget.ShapeRange.PictureFormat.ColorType = oobSource.ShapeRange.PictureFormat.ColorType
            Case enOOBShapeRangePictureFormatContrast:              oobTarget.ShapeRange.PictureFormat.Contrast = oobSource.ShapeRange.PictureFormat.Contrast
            Case enOOBShapeRangePictureFormatTransparencyColor:     oobTarget.ShapeRange.PictureFormat.TransparencyColor = oobSource.ShapeRange.PictureFormat.TransparencyColor
            Case enOOBShapeRangePictureFormatTransparentBackground: oobTarget.ShapeRange.PictureFormat.TransparentBackground = oobSource.ShapeRange.PictureFormat.TransparentBackground
            Case enOOBShapeRangeRotation:                           oobTarget.ShapeRange.Rotation = oobSource.ShapeRange.Rotation
            Case enOOBShapeRangeShadowBlur:                         oobTarget.ShapeRange.Shadow.Blur = oobSource.ShapeRange.Shadow.Blur
            Case enOOBShapeRangeShadowForeColorBrightness:          oobTarget.ShapeRange.Shadow.ForeColor.Brightness = oobSource.ShapeRange.Shadow.ForeColor.Brightness
            Case enOOBShapeRangeShadowForeColorObjectThemeColor:    oobTarget.ShapeRange.Shadow.ForeColor.ObjectThemeColor = oobSource.ShapeRange.Shadow.ForeColor.ObjectThemeColor
            Case enOOBShapeRangeShadowForeColorRGB:                 oobTarget.ShapeRange.Shadow.ForeColor.RGB = oobSource.ShapeRange.Shadow.ForeColor.RGB
            Case enOOBShapeRangeShadowForeColorSchemeColor:         oobTarget.ShapeRange.Shadow.ForeColor.SchemeColor = oobSource.ShapeRange.Shadow.ForeColor.SchemeColor
            Case enOOBShapeRangeShadowForeColorTintAndShade:        oobTarget.ShapeRange.Shadow.ForeColor.TintAndShade = oobSource.ShapeRange.Shadow.ForeColor.TintAndShade
            Case enOOBShapeRangeShadowObscured:                     oobTarget.ShapeRange.Shadow.Obscured = oobSource.ShapeRange.Shadow.Obscured
            Case enOOBShapeRangeShadowOffsetX:                      oobTarget.ShapeRange.Shadow.OffsetX = oobSource.ShapeRange.Shadow.OffsetX
            Case enOOBShapeRangeShadowOffsetY:                      oobTarget.ShapeRange.Shadow.OffsetY = oobSource.ShapeRange.Shadow.OffsetY
            Case enOOBShapeRangeShadowRotateWithShape:              oobTarget.ShapeRange.Shadow.RotateWithShape = oobSource.ShapeRange.Shadow.RotateWithShape
            Case enOOBShapeRangeShadowSize:                         oobTarget.ShapeRange.Shadow.Size = oobSource.ShapeRange.Shadow.Size
            Case enOOBShapeRangeShadowStyle:                        oobTarget.ShapeRange.Shadow.Style = oobSource.ShapeRange.Shadow.Style
            Case enOOBShapeRangeShadowTransparency:                 oobTarget.ShapeRange.Shadow.Transparency = oobSource.ShapeRange.Shadow.Transparency
            Case enOOBShapeRangeShadowType:                         oobTarget.ShapeRange.Shadow.Type = oobSource.ShapeRange.Shadow.Type
            Case enOOBShapeRangeShadowVisible:                      oobTarget.ShapeRange.Shadow.Visible = oobSource.ShapeRange.Shadow.Visible
            Case enOOBShapeRangeShapeStyle:                         oobTarget.ShapeRange.ShapeStyle = oobSource.ShapeRange.ShapeStyle
            Case enOOBShapeRangeSoftEdgeRadius:                     oobTarget.ShapeRange.SoftEdge.Radius = oobSource.ShapeRange.SoftEdge.Radius
            Case enOOBShapeRangeSoftEdgeType:                       oobTarget.ShapeRange.SoftEdge.Type = oobSource.ShapeRange.SoftEdge.Type
            Case enOOBShapeRangeTextEffectFontBold:                 oobTarget.ShapeRange.TextEffect.FontBold = oobSource.ShapeRange.TextEffect.FontBold
            Case enOOBShapeRangeTextEffectFontItalic:               oobTarget.ShapeRange.TextEffect.FontItalic = oobSource.ShapeRange.TextEffect.FontItalic
            Case enOOBShapeRangeTextEffectFontName:                 oobTarget.ShapeRange.TextEffect.FontName = oobSource.ShapeRange.TextEffect.FontName
            Case enOOBShapeRangeTextEffectFontSize:                 oobTarget.ShapeRange.TextEffect.FontSize = oobSource.ShapeRange.TextEffect.FontSize
            Case enOOBShapeRangeTextEffectKernedPairs:              oobTarget.ShapeRange.TextEffect.KernedPairs = oobSource.ShapeRange.TextEffect.KernedPairs
            Case enOOBShapeRangeTextEffectNormalizedHeight:         oobTarget.ShapeRange.TextEffect.NormalizedHeight = oobSource.ShapeRange.TextEffect.NormalizedHeight
            Case enOOBShapeRangeTextEffectPresetShape:              oobTarget.ShapeRange.TextEffect.PresetShape = oobSource.ShapeRange.TextEffect.PresetShape
            Case enOOBShapeRangeTextEffectPresetTextEffect:         oobTarget.ShapeRange.TextEffect.PresetTextEffect = oobSource.ShapeRange.TextEffect.PresetTextEffect
            Case enOOBShapeRangeTextEffectRotatedChars:             oobTarget.ShapeRange.TextEffect.RotatedChars = oobSource.ShapeRange.TextEffect.RotatedChars
            Case enOOBShapeRangeTextEffectText:                     oobTarget.ShapeRange.TextEffect.Text = oobSource.ShapeRange.TextEffect.Text
            Case enOOBShapeRangeTextFrame2AutoSize:                 oobTarget.ShapeRange.TextFrame2.AutoSize = oobSource.ShapeRange.TextFrame2.AutoSize
            Case enOOBShapeRangeTextFrame2HorizontalAnchor:         oobTarget.ShapeRange.TextFrame2.HorizontalAnchor = oobSource.ShapeRange.TextFrame2.HorizontalAnchor
            Case enOOBShapeRangeTextFrame2MarginBottom:             oobTarget.ShapeRange.TextFrame2.MarginBottom = oobSource.ShapeRange.TextFrame2.MarginBottom
            Case enOOBShapeRangeTextFrame2MarginLeft:               oobTarget.ShapeRange.TextFrame2.MarginRight = oobSource.ShapeRange.TextFrame2.MarginRight
            Case enOOBShapeRangeTextFrame2MarginRight:              oobTarget.ShapeRange.TextFrame2.MarginLeft = oobSource.ShapeRange.TextFrame2.MarginLeft
            Case enOOBShapeRangeTextFrame2MarginTop:                oobTarget.ShapeRange.TextFrame2.MarginTop = oobSource.ShapeRange.TextFrame2.MarginTop
            Case enOOBShapeRangeTextFrame2NoTextRotation:           oobTarget.ShapeRange.TextFrame2.NoTextRotation = oobSource.ShapeRange.TextFrame2.NoTextRotation
            Case enOOBShapeRangeTextFrame2Orientation:              oobTarget.ShapeRange.TextFrame2.Orientation = oobSource.ShapeRange.TextFrame2.Orientation
            Case enOOBShapeRangeTextFrame2PathFormat:               oobTarget.ShapeRange.TextFrame2.PathFormat = oobSource.ShapeRange.TextFrame2.PathFormat
            Case enOOBShapeRangeTextFrame2WarpFormat:               oobTarget.ShapeRange.TextFrame2.WarpFormat = oobSource.ShapeRange.TextFrame2.WarpFormat
            Case enOOBShapeRangeTextFrame2WordArtformat:            oobTarget.ShapeRange.TextFrame2.WordArtformat = oobSource.ShapeRange.TextFrame2.WordArtformat
            Case enOOBShapeRangeTextFrameAutoMargins:               oobTarget.ShapeRange.TextFrame.AutoMargins = oobSource.ShapeRange.TextFrame.AutoMargins
            Case enOOBShapeRangeTextFrameAutoSize:                  oobTarget.ShapeRange.TextFrame.AutoSize = oobSource.ShapeRange.TextFrame.AutoSize
            Case enOOBShapeRangeTextFrameCharactersCaption:         oobTarget.ShapeRange.TextFrame.Characters.Caption = oobSource.ShapeRange.TextFrame.Characters.Caption
            Case enOOBShapeRangeTextFrameHorizontalAlignment:       oobTarget.ShapeRange.TextFrame.HorizontalAlignment = oobSource.ShapeRange.TextFrame.HorizontalAlignment
            Case enOOBShapeRangeTextFrameHorizontalOverflow:        oobTarget.ShapeRange.TextFrame.HorizontalOverflow = oobSource.ShapeRange.TextFrame.HorizontalOverflow
            Case enOOBShapeRangeTextFrameMarginBottom:              oobTarget.ShapeRange.TextFrame.MarginBottom = oobSource.ShapeRange.TextFrame.MarginBottom
            Case enOOBShapeRangeTextFrameMarginLeft:                oobTarget.ShapeRange.TextFrame.MarginRight = oobSource.ShapeRange.TextFrame.MarginRight
            Case enOOBShapeRangeTextFrameMarginRight:               oobTarget.ShapeRange.TextFrame.MarginLeft = oobSource.ShapeRange.TextFrame.MarginLeft
            Case enOOBShapeRangeTextFrameMarginTop:                 oobTarget.ShapeRange.TextFrame.MarginTop = oobSource.ShapeRange.TextFrame.MarginTop
            Case enOOBShapeRangeTextFrameOrientation:               oobTarget.ShapeRange.TextFrame.Orientation = oobSource.ShapeRange.TextFrame.Orientation
            Case enOOBShapeRangeTextFrameReadingOrder:              oobTarget.ShapeRange.TextFrame.ReadingOrder = oobSource.ShapeRange.TextFrame.ReadingOrder
            Case enOOBShapeRangeTextFrameVerticalAlignment:         oobTarget.ShapeRange.TextFrame.VerticalAlignment = oobSource.ShapeRange.TextFrame.VerticalAlignment
            Case enOOBShapeRangeTextFrameVerticalOverflow:          oobTarget.ShapeRange.TextFrame.VerticalOverflow = oobSource.ShapeRange.TextFrame.VerticalOverflow
            Case enOOBShapeRangeTitle:                              oobTarget.ShapeRange.Title = oobSource.ShapeRange.TextFrame.VerticalOverflow
            Case enOOBShapeRangeTop:                                oobTarget.ShapeRange.Top = oobSource.ShapeRange.Top
            Case enOOBShapeRangeVisible:                            oobTarget.ShapeRange.Visible = oobSource.ShapeRange.Visible
            Case enOOBShapeRangeWidth:                              oobTarget.ShapeRange.Width = oobSource.ShapeRange.Width
            Case enOOBSourceName:                                   oobTarget.SourceName = oobSource.SourceName
            Case enOOBTop:                                          oobTarget.Top = oobSource.Top
            Case enOOBVisible:                                      oobTarget.Visible = oobSource.Visible
            Case enOOBWidth:                                        oobTarget.Width = oobSource.Width
        End Select
    End With
    
    With shpTarget
        Select Case sp_en_property
            Case enShapeAccelerator:                                .Accelerator = shpSource.Accelerator
            Case enShapeAlignment:                                  .Alignment = shpSource.Alignment
            Case enShapeAlternativeText:                            .AlternativeText = shpSource.AlternativeText
            Case enShapeAltHTML:                                    .AltHTML = shpSource.AltHTML
            Case enShapeAutoLoad:                                   .AutoLoad = shpSource.AutoLoad
            Case enShapeAutoShapeType:                              .AutoShapeType = shpSource.AutoShapeType
            Case enShapeAutoSize:                                   .AutoSize = shpSource.AutoSize
            Case enShapeAutoUpdate:                                 .AutoUpdate = shpSource.AutoUpdate
            Case enShapeBackColor:                                  .BackColor = shpSource.BackColor
            Case enShapeBackgroundStyle:                            .BackgroundStyle = shpSource.BackgroundStyle
            Case enShapeBackStyle:                                  .BackStyle = shpSource.BackStyle
            Case enShapeBlackWhiteMode:                             .BlackWhiteMode = shpSource.BlackWhiteMode
            Case enShapeBorderColor:                                .BorderColor = shpSource.BorderColor
            Case enShapeBorderColorIndex:                           .BorderColorIndex = shpSource.BorderColorIndex
            Case enShapeBorderLineStyle:                            .BorderLineStyle = shpSource.BorderLineStyle
            Case enShapeBorderParent:                               .BorderParent = shpSource.BorderParent
            Case enShapeBordersSuppress:                            .BordersSuppress = shpSource.BordersSuppress
            Case enShapeBorderThemeColor:                           .BorderThemeColor = shpSource.BorderThemeColor
            Case enShapeBorderTintAndShade:                         .BorderTintAndShade = shpSource.BorderTintAndShade
            Case enShapeBorderWeight:                               .BorderWeight = shpSource.BorderWeight
            Case enShapeCaption:                                    .Caption = shpSource.Caption
            Case enShapeControlFormatEnabled:                       .ControlFormat.Enabled = shpSource.ControlFormat.Enabled
            Case enShapeControlFormatDropDownLines:                 .ControlFormat.DropDownLines = shpSource.ControlFormat.DropDownLines
            Case enShapeControlFormatLargeChange:                   .ControlFormat.LargeChange = shpSource.ControlFormat.LargeChange
            Case enShapeControlFormatLinkedCell:                    .ControlFormat.LinkedCell = shpSource.ControlFormat.LinkedCell
            Case enShapeControlFormatListFillRange:                 .ControlFormat.ListFillRange = shpSource.ControlFormat.ListFillRange
            Case enShapeControlFormatListIndex:                     .ControlFormat.ListIndex = shpSource.ControlFormatListIndex
            Case enShapeControlFormatLockedText:                    .ControlFormat.LockedText = shpSource.ControlFormatLockedText
            Case enShapeControlFormatMax:                           .ControlFormat.Max = shpSource.ControlFormat.Max
            Case enShapeControlFormatMin:                           .ControlFormat.Min = shpSource.ControlFormat.Min
            Case enShapeControlFormatMultiSelect:                   .ControlFormat.MultiSelect = shpSource.ControlFormat.MultiSelect
            Case enShapeControlFormatPrintObject:                   .ControlFormat.PrintObject = shpSource.ControlFormat.PrintObject
            Case enShapeControlFormatSmallChange:                   .ControlFormat.SmallCha = shpSource.ControlFormat.SmallChange
            Case enShapeDecorative:                                 .Decorative = shpSource.Decorative
            Case enShapeEnabled:                                    .Enabled = shpSource.Enabled
            Case enShapeForeColorBrightness:                        .Fill.ForeColor.Brightness = shpSource.Shadow.ForeColor.Brightness
            Case enShapeForeColorObjectThemeColor:                  .Fill.ForeColor.ObjectThemeColor = shpSource.Shadow.ForeColor.ObjectThemeColor
            Case enShapeForeColorRGB:                               .Fill.ForeColor.RGB = shpSource.Shadow.ForeColor.RGB
            Case enShapeForeColorSchmeColor:                        .Fill.ForeColor.SchemeColor = shpSource.Shadow.ForeColor.SchemeColor
            Case enShapeForeColorTintAndShade:                      .Fill.ForeColor.TintAndShade = shpSource.Shadow.ForeColor.TintAndShade
            Case enShapeGraphicStyle:                               .GraphicStyle = shpSource.GraphicStyle
            Case enShapeGroupName:                                  .GroupName = shpSource.GroupName
            Case enShapeHeight:                                     .Height = shpSource.Height
            Case enShapeLeft:                                       .Left = shpSource.Left
            Case enShapeLinkedCell:                                 .LinkedCell = shpSource.LinkedCell
            Case enShapeListFillRange:                              .ListFillRange = shpSource.ListFillRange
            Case enShapeLockAspectRatio:                            .LockAspectRatio = shpSource.LockAspectRatio
            Case enShapeLocked:                                     .Locked = shpSource.Locked
            Case enShapeMouseIcon:                                  .MouseIcon = shpSource.MouseIcon
            Case enShapeMousePointer:                               .MousePointer = shpSource.MousePointer
            Case enShapeMultiSelect:                                .MultiSelect = shpSource.MultiSelect
            Case enShapeName:                                       .Name = shpSource.Name
            Case enShapeOnAction:                                   .OnAction = shpSource.OnAction
            Case enShapePicture:                                    .Picture = shpSource.Picture
            Case enShapePicturePosition:                            .PicturePosition = shpSource.PicturePosition
            Case enShapePlacement:                                  .Placement = shpSource.Placement
            Case enShapePrintObject:                                .PrintObject = shpSource.PrintObject
            Case enShapeRotation:                                   .Rotation = shpSource.Rotation
            Case enShapeShadowBlur:                                 .Shadow.Blur = shpSource.Shadow.Blur
            Case enShapeShadowForeColorObjectThemeColor:            .Shadow.ForeColor.ObjectThemeColor = shpSource.Shadow.ForeColor.ObjectThemeColor
'            Case enShapeShadowForeColorRGB:                         .Shadow.ForeColor.RGB = shpSource.Shadow.ForeColor.RGB
            Case enShapeShadowForeColorSchmeColor:                  .Shadow.ForeColor.SchemeColor = shpSource.Shadow.ForeColor.SchemeColor
            Case enShapeShadowForeColorTintAndShade:                .Shadow.ForeColor.TintAndShade = shpSource.Shadow.ForeColor.TintAndShade
            Case enShapeShadowObscured:                             .Shadow.Obscured = shpSource.Shadow.Obscured
            Case enShapeShadowRotateWithShape:                      .Shadow.RotateWithShape = shpSource.Shadow.RotateWithShape
            Case enShapeShadowSize:                                 .Shadow.Size = shpSource.Shadow.Size
            Case enShapeShadowStyle:                                .Shadow.Style = shpSource.Shadow.Style
            Case enShapeShadowTransparency:                         .Shadow.Transparency = shpSource.Shadow.Transparency
            Case enShapeShadowType:                                 .Shadow.Type = shpSource.Shadow.Type
            Case enShapeShadowVisible:                              .Shadow.Visible = shpSource.Shadow.Visible
            Case enShapeShapeStyle:                                 .ShapeStyle = shpSource.ShapeStyle
            Case enShapeSourceName:                                 .SourceName = shpSource.SourceName
            Case enShapeSpecialEffect:                              .SpecialEffect = shpSource.SpecialEffect
            Case enShapeTextEffectFontBold:                         .TextEffect.FontBold = shpSource.TextEffect.FontBold
            Case enShapeTextEffectFontItalic:                       .TextEffect.FontItalic = shpSource.TextEffect.FontItalic
            Case enShapeTextEffectFontName:                         .TextEffect.FontName = shpSource.TextEffect.FontName
            Case enShapeTextEffectFontSize:                         .TextEffect.FontSize = shpSource.TextEffect.FontSize
            Case enShapeTextEffectKernedPairs:                      .TextEffect.KernedPairs = shpSource.TextEffect.KernedPairs
            Case enShapeTextEffectNormalizedHeight:                 .TextEffect.NormalizedHeight = shpSource.TextEffect.NormalizedHeight
            Case enShapeTextEffectPresetShape:                      .TextEffect.PresetShape = shpSource.TextEffect.PresetShape
            Case enShapeTextEffectPresetTextEffect:                 .TextEffect.PresetTextEffect = shpSource.TextEffect.PresetTextEffect
            Case enShapeTextEffectRotatedChars:                     .TextEffect.RotatedChars = shpSource.TextEffect.RotatedChars
            Case enShapeTextEffectText:                             .TextEffect.Text = shpSource.TextEffect.Text
            Case enShapeTextFrame2AutoSize:                         .TextFrame2.AutoSize = shpSource.TextFrame2.AutoSize
            Case enShapeTextFrame2HorizontalAnchor:                 .TextFrame2.HorizontalAnchor = shpSource.TextFrame2.HorizontalAnchor
            Case enShapeTextFrame2MarginBottom:                     .TextFrame2.MarginBottom = shpSource.TextFrame2.MarginBottom
            Case enShapeTextFrame2MarginLeft:                       .TextFrame2.MarginRight = shpSource.TextFrame2.MarginRight
            Case enShapeTextFrame2MarginRight:                      .TextFrame2.MarginLeft = shpSource.TextFrame2.MarginLeft
            Case enShapeTextFrame2MarginTop:                        .TextFrame2.MarginTop = shpSource.TextFrame2.MarginTop
            Case enShapeTextFrame2NoTextRotation:                   .TextFrame2.NoTextRotation = shpSource.TextFrame2.NoTextRotation
            Case enShapeTextFrame2Orientation:                      .TextFrame2.Orientation = shpSource.TextFrame2.Orientation
            Case enShapeTextFrame2PathFormat:                       .TextFrame2.PathFormat = shpSource.TextFrame2.PathFormat
            Case enShapeTextFrame2WarpFormat:                       .TextFrame2.WarpFormat = shpSource.TextFrame2.WarpFormat
            Case enShapeTextFrame2WordArtformat:                    .TextFrame2.WordArtformat = shpSource.TextFrame2.WordArtformat
            Case enShapeTextFrameAutoMargins:                       .TextFrame.AutoMargins = shpSource.TextFrame.AutoMargins
            Case enShapeTextFrameAutoSize:                          .TextFrame.AutoSize = shpSource.TextFrame.AutoSize
            Case enShapeTextFrameCharactersCaption:                 .TextFrame.Characters.Caption = shpSource.TextFrame.Characters.Caption
            Case enShapeTextFrameCharactersText:                    .TextFrame.Characters.Text = shpSource.TextFrame.Characters.Text
            Case enShapeTextFrameHorizontalAlignment:               .TextFrame.HorizontalAlignment = shpSource.TextFrame.HorizontalAlignment
            Case enShapeTextFrameHorizontalOverflow:                .TextFrame.HorizontalOverflow = shpSource.TextFrame.HorizontalOverflow
            Case enShapeTextFrameMarginBottom:                      .TextFrame.MarginBottom = shpSource.TextFrame.MarginBottom
            Case enShapeTextFrameMarginLeft:                        .TextFrame.MarginRight = shpSource.TextFrame.MarginRight
            Case enShapeTextFrameMarginRight:                       .TextFrame.MarginLeft = shpSource.TextFrame.MarginLeft
            Case enShapeTextFrameMarginTop:                         .TextFrame.MarginTop = shpSource.TextFrame.MarginTop
            Case enShapeTextFrameOrientation:                       .TextFrame.Orientation = shpSource.TextFrame.Orientation
            Case enShapeTextFrameReadingOrder:                      .TextFrame.ReadingOrder = shpSource.TextFrame.ReadingOrder
            Case enShapeTextFrameVerticalAlignment:                 .TextFrame.VerticalAlignment = shpSource.TextFrame.VerticalAlignment
            Case enShapeTextFrameVerticalOverflow:                  .TextFrame.VerticalOverflow = shpSource.TextFrame.VerticalOverflow
            Case enShapeTitle:                                      .Title = shpSource.Title
            Case enShapeTop:                                        .Top = shpSource.Top
            Case enShapeVisible:                                    .Visible = shpSource.Visible
            Case enShapeWidth:                                      .Width = shpSource.Width
            Case enShapeWordWrap:                                   .WordWrap = shpSource.WordWrap
        End Select
    End With
    DoEvents
    
End Sub

Private Function SyncPropertyAsserted(ByVal sp_enProperty As enProperties) As Boolean
' ------------------------------------------------------------------------------
' Synchronizes the Properties vf the shape (shpSource) with the corresponding
' shape in the corresponding target sheet in the Sync-Target-Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "SyncProperty"
    
    If Not oobTarget Is Nothing And Not oobSource Is Nothing Then
        Select Case sp_enProperty
            Case enOOBBorderColor:                                  SyncPropertyAsserted = oobTarget.Border.Color = oobSource.Border.Color
            Case enOOBBorderColorIndex:                             SyncPropertyAsserted = oobTarget.Border.ColorIndex = oobSource.Border.ColorIndex
            Case enOOBBorderLineStyle:                              SyncPropertyAsserted = oobTarget.Border.LineStyle = oobSource.Border.LineStyle
            Case enOOBLinkedCell:                                   SyncPropertyAsserted = oobTarget.LinkedCell = oobSource.LinkedCell
            Case enOOBName:                                         SyncPropertyAsserted = oobTarget.Name = oobSource.Name
            Case enOOBAutoLoad:                                     SyncPropertyAsserted = oobTarget.AutoLoad = oobSource.AutoLoad
            Case enOOBObjectCaption:                                SyncPropertyAsserted = oobTarget.Object.Caption = oobSource.Object.Caption
            Case enOOBObjectEnabled:                                SyncPropertyAsserted = oobTarget.Object.Enabled = oobSource.Object.Enabled
            Case enOOBObjectForeColor:                              SyncPropertyAsserted = oobTarget.Object.ForeColor = oobSource.Object.ForeColor
            Case enOOBObjectBackColor:                              SyncPropertyAsserted = oobTarget.Object.BackColor = oobSource.Object.BackColor
            Case enOOBObjectBackStyle:                              SyncPropertyAsserted = oobTarget.Object.BackStyle = oobSource.Object.BackStyle
            Case enOOBObjectFontBold:                               SyncPropertyAsserted = oobTarget.Object.Font.Bold = oobSource.Object.Font.Bold
            Case enOOBObjectFontItalic:                             SyncPropertyAsserted = oobTarget.Object.Font.Italic = oobSource.Object.Font.Italic
            Case enOOBObjectFontName:                               SyncPropertyAsserted = oobTarget.Object.Font.Name = oobSource.Object.Font.Name
            Case enOOBObjectFontSize:                               SyncPropertyAsserted = oobTarget.Object.Font.Size = oobSource.Object.Font.Size
            Case enOOBObjectLocked:                                 SyncPropertyAsserted = oobTarget.Object.Locked = oobSource.Object.Locked
            Case enOOBObjectMouseIcon:                              SyncPropertyAsserted = oobTarget.Object.MouseIcon = oobSource.Object.MouseIcon
            Case enOOBObjectMousePointer:                           SyncPropertyAsserted = oobTarget.Object.MousePointer = oobSource.Object.MousePointer
            Case enOOBObjectPicture:                                SyncPropertyAsserted = oobTarget.Object.Picture = oobSource.Object.Picture
            Case enOOBObjectPictureposition:                        SyncPropertyAsserted = oobTarget.Object.PicturePosition = oobSource.Object.PicturePosition
            Case enOOBObjectTakeFocusOnClick:                       SyncPropertyAsserted = oobTarget.Object.TakeFocusOnClick = oobSource.Object.TakeFocusOnClick
            Case enOOBObjectWordWrap:                               SyncPropertyAsserted = oobTarget.Object.WordWrap = oobSource.Object.WordWrap
            Case enOOBHeight:                                       SyncPropertyAsserted = oobTarget.Height = oobSource.Height
            Case enOOBLeft:                                         SyncPropertyAsserted = oobTarget.Left = oobSource.Left
            Case enOOBPlacement:                                    SyncPropertyAsserted = oobTarget.Placement = oobSource.Placement
            Case enOOBPrintObject:                                  SyncPropertyAsserted = oobTarget.PrintObject = oobSource.PrintObject
            Case enOOBShadow:                                       SyncPropertyAsserted = oobTarget.Shadow = oobSource.Shadow
            Case enOOBTop:                                          SyncPropertyAsserted = oobTarget.Top = oobSource.Top
            Case enOOBVisible:                                      SyncPropertyAsserted = oobTarget.Visible = oobSource.Visible
            Case enOOBWidth:                                        SyncPropertyAsserted = oobTarget.Width = oobSource.Width
            Case enOOBOLEType:                                      SyncPropertyAsserted = oobTarget.OLEType = oobSource.OLEType
            Case enOOBShapeRangeAlternativeText:                    SyncPropertyAsserted = oobTarget.ShapeRange.AlternativeText = oobSource.ShapeRange.AlternativeText
            Case enOOBShapeRangeAutoShapeType:                      SyncPropertyAsserted = oobTarget.ShapeRange.AutoShapeType = oobSource.ShapeRange.AutoShapeType
            Case enOOBShapeRangeBackgroundStyle:                    SyncPropertyAsserted = oobTarget.ShapeRange.BackgroundStyle = oobSource.ShapeRange.BackgroundStyle
            Case enOOBShapeRangeBlackWhiteMode:                     SyncPropertyAsserted = oobTarget.ShapeRange.BlackWhiteMode = oobSource.ShapeRange.BlackWhiteMode
            Case enOOBShapeRangeDecorative:                         SyncPropertyAsserted = oobTarget.ShapeRange.Decorative = oobSource.ShapeRange.Decorative
            Case enOOBShapeRangeGraphicStyle:                       SyncPropertyAsserted = oobTarget.ShapeRange.GraphicStyle = oobSource.ShapeRange.GraphicStyle
            Case enOOBShapeRangeHeight:                             SyncPropertyAsserted = oobTarget.ShapeRange.Height = oobSource.ShapeRange.Height
            Case enOOBShapeRangeLeft:                               SyncPropertyAsserted = oobTarget.ShapeRange.Left = oobSource.ShapeRange.Left
            Case enOOBShapeRangeLineBackColor:                      SyncPropertyAsserted = oobTarget.ShapeRange.Line.BackColor = oobSource.ShapeRange.Line.BackColor
            Case enOOBShapeRangeLineBeginArrowheadLength:           SyncPropertyAsserted = oobTarget.ShapeRange.Line.BeginArrowheadLength = oobSource.ShapeRange.Line.BeginArrowheadLength
            Case enOOBShapeRangeLineBeginArrowheadStyle:            SyncPropertyAsserted = oobTarget.ShapeRange.Line.BeginArrowheadStyle = oobSource.ShapeRange.Line.BeginArrowheadStyle
            Case enOOBShapeRangeLineBeginArrowheadWidth:            SyncPropertyAsserted = oobTarget.ShapeRange.Line.BeginArrowheadWidth = oobSource.ShapeRange.Line.BeginArrowheadWidth
            Case enOOBShapeRangeLineDashStyle:                      SyncPropertyAsserted = oobTarget.ShapeRange.Line.DashStyle = oobSource.ShapeRange.Line.DashStyle
            Case enOOBShapeRangeLineEndArrowheadLength:             SyncPropertyAsserted = oobTarget.ShapeRange.Line.EndArrowheadLength = oobSource.ShapeRange.Line.EndArrowheadLength
            Case enOOBShapeRangeLineEndArrowheadStyle:              SyncPropertyAsserted = oobTarget.ShapeRange.Line.EndArrowheadStyle = oobSource.ShapeRange.Line.EndArrowheadStyle
            Case enOOBShapeRangeLineEndArrowheadWidth:              SyncPropertyAsserted = oobTarget.ShapeRange.Line.EndArrowheadWidth = oobSource.ShapeRange.Line.EndArrowheadWidth
            Case enOOBShapeRangeLineForeColorBrightness:            SyncPropertyAsserted = oobTarget.ShapeRange.Line.ForeColor.Brightness = oobSource.ShapeRange.Line.ForeColor.Brightness
            Case enOOBShapeRangeLineForeColorObjectThemeColor:      SyncPropertyAsserted = oobTarget.ShapeRange.Line.ForeColor.ObjectThemeColor = oobSource.ShapeRange.Line.ForeColor.ObjectThemeColor
            Case enOOBShapeRangeLineForeColorRGB:                   SyncPropertyAsserted = oobTarget.ShapeRange.Line.ForeColor.RGB = oobSource.ShapeRange.Line.ForeColor.RGB
            Case enOOBShapeRangeLineForeColorSchemeColor:           SyncPropertyAsserted = oobTarget.ShapeRange.Line.ForeColor.SchemeColor = oobSource.ShapeRange.Line.ForeColor.SchemeColor
            Case enOOBShapeRangeLineForeColorTintAndShade:          SyncPropertyAsserted = oobTarget.ShapeRange.Line.ForeColor.TintAndShade = oobSource.ShapeRange.Line.ForeColor.TintAndShade
            Case enOOBShapeRangeLineInsetPen:                       SyncPropertyAsserted = oobTarget.ShapeRange.Line.InsetPen = oobSource.ShapeRange.Line.InsetPen
            Case enOOBShapeRangeLinePattern:                        SyncPropertyAsserted = oobTarget.ShapeRange.Line.Pattern = oobSource.ShapeRange.Line.Pattern
            Case enOOBShapeRangeLineStyle:                          SyncPropertyAsserted = oobTarget.ShapeRange.Line.Style = oobSource.ShapeRange.Line.Style
            Case enOOBShapeRangeLineTransparency:                   SyncPropertyAsserted = oobTarget.ShapeRange.Line.Transparency = oobSource.ShapeRange.Line.Transparency
            Case enOOBShapeRangeLineVisible:                        SyncPropertyAsserted = oobTarget.ShapeRange.Line.Visible = oobSource.ShapeRange.Line.Visible
            Case enOOBShapeRangeLineWeight:                         SyncPropertyAsserted = oobTarget.ShapeRange.Line.Weight = oobSource.ShapeRange.Line.Weight
            Case enOOBShapeRangeLockAspectRatio:                    SyncPropertyAsserted = oobTarget.ShapeRange.LockAspectRatio = oobSource.ShapeRange.LockAspectRatio
            Case enOOBShapeRangeName:                               SyncPropertyAsserted = oobTarget.ShapeRange.Name = oobSource.ShapeRange.Name
            Case enOOBShapeRangePictureFormatBrightness:            SyncPropertyAsserted = oobTarget.ShapeRange.PictureFormat.Brightness = oobSource.ShapeRange.Picture.Brightness
            Case enOOBShapeRangePictureFormatColorType:             SyncPropertyAsserted = oobTarget.ShapeRange.PictureFormat.ColorType = oobSource.ShapeRange.PictureFormat.ColorType
            Case enOOBShapeRangePictureFormatContrast:              SyncPropertyAsserted = oobTarget.ShapeRange.PictureFormat.Contrast = oobSource.ShapeRange.PictureFormat.Contrast
            Case enOOBShapeRangePictureFormatTransparencyColor:     SyncPropertyAsserted = oobTarget.ShapeRange.PictureFormat.TransparencyColor = oobSource.ShapeRange.PictureFormat.TransparencyColor
            Case enOOBShapeRangePictureFormatTransparentBackground: SyncPropertyAsserted = oobTarget.ShapeRange.PictureFormat.TransparentBackground = oobSource.ShapeRange.PictureFormat.TransparentBackground
            Case enOOBShapeRangeRotation:                           SyncPropertyAsserted = oobTarget.ShapeRange.Rotation = oobSource.ShapeRange.Rotation
            Case enOOBShapeRangeShadowBlur:                         SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Blur = oobSource.ShapeRange.Shadow.Blur
            Case enOOBShapeRangeShadowForeColorBrightness:          SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.ForeColor.Brightness = oobSource.ShapeRange.Shadow.ForeColor.Brightness
            Case enOOBShapeRangeShadowForeColorObjectThemeColor:    SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.ForeColor.ObjectThemeColor = oobSource.ShapeRange.Shadow.ForeColor.ObjectThemeColor
            Case enOOBShapeRangeShadowForeColorRGB:                 SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.ForeColor.RGB = oobSource.ShapeRange.Shadow.ForeColor.RGB
            Case enOOBShapeRangeShadowForeColorSchemeColor:         SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.ForeColor.SchemeColor = oobSource.ShapeRange.Shadow.ForeColor.SchemeColor
            Case enOOBShapeRangeShadowForeColorTintAndShade:        SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.ForeColor.TintAndShade = oobSource.ShapeRange.Shadow.ForeColor.TintAndShade
            Case enOOBShapeRangeShadowObscured:                     SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Obscured = oobSource.ShapeRange.Shadow.Obscured
            Case enOOBShapeRangeShadowOffsetX:                      SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.OffsetX = oobSource.ShapeRange.Shadow.OffsetX
            Case enOOBShapeRangeShadowOffsetY:                      SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.OffsetY = oobSource.ShapeRange.Shadow.OffsetY
            Case enOOBShapeRangeShadowRotateWithShape:              SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.RotateWithShape = oobSource.ShapeRange.Shadow.RotateWithShape
            Case enOOBShapeRangeShadowSize:                         SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Size = oobSource.ShapeRange.Shadow.Size
            Case enOOBShapeRangeShadowStyle:                        SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Style = oobSource.ShapeRange.Shadow.Style
            Case enOOBShapeRangeShadowTransparency:                 SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Transparency = oobSource.ShapeRange.Shadow.Transparency
            Case enOOBShapeRangeShadowType:                         SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Type = oobSource.ShapeRange.Shadow.Type
            Case enOOBShapeRangeShadowVisible:                      SyncPropertyAsserted = oobTarget.ShapeRange.Shadow.Visible = oobSource.ShapeRange.Shadow.Visible
            Case enOOBShapeRangeShapeStyle:                         SyncPropertyAsserted = oobTarget.ShapeRange.ShapeStyle = oobSource.ShapeRange.ShapeStyle
            Case enOOBShapeRangeSoftEdgeRadius:                     SyncPropertyAsserted = oobTarget.ShapeRange.SoftEdge.Radius = oobSource.ShapeRange.SoftEdge.Radius
            Case enOOBShapeRangeSoftEdgeType:                       SyncPropertyAsserted = oobTarget.ShapeRange.SoftEdge.Type = oobSource.ShapeRange.SoftEdge.Type
            Case enOOBShapeRangeTextEffectFontBold:                 SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.FontBold = oobSource.ShapeRange.TextEffect.FontBold
            Case enOOBShapeRangeTextEffectFontItalic:               SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.FontItalic = oobSource.ShapeRange.TextEffect.FontItalic
            Case enOOBShapeRangeTextEffectFontName:                 SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.FontName = oobSource.ShapeRange.TextEffect.FontName
            Case enOOBShapeRangeTextEffectFontSize:                 SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.FontSize = oobSource.ShapeRange.TextEffect.FontSize
            Case enOOBShapeRangeTextEffectKernedPairs:              SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.KernedPairs = oobSource.ShapeRange.TextEffect.KernedPairs
            Case enOOBShapeRangeTextEffectNormalizedHeight:         SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.NormalizedHeight = oobSource.ShapeRange.TextEffect.NormalizedHeight
            Case enOOBShapeRangeTextEffectPresetShape:              SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.PresetShape = oobSource.ShapeRange.TextEffect.PresetShape
            Case enOOBShapeRangeTextEffectPresetTextEffect:         SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.PresetTextEffect = oobSource.ShapeRange.TextEffect.PresetTextEffect
            Case enOOBShapeRangeTextEffectRotatedChars:             SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.RotatedChars = oobSource.ShapeRange.TextEffect.RotatedChars
            Case enOOBShapeRangeTextEffectText:                     SyncPropertyAsserted = oobTarget.ShapeRange.TextEffect.Text = oobSource.ShapeRange.TextEffect.Text
            Case enOOBShapeRangeTextFrame2AutoSize:                 SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.AutoSize = oobSource.ShapeRange.TextFrame2.AutoSize
            Case enOOBShapeRangeTextFrame2HorizontalAnchor:         SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.HorizontalAnchor = oobSource.ShapeRange.TextFrame2.HorizontalAnchor
            Case enOOBShapeRangeTextFrame2MarginBottom:             SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.MarginBottom = oobSource.ShapeRange.TextFrame2.MarginBottom
            Case enOOBShapeRangeTextFrame2MarginLeft:               SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.MarginRight = oobSource.ShapeRange.TextFrame2.MarginRight
            Case enOOBShapeRangeTextFrame2MarginRight:              SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.MarginLeft = oobSource.ShapeRange.TextFrame2.MarginLeft
            Case enOOBShapeRangeTextFrame2MarginTop:                SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.MarginTop = oobSource.ShapeRange.TextFrame2.MarginTop
            Case enOOBShapeRangeTextFrame2NoTextRotation:           SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.NoTextRotation = oobSource.ShapeRange.TextFrame2.NoTextRotation
            Case enOOBShapeRangeTextFrame2Orientation:              SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.Orientation = oobSource.ShapeRange.TextFrame2.Orientation
            Case enOOBShapeRangeTextFrame2PathFormat:               SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.PathFormat = oobSource.ShapeRange.TextFrame2.PathFormat
            Case enOOBShapeRangeTextFrame2WarpFormat:               SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.WarpFormat = oobSource.ShapeRange.TextFrame2.WarpFormat
            Case enOOBShapeRangeTextFrame2WordArtformat:            SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame2.WordArtformat = oobSource.ShapeRange.TextFrame2.WordArtformat
            Case enOOBShapeRangeTextFrameAutoMargins:               SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.AutoMargins = oobSource.ShapeRange.TextFrame.AutoMargins
            Case enOOBShapeRangeTextFrameAutoSize:                  SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.AutoSize = oobSource.ShapeRange.TextFrame.AutoSize
            Case enOOBShapeRangeTextFrameCharactersCaption:         SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.Characters.Caption = oobSource.ShapeRange.TextFrame.Characters.Caption
            Case enOOBShapeRangeTextFrameHorizontalAlignment:       SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.HorizontalAlignment = oobSource.ShapeRange.TextFrame.HorizontalAlignment
            Case enOOBShapeRangeTextFrameHorizontalOverflow:        SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.HorizontalOverflow = oobSource.ShapeRange.TextFrame.HorizontalOverflow
            Case enOOBShapeRangeTextFrameMarginBottom:              SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.MarginBottom = oobSource.ShapeRange.TextFrame.MarginBottom
            Case enOOBShapeRangeTextFrameMarginLeft:                SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.MarginRight = oobSource.ShapeRange.TextFrame.MarginRight
            Case enOOBShapeRangeTextFrameMarginRight:               SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.MarginLeft = oobSource.ShapeRange.TextFrame.MarginLeft
            Case enOOBShapeRangeTextFrameMarginTop:                 SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.MarginTop = oobSource.ShapeRange.TextFrame.MarginTop
            Case enOOBShapeRangeTextFrameOrientation:               SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.Orientation = oobSource.ShapeRange.TextFrame.Orientation
            Case enOOBShapeRangeTextFrameReadingOrder:              SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.ReadingOrder = oobSource.ShapeRange.TextFrame.ReadingOrder
            Case enOOBShapeRangeTextFrameVerticalAlignment:         SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.VerticalAlignment = oobSource.ShapeRange.TextFrame.VerticalAlignment
            Case enOOBShapeRangeTextFrameVerticalOverflow:          SyncPropertyAsserted = oobTarget.ShapeRange.TextFrame.VerticalOverflow = oobSource.ShapeRange.TextFrame.VerticalOverflow
            Case enOOBShapeRangeTitle:                              SyncPropertyAsserted = oobTarget.ShapeRange.Title = oobSource.ShapeRange.TextFrame.VerticalOverflow
            Case enOOBShapeRangeTop:                                SyncPropertyAsserted = oobTarget.ShapeRange.Top = oobSource.ShapeRange.Top
            Case enOOBShapeRangeVisible:                            SyncPropertyAsserted = oobTarget.ShapeRange.Visible = oobSource.ShapeRange.Visible
            Case enOOBShapeRangeWidth:                              SyncPropertyAsserted = oobTarget.ShapeRange.Width = oobSource.ShapeRange.Width
            Case enOOBSourceName:                                   SyncPropertyAsserted = oobTarget.SourceName = oobSource.SourceName
        End Select
    End If
    If Not shpTarget Is Nothing And Not shpSource Is Nothing Then
        Select Case sp_enProperty
            Case enShapeAccelerator:                                SyncPropertyAsserted = shpTarget.Accelerator = shpSource.Accelerator
            Case enShapeAlignment:                                  SyncPropertyAsserted = shpTarget.Alignment = shpSource.Alignment
            Case enShapeAlternativeText:                            SyncPropertyAsserted = shpTarget.AlternativeText = shpSource.AlternativeText
            Case enShapeAltHTML:                                    SyncPropertyAsserted = shpTarget.AltHTML = shpSource.AltHTML
            Case enShapeAutoLoad:                                   SyncPropertyAsserted = shpTarget.AutoLoad = shpSource.AutoLoad
            Case enShapeAutoShapeType:                              SyncPropertyAsserted = shpTarget.AutoShapeType = shpSource.AutoShapeType
            Case enShapeAutoSize:                                   SyncPropertyAsserted = shpTarget.AutoSize = shpSource.AutoSize
            Case enShapeAutoUpdate:                                 SyncPropertyAsserted = shpTarget.AutoUpdate = shpSource.AutoUpdate
            Case enShapeBackColor:                                  SyncPropertyAsserted = shpTarget.BackColor = shpSource.BackColor
            Case enShapeBackgroundStyle:                            SyncPropertyAsserted = shpTarget.BackgroundStyle = shpSource.BackgroundStyle
            Case enShapeBackStyle:                                  SyncPropertyAsserted = shpTarget.BackStyle = shpSource.BackStyle
            Case enShapeBlackWhiteMode:                             SyncPropertyAsserted = shpTarget.BlackWhiteMode = shpSource.BlackWhiteMode
            Case enShapeBorderColor:                                SyncPropertyAsserted = shpTarget.BorderColor = shpSource.BorderColor
            Case enShapeBorderColorIndex:                           SyncPropertyAsserted = shpTarget.BorderColorIndex = shpSource.BorderColorIndex
            Case enShapeBorderLineStyle:                            SyncPropertyAsserted = shpTarget.BorderLineStyle = shpSource.BorderLineStyle
            Case enShapeBorderParent:                               SyncPropertyAsserted = shpTarget.BorderParent = shpSource.BorderParent
            Case enShapeBordersSuppress:                            SyncPropertyAsserted = shpTarget.BordersSuppress = shpSource.BordersSuppress
            Case enShapeBorderThemeColor:                           SyncPropertyAsserted = shpTarget.BorderThemeColor = shpSource.BorderThemeColor
            Case enShapeBorderTintAndShade:                         SyncPropertyAsserted = shpTarget.BorderTintAndShade = shpSource.BorderTintAndShade
            Case enShapeBorderWeight:                               SyncPropertyAsserted = shpTarget.BorderWeight = shpSource.BorderWeight
            Case enShapeCaption:                                    SyncPropertyAsserted = shpTarget.Caption = shpSource.Caption
            Case enShapeDecorative:                                 SyncPropertyAsserted = shpTarget.Decorative = shpSource.Decorative
            Case enShapeEnabled:                                    SyncPropertyAsserted = shpTarget.Enabled = shpSource.Enabled
            Case enShapeForeColorBrightness:                        SyncPropertyAsserted = shpTarget.ForeColor.Brightness = shpSource.ForeColor.Brightness
            Case enShapeForeColorObjectThemeColor:                  SyncPropertyAsserted = shpTarget.ForeColor.ObjectThemeColor = shpSource.ForeColor.ObjectThemeColor
            Case enShapeForeColorRGB:                               SyncPropertyAsserted = shpTarget.ForeColor.RGB = shpSource.ForeColor.RGB
            Case enShapeForeColorSchmeColor:                        SyncPropertyAsserted = shpTarget.ForeColor.SchemeColor = shpSource.ForeColor.SchemeColor
            Case enShapeForeColorTintAndShade:                      SyncPropertyAsserted = shpTarget.ForeColor.TintAndShade = shpSource.ForeColor.TintAndShade
            Case enShapeGraphicStyle:                               SyncPropertyAsserted = shpTarget.GraphicStyle = shpSource.GraphicStyle
            Case enShapeGroupName:                                  SyncPropertyAsserted = shpTarget.GroupName = shpSource.GroupName
            Case enShapeHeight:                                     SyncPropertyAsserted = shpTarget.Height = Round(shpSource.Height, 0)
            Case enShapeLeft:                                       SyncPropertyAsserted = shpTarget.Left = shpSource.Left
            Case enShapeLinkedCell:                                 SyncPropertyAsserted = shpTarget.LinkedCell = shpSource.LinkedCell
            Case enShapeListFillRange:                              SyncPropertyAsserted = shpTarget.ListFillRange = shpSource.ListFillRange
            Case enShapeLockAspectRatio:                            SyncPropertyAsserted = shpTarget.LockAspectRatio = shpSource.LockAspectRatio
            Case enShapeLocked:                                     SyncPropertyAsserted = shpTarget.Locked = shpSource.Locked
            Case enShapeMouseIcon:                                  SyncPropertyAsserted = shpTarget.MouseIcon = shpSource.MouseIcon
            Case enShapeMousePointer:                               SyncPropertyAsserted = shpTarget.MousePointer = shpSource.MousePointer
            Case enShapeMultiSelect:                                SyncPropertyAsserted = shpTarget.MultiSelect = shpSource.MultiSelect
            Case enShapeName:                                       SyncPropertyAsserted = shpTarget.Name = shpSource.Name
            Case enShapeOnAction:                                   SyncPropertyAsserted = shpTarget.OnAction = shpSource.OnAction
            Case enShapePicture:                                    SyncPropertyAsserted = shpTarget.Picture = shpSource.Picture
            Case enShapePicturePosition:                            SyncPropertyAsserted = shpTarget.PicturePosition = shpSource.PicturePosition
            Case enShapePlacement:                                  SyncPropertyAsserted = shpTarget.Placement = shpSource.Placement
            Case enShapePrintObject:                                SyncPropertyAsserted = shpTarget.PrintObject = shpSource.PrintObject
            Case enShapeRotation:                                   SyncPropertyAsserted = shpTarget.Rotation = shpSource.Rotation
            Case enShapeShadowBlur:                                 SyncPropertyAsserted = shpTarget.Shadow.Blur = shpSource.Shadow.Blur
            Case enShapeShadowForeColorObjectThemeColor:            SyncPropertyAsserted = shpTarget.Shadow.ForeColor.ObjectThemeColor = shpSource.Shadow.ForeColor.ObjectThemeColor
'            Case enShapeShadowForeColorRGB:                         SyncPropertyAsserted = shpTarget.Shadow.ForeColor.RGB = shpSource.Shadow.ForeColor.RGB
            Case enShapeShadowForeColorSchmeColor:                  SyncPropertyAsserted = shpTarget.Shadow.ForeColor.SchemeColor = shpSource.Shadow.ForeColor.SchemeColor
            Case enShapeShadowForeColorTintAndShade:                SyncPropertyAsserted = shpTarget.Shadow.ForeColor.TintAndShade = shpSource.Shadow.ForeColor.TintAndShade
            Case enShapeShadowObscured:                             SyncPropertyAsserted = shpTarget.Shadow.Obscured = shpSource.Shadow.Obscured
            Case enShapeShadowRotateWithShape:                      SyncPropertyAsserted = shpTarget.Shadow.RotateWithShape = shpSource.Shadow.RotateWithShape
            Case enShapeShadowSize:                                 SyncPropertyAsserted = shpTarget.Shadow.Size = shpSource.Shadow.Size
            Case enShapeShadowStyle:                                SyncPropertyAsserted = shpTarget.Shadow.Style = shpSource.Shadow.Style
            Case enShapeShadowTransparency:                         SyncPropertyAsserted = shpTarget.Shadow.Transparency = shpSource.Shadow.Transparency
            Case enShapeShadowType:                                 SyncPropertyAsserted = shpTarget.Shadow.Type = shpSource.Shadow.Type
            Case enShapeShadowVisible:                              SyncPropertyAsserted = shpTarget.Shadow.Visible = shpSource.Shadow.Visible
            Case enShapeShapeStyle:                                 SyncPropertyAsserted = shpTarget.ShapeStyle = shpSource.ShapeStyle
            Case enShapeSourceName:                                 SyncPropertyAsserted = shpTarget.SourceName = shpSource.SourceName
            Case enShapeSpecialEffect:                              SyncPropertyAsserted = shpTarget.SpecialEffect = shpSource.SpecialEffect
            Case enShapeTextEffectFontBold:                         SyncPropertyAsserted = shpTarget.TextEffect.FontBold = shpSource.TextEffect.FontBold
            Case enShapeTextEffectFontItalic:                       SyncPropertyAsserted = shpTarget.TextEffect.FontItalic = shpSource.TextEffect.FontItalic
            Case enShapeTextEffectFontName:                         SyncPropertyAsserted = shpTarget.TextEffect.FontName = shpSource.TextEffect.FontName
            Case enShapeTextEffectFontSize:                         SyncPropertyAsserted = shpTarget.TextEffect.FontSize = shpSource.TextEffect.FontSize
            Case enShapeTextEffectKernedPairs:                      SyncPropertyAsserted = shpTarget.TextEffect.KernedPairs = shpSource.TextEffect.KernedPairs
            Case enShapeTextEffectNormalizedHeight:                 SyncPropertyAsserted = shpTarget.TextEffect.NormalizedHeight = shpSource.TextEffect.NormalizedHeight
            Case enShapeTextEffectPresetShape:                      SyncPropertyAsserted = shpTarget.TextEffect.PresetShape = shpSource.TextEffect.PresetShape
            Case enShapeTextEffectPresetTextEffect:                 SyncPropertyAsserted = shpTarget.TextEffect.PresetTextEffect = shpSource.TextEffect.PresetTextEffect
            Case enShapeTextEffectRotatedChars:                     SyncPropertyAsserted = shpTarget.TextEffect.RotatedChars = shpSource.TextEffect.RotatedChars
            Case enShapeTextEffectText:                             SyncPropertyAsserted = shpTarget.TextEffect.Text = shpSource.TextEffect.Text
            Case enShapeTextFrame2AutoSize:                         SyncPropertyAsserted = shpTarget.TextFrame2.AutoSize = shpSource.TextFrame2.AutoSize
            Case enShapeTextFrame2HorizontalAnchor:                 SyncPropertyAsserted = shpTarget.TextFrame2.HorizontalAnchor = shpSource.TextFrame2.HorizontalAnchor
            Case enShapeTextFrame2MarginBottom:                     SyncPropertyAsserted = shpTarget.TextFrame2.MarginBottom = shpSource.TextFrame2.MarginBottom
            Case enShapeTextFrame2MarginLeft:                       SyncPropertyAsserted = shpTarget.TextFrame2.MarginRight = shpSource.TextFrame2.MarginRight
            Case enShapeTextFrame2MarginRight:                      SyncPropertyAsserted = shpTarget.TextFrame2.MarginLeft = shpSource.TextFrame2.MarginLeft
            Case enShapeTextFrame2MarginTop:                        SyncPropertyAsserted = shpTarget.TextFrame2.MarginTop = shpSource.TextFrame2.MarginTop
            Case enShapeTextFrame2NoTextRotation:                   SyncPropertyAsserted = shpTarget.TextFrame2.NoTextRotation = shpSource.TextFrame2.NoTextRotation
            Case enShapeTextFrame2Orientation:                      SyncPropertyAsserted = shpTarget.TextFrame2.Orientation = shpSource.TextFrame2.Orientation
            Case enShapeTextFrame2PathFormat:                       SyncPropertyAsserted = shpTarget.TextFrame2.PathFormat = shpSource.TextFrame2.PathFormat
            Case enShapeTextFrame2WarpFormat:                       SyncPropertyAsserted = shpTarget.TextFrame2.WarpFormat = shpSource.TextFrame2.WarpFormat
            Case enShapeTextFrame2WordArtformat:                    SyncPropertyAsserted = shpTarget.TextFrame2.WordArtformat = shpSource.TextFrame2.WordArtformat
            Case enShapeTextFrameAutoMargins:                       SyncPropertyAsserted = shpTarget.TextFrame.AutoMargins = shpSource.TextFrame.AutoMargins
            Case enShapeTextFrameAutoSize:                          SyncPropertyAsserted = shpTarget.TextFrame.AutoSize = shpSource.TextFrame.AutoSize
            Case enShapeTextFrameCharactersCaption:                 SyncPropertyAsserted = shpTarget.TextFrame.Characters.Caption = shpSource.TextFrame.Characters.Caption
            Case enShapeTextFrameCharactersText:                    SyncPropertyAsserted = shpTarget.TextFrame.Characters.Text = shpSource.TextFrame.Characters.Text
            Case enShapeTextFrameHorizontalAlignment:               SyncPropertyAsserted = shpTarget.TextFrame.HorizontalAlignment = shpSource.TextFrame.HorizontalAlignment
            Case enShapeTextFrameHorizontalOverflow:                SyncPropertyAsserted = shpTarget.TextFrame.HorizontalOverflow = shpSource.TextFrame.HorizontalOverflow
            Case enShapeTextFrameMarginBottom:                      SyncPropertyAsserted = shpTarget.TextFrame.MarginBottom = shpSource.TextFrame.MarginBottom
            Case enShapeTextFrameMarginLeft:                        SyncPropertyAsserted = shpTarget.TextFrame.MarginRight = shpSource.TextFrame.MarginRight
            Case enShapeTextFrameMarginRight:                       SyncPropertyAsserted = shpTarget.TextFrame.MarginLeft = shpSource.TextFrame.MarginLeft
            Case enShapeTextFrameMarginTop:                         SyncPropertyAsserted = shpTarget.TextFrame.MarginTop = shpSource.TextFrame.MarginTop
            Case enShapeTextFrameOrientation:                       SyncPropertyAsserted = shpTarget.TextFrame.Orientation = shpSource.TextFrame.Orientation
            Case enShapeTextFrameReadingOrder:                      SyncPropertyAsserted = shpTarget.TextFrame.ReadingOrder = shpSource.TextFrame.ReadingOrder
            Case enShapeTextFrameVerticalAlignment:                 SyncPropertyAsserted = shpTarget.TextFrame.VerticalAlignment = shpSource.TextFrame.VerticalAlignment
            Case enShapeTextFrameVerticalOverflow:                  SyncPropertyAsserted = shpTarget.TextFrame.VerticalOverflow = shpSource.TextFrame.VerticalOverflow
            Case enShapeTitle:                                      SyncPropertyAsserted = shpTarget.Title = shpSource.Title
            Case enShapeTop:                                        SyncPropertyAsserted = shpTarget.Top = shpSource.Top
            Case enShapeVisible:                                    SyncPropertyAsserted = shpTarget.Visible = shpSource.Visible
            Case enShapeWidth:                                      SyncPropertyAsserted = shpTarget.Width = shpSource.Width
            Case enShapeWordWrap:                                   SyncPropertyAsserted = shpTarget.WordWrap = shpSource.WordWrap
        End Select
    End If
    
xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

