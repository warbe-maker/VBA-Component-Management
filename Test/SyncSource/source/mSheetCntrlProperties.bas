Attribute VB_Name = "mSheetCntrlProperties"
Option Explicit

Private sCntrlType As String
Private sCntrlName As String

'Public Enum enCntrlPropertiesLet
'    '~~ -----------------------------------------------------
'    '~~ All properties applicable and 'writeable'
'    '~~ for any kind of Form-Control (Shape, OLEObject, etc.)
'    '~~ -----------------------------------------------------
'    enType
'    enName
'    enAccelerator
'    enAlignment
'    enAlternativeText
'    enAltHTML
'    enAutoLoad
'    enAutoShapeType
'    enAutoSize
'    enAutoUpdate
'    enBackColor
'    enBackgroundStyle
'    enBackStyle
'    enBlackWhiteMode
'    enBordersSuppress
'    enCaption
'    enDecorative
'    enEnabled
'    enFontBold
'    enFontItalic
'    enFontName
'    enFontSize
'    enFontStrikethru
'    enFontUnderline
'    enFontWeight
'    enForeColor
'    enGraphicStyle
'    enGroupName
'    enHeight
'    enLeft
'    enLinkedCell
'    enListFillRange
'    enLockAspectRatio
'    enLocked
'    enMouseIcon
'    enMousePointer
'    enMultiSelect
'    enOnAction
'    enPicture
'    enPicturePosition
'    enPlacement
'    enPrintObject
'    enRotation
'    enShadow
'    enShapeStyle
'    enSourceName
'    enSpecialEffect
'    enTakeFocusOnClick
'    enTextAlign
'    enTitle
'    enTop
'    enTripleState
'    enVisible
'    enWidth
'    enWordWrap
'    en_Font_Reserved
'End Enum

Private Property Get col(Optional ByVal o_name As String) As String
    Select Case o_name
        Case "Sync1cmbShape":               col = "C"
        
        Case "Sync1cmbActiveXoob":          col = "D"
        Case "Sync1cmbActiveXoobObject":    col = "E"
        
        Case "Sync2cbxShape":               col = "F"
        
        Case "Sync2cbxActiveXoob":          col = "G"
        Case "Sync2cbxActiveXoobObject":    col = "H"
        
        Case "Sync3optShape":               col = "I"
        
        Case "Sync3optActiveXoob":          col = "J"
        Case "Sync3optActiveXoobObject":    col = "K"
    End Select
End Property

Private Property Get CntrlType() As String:         CntrlType = sCntrlType: End Property
Private Property Let CntrlType(ByVal s As String):  sCntrlType = s:         End Property
Private Property Get CntrlName() As String:         CntrlName = sCntrlName: End Property
Private Property Let CntrlName(ByVal s As String):  sCntrlName = s:         End Property

Private Sub AssignColName(ByVal col As String)
    Dim rng As Range
    Dim nme As String
    Dim ref As String
    
    On Error GoTo eh
    If col <> vbNullString Then
        nme = "col" & CntrlName & "_" & CntrlType
        Set rng = wsSyncTestB.Range(col & 1).EntireColumn
        ref = "=" & rng.Parent.name & "!" & rng.Address
        ThisWorkbook.Names.Add name:=nme, RefersTo:=ref
    End If
xt: Exit Sub
eh: Debug.Print Err.Description: Stop: Resume
End Sub

Private Sub CntrlProperties(ByRef cntrl As Variant, _
                           ByVal cntrl_name As String, _
                           ByVal cntrl_col As String, _
                           ByVal cntrl_type As String)
    
    Dim i                   As Long
    Dim en                  As enCntrlProperties
    Dim s                   As String
    Dim oFont               As Font
    Dim oBorder             As Border
    Dim oTextEffectFormat   As TextEffectFormat
    Dim oTextFrame          As TextFrame
    Dim oTextFrame2         As TextFrame2
    Dim oAdjustments        As Adjustments
    
    On Error GoTo eh
    en = enType:                                CntrlPropertyValue en, cntrl_col, cntrl_type:           CntrlType = cntrl_type
    en = enName:                                CntrlPropertyValue en, cntrl_col, cntrl_name:           CntrlName = cntrl_name: AssignColName cntrl_col
    en = enOLEType:                             CntrlPropertyValue en, cntrl_col, cntrl.OLEType
    If cntrl.OLEType = xlOLELink Then
        en = enLinkStatus:                      CntrlPropertyValue en, cntrl_col, "Linked"
    Else
        en = enLinkStatus:                      CntrlPropertyValue en, cntrl_col, "Embedded"
    End If
    
    en = enAccelerator:                         CntrlPropertyValue en, cntrl_col, cntrl.Accelerator
    
    On Error GoTo c0
    Set oAdjustments = cntrl.Adjustments
    en = enAdjustmentsApplication:              CntrlPropertyValue en, cntrl_col, oAdjustments.Application
    en = enAdjustmentsCount:                    CntrlPropertyValue en, cntrl_col, oAdjustments.Count
    en = enAdjustmentsCreator:                  CntrlPropertyValue en, cntrl_col, oAdjustments.Creator
    For i = 1 To oAdjustments.Count
        en = enAdjustmentsItem
        On Error Resume Next
        CntrlPropertyValue en, cntrl_col, oAdjustments.Item(i) ' Access denied !!!
    Next i
    en = enAdjustmentsParentName:               CntrlPropertyValue en, cntrl_col, oAdjustments.Parent.name
    
c0: en = enAlignment:                           CntrlPropertyValue en, cntrl_col, cntrl.Alignment
    en = enAlternativeText:                     CntrlPropertyValue en, cntrl_col, cntrl.AlternativeText
    en = enAltHTML:                             CntrlPropertyValue en, cntrl_col, cntrl.AltHTML
    en = enApplication:                         CntrlPropertyValue en, cntrl_col, cntrl.Application
    en = enAutoLoad:                            CntrlPropertyValue en, cntrl_col, cntrl.AutoLoad
    en = enAutoShapeType:                       CntrlPropertyValue en, cntrl_col, cntrl.AutoShapeType
    en = enAutoSize:                            CntrlPropertyValue en, cntrl_col, cntrl.AutoSize
    en = enAutoUpdate:                          CntrlPropertyValue en, cntrl_col, cntrl.AutoUpdate
    en = enBackColor:                           CntrlPropertyValue en, cntrl_col, cntrl.BackColor
    en = enBackgroundStyle:                     CntrlPropertyValue en, cntrl_col, cntrl.BackgroundStyle
    en = enBackStyle:                           CntrlPropertyValue en, cntrl_col, cntrl.BackStyle
    en = enBlackWhiteMode:                      CntrlPropertyValue en, cntrl_col, cntrl.BlackWhiteMode
    en = enBordersSuppress:                     CntrlPropertyValue en, cntrl_col, cntrl.BordersSuppress
    
    On Error GoTo c1
    Set oBorder = cntrl.Border
    On Error GoTo eh
    en = enBorderColor:                         CntrlPropertyValue en, cntrl_col, oBorder.Color
    en = enBorderColorIndex:                    CntrlPropertyValue en, cntrl_col, oBorder.ColorIndex
    en = enBorderLineStyle:                     CntrlPropertyValue en, cntrl_col, oBorder.LineStyle
    en = enBorderParent:                        CntrlPropertyValue en, cntrl_col, oBorder.Parent
    en = enBorderThemeColor:                    CntrlPropertyValue en, cntrl_col, oBorder.ThemeColor
    en = enBorderTintAndShade:                  CntrlPropertyValue en, cntrl_col, oBorder.TintAndShade
    en = enBorderWeight:                        CntrlPropertyValue en, cntrl_col, oBorder.Weight
    
c1: en = enBottomRightCell:                     CntrlPropertyValue en, cntrl_col, cntrl.BottomRigthCell
    en = enCallout:                             CntrlPropertyValue en, cntrl_col, cntrl.Callout
    en = enCaption:                             CntrlPropertyValue en, cntrl_col, cntrl.Caption
    en = enChart:                               CntrlPropertyValue en, cntrl_col, cntrl.Chart
    en = enChild:                               CntrlPropertyValue en, cntrl_col, cntrl.Child
    en = enConnectionSiteCount:                 CntrlPropertyValue en, cntrl_col, cntrl.ConnectionSiteCount
    en = enConnector:                           CntrlPropertyValue en, cntrl_col, cntrl.Connector
    en = enConnectorFormat:                     CntrlPropertyValue en, cntrl_col, cntrl.ConnectorFormat
    en = enCount:                               CntrlPropertyValue en, cntrl_col, cntrl.Count
    en = enCreator:                             CntrlPropertyValue en, cntrl_col, cntrl.Creator
    en = enDecorative:                          CntrlPropertyValue en, cntrl_col, cntrl.Decorative
    en = enEnabled:                             CntrlPropertyValue en, cntrl_col, cntrl.Enabled
    en = enFill:                                CntrlPropertyValue en, cntrl_col, cntrl.fill
    
    On Error GoTo c2
    Set oFont = cntrl.Font
    On Error GoTo eh
    en = enFontApplication:                     CntrlPropertyValue en, cntrl_col, cntrl.Font.Application
    en = enFontBackground:                      CntrlPropertyValue en, cntrl_col, cntrl.Font.Background
    en = enFontBold:                            CntrlPropertyValue en, cntrl_col, cntrl.Font.Bold
    en = enFontColor:                           CntrlPropertyValue en, cntrl_col, cntrl.Font.Color
    en = enFontColorIndex:                      CntrlPropertyValue en, cntrl_col, cntrl.Font.ColorIndex
    en = enFontCreator:                         CntrlPropertyValue en, cntrl_col, cntrl.Font.Creator
    en = enFontFontStyle:                       CntrlPropertyValue en, cntrl_col, cntrl.Font.FontStyle
    en = enFontItalic:                          CntrlPropertyValue en, cntrl_col, cntrl.Font.Italic
    en = enFontName:                            CntrlPropertyValue en, cntrl_col, cntrl.Font.name
    en = enFontParent:                          CntrlPropertyValue en, cntrl_col, cntrl.Font.Parent
    en = enFontSize:                            CntrlPropertyValue en, cntrl_col, cntrl.Font.Size
    en = enFontStrikethrough:                   CntrlPropertyValue en, cntrl_col, cntrl.Font.Strikethrough
    en = enFontStrikethru:                      CntrlPropertyValue en, cntrl_col, cntrl.Font.Strikethru
    en = enFontSubscript:                       CntrlPropertyValue en, cntrl_col, cntrl.Font.Subscript
    en = enFontSuperscript:                     CntrlPropertyValue en, cntrl_col, cntrl.Font.Superscript
    en = enFontThemeColor:                      CntrlPropertyValue en, cntrl_col, cntrl.Font.ThemeColor
    en = enFontThemeFont:                       CntrlPropertyValue en, cntrl_col, cntrl.Font.ThemeFont
    en = enFontTintAndShade:                    CntrlPropertyValue en, cntrl_col, cntrl.Font.TintAndShade
    en = enFontUnderline:                       CntrlPropertyValue en, cntrl_col, cntrl.Font.Underline
    en = enFontWeight:                          CntrlPropertyValue en, cntrl_col, cntrl.Font.Weight
c2: en = enForeColor:                           CntrlPropertyValue en, cntrl_col, cntrl.ForeColor
    en = enGlow:                                CntrlPropertyValue en, cntrl_col, cntrl.Glow
    en = enGraphicStyle:                        CntrlPropertyValue en, cntrl_col, cntrl.GraphicStyle
    en = enGroupItems:                          CntrlPropertyValue en, cntrl_col, cntrl.Items
    en = enGroupName:                           CntrlPropertyValue en, cntrl_col, cntrl.GroupName
    en = enHasChart:                            CntrlPropertyValue en, cntrl_col, cntrl.Chart
    en = enHeight:                              CntrlPropertyValue en, cntrl_col, cntrl.Height
    en = enHorizontalFlip:                      CntrlPropertyValue en, cntrl_col, cntrl.HorizontalFlip
    en = enID:                                  CntrlPropertyValue en, cntrl_col, cntrl.ID
    en = enIndex:                               CntrlPropertyValue en, cntrl_col, cntrl.Index
    en = enInterior:                            CntrlPropertyValue en, cntrl_col, cntrl.Interior
    en = enLeft:                                CntrlPropertyValue en, cntrl_col, cntrl.Left
    en = enLine:                                CntrlPropertyValue en, cntrl_col, cntrl.Line
    en = enLinkedCell:                          CntrlPropertyValue en, cntrl_col, cntrl.LinkedCell
    en = enLinkStatus:                          CntrlPropertyValue en, cntrl_col, cntrl.LinkStatus
    en = enListFillRange:                       CntrlPropertyValue en, cntrl_col, cntrl.ListFillRange
    en = enLockAspectRatio:                     CntrlPropertyValue en, cntrl_col, cntrl.Ratio
    en = enLocked:                              CntrlPropertyValue en, cntrl_col, cntrl.Locked
    en = enModel3D:                             CntrlPropertyValue en, cntrl_col, cntrl.Model3D
    en = enMouseIcon:                           CntrlPropertyValue en, cntrl_col, cntrl.MouseIcon
    en = enMousePointer:                        CntrlPropertyValue en, cntrl_col, cntrl.MousePointer
    en = enMultiSelect:                         CntrlPropertyValue en, cntrl_col, cntrl.MultiSelect
    en = enNodes:                               CntrlPropertyValue en, cntrl_col, cntrl.Nodes
    en = enOLEType:                             CntrlPropertyValue en, cntrl_col, cntrl.OLEType
    en = enOnAction:                            CntrlPropertyValue en, cntrl_col, cntrl.OnAction
    en = enParent:                              CntrlPropertyValue en, cntrl_col, cntrl.Parent
    en = enParentGroup:                         CntrlPropertyValue en, cntrl_col, cntrl.ParebtGroup
    en = enPicture:                             CntrlPropertyValue en, cntrl_col, cntrl.Picture
    en = enPictureFormat:                       CntrlPropertyValue en, cntrl_col, cntrl.PictureFormat
    en = enPicturePosition:                     CntrlPropertyValue en, cntrl_col, cntrl.PicturePosition
    en = enPlacement:                           CntrlPropertyValue en, cntrl_col, cntrl.Placement
    en = enPrintObject:                         CntrlPropertyValue en, cntrl_col, cntrl.PrintObject
    en = enReflection:                          CntrlPropertyValue en, cntrl_col, cntrl.Reflection
    en = enRotation:                            CntrlPropertyValue en, cntrl_col, cntrl.Rotation
    en = enShadow:                              CntrlPropertyValue en, cntrl_col, cntrl.Shadow
    en = enShapeStyle:                          CntrlPropertyValue en, cntrl_col, cntrl.ShapeStyle
    en = enSoftEdge:                            CntrlPropertyValue en, cntrl_col, cntrl.SoftEdge
    en = enSourceName:                          CntrlPropertyValue en, cntrl_col, cntrl.SourceName
    en = enSpecialEffect:                       CntrlPropertyValue en, cntrl_col, cntrl.SpecialEffect
    en = enTakeFocusOnClick:                    CntrlPropertyValue en, cntrl_col, cntrl.TakeFocusOnClick
    en = enTextAlign:                           CntrlPropertyValue en, cntrl_col, cntrl.TextAlign
    
    On Error GoTo c3
    Set oTextEffectFormat = cntrl.TeextEffectFormat
    On Error GoTo eh
    en = enTextEffectFormatAlignment:           CntrlPropertyValue en, cntrl_col, oTextEffectFormat.Alignment
    en = enTextEffectFormatFontBold:            CntrlPropertyValue en, cntrl_col, oTextEffectFormat.FontBold
    en = enTextEffectFormatFontItalic:          CntrlPropertyValue en, cntrl_col, oTextEffectFormat.FontItalic
    en = enTextEffectFormatFontName:            CntrlPropertyValue en, cntrl_col, oTextEffectFormat.FontName
    en = enTextEffectFormatFontSize:            CntrlPropertyValue en, cntrl_col, oTextEffectFormat.FontSize
    en = enTextEffectFormatKernedPairs:         CntrlPropertyValue en, cntrl_col, oTextEffectFormat.KernedPairs
    en = enTextEffectFormatNormalizedHeight:    CntrlPropertyValue en, cntrl_col, oTextEffectFormat.NormalizedHeight
    en = enTextEffectFormatPresetShape:         CntrlPropertyValue en, cntrl_col, oTextEffectFormat.PresetShape
    en = enTextEffectFormatPresetTextEffect:    CntrlPropertyValue en, cntrl_col, oTextEffectFormat.PresetTextEffect
    en = enTextEffectFormatRotatedChars:        CntrlPropertyValue en, cntrl_col, oTextEffectFormat.RotatedChars
    en = enTextEffectFormattext:                CntrlPropertyValue en, cntrl_col, oTextEffectFormat.Text
    en = enTextEffectFormatTracking:            CntrlPropertyValue en, cntrl_col, oTextEffectFormat.Tracking
    
c3: On Error GoTo c4
    Set oTextFrame2 = cntrl.TextFrame2
    On Error GoTo eh
    en = enTextFrame2AutoSize:                  CntrlPropertyValue en, cntrl_col, oTextFrame2.AutoSize
    en = enTextFrame2HorizontalAnchor:          CntrlPropertyValue en, cntrl_col, oTextFrame2.HorizontalAnchor
    en = enTextFrame2MarginBottom:              CntrlPropertyValue en, cntrl_col, oTextFrame2.MarginBottom
    en = enTextFrame2MarginLeft:                CntrlPropertyValue en, cntrl_col, oTextFrame2.MarginLeft
    en = enTextFrame2MarginRight:               CntrlPropertyValue en, cntrl_col, oTextFrame2.MarginRight
    en = enTextFrame2MarginTop:                 CntrlPropertyValue en, cntrl_col, oTextFrame2.MarginTop
    en = enTextFrame2NoTextRotation:            CntrlPropertyValue en, cntrl_col, oTextFrame2.NoTextRotation
    en = enTextFrame2Orientation:               CntrlPropertyValue en, cntrl_col, oTextFrame2.Orientation
    en = enTextFrame2PathFormat:                CntrlPropertyValue en, cntrl_col, oTextFrame2.PathFormat
    en = enTextFrame2VerticalAnchor:            CntrlPropertyValue en, cntrl_col, oTextFrame2.VerticalAnchor
    en = enTextFrame2WarpForma:                 CntrlPropertyValue en, cntrl_col, oTextFrame2.WarpFormat
    en = enTextFrame2WordArtformat:             CntrlPropertyValue en, cntrl_col, oTextFrame2.WordArtformat
    
c4: On Error GoTo c5
    Set oTextFrame = cntrl.TextFrame
    On Error GoTo eh
    en = enTextFrameAutoMargins:                CntrlPropertyValue en, cntrl_col, oTextFrame.AutoMargins
    en = enTextFrameAutoSize:                   CntrlPropertyValue en, cntrl_col, oTextFrame.AutoSize
    en = enTextFrameHorizontalAlignment:        CntrlPropertyValue en, cntrl_col, oTextFrame.HorizontalAlignment
    en = enTextFrameHorizontalOverflow:         CntrlPropertyValue en, cntrl_col, oTextFrame.HorizontalOverflow
    en = enTextFrameMarginBottom:               CntrlPropertyValue en, cntrl_col, oTextFrame.MarginBottom
    en = enTextFrameMarginLeft:                 CntrlPropertyValue en, cntrl_col, oTextFrame.MarginLeft
    en = enTextFrameMarginRight:                CntrlPropertyValue en, cntrl_col, oTextFrame.MarginRight
    en = enTextFrameMarginTop:                  CntrlPropertyValue en, cntrl_col, oTextFrame.MarginTop
    en = enTextFrameOrientation:                CntrlPropertyValue en, cntrl_col, oTextFrame.Orientation
    en = enTextFrameReadingOrder:               CntrlPropertyValue en, cntrl_col, oTextFrame.ReadingOrder
    en = enTextFrameVerticalAlignment:          CntrlPropertyValue en, cntrl_col, oTextFrame.VerticalAlignment
    en = enTextFrameVerticalOverflow:           CntrlPropertyValue en, cntrl_col, oTextFrame.VerticalOverflow
    
c5: en = enThreeD:                              CntrlPropertyValue en, cntrl_col, cntrl.ThreeD
    en = enTitle:                               CntrlPropertyValue en, cntrl_col, cntrl.Title
    en = enTop:                                 CntrlPropertyValue en, cntrl_col, cntrl.Top
    en = enTripleState:                         CntrlPropertyValue en, cntrl_col, cntrl.TripleState
    en = enVerticalFlip:                        CntrlPropertyValue en, cntrl_col, cntrl.VerticalFlip
    en = enVertices:                            CntrlPropertyValue en, cntrl_col, cntrl.Vertices
    en = enVisible:                             CntrlPropertyValue en, cntrl_col, cntrl.Visible
    en = enWidth:                               CntrlPropertyValue en, cntrl_col, cntrl.Width
    en = enWordWrap:                            CntrlPropertyValue en, cntrl_col, cntrl.WordWrap
    en = enZOrderPosition:                      CntrlPropertyValue en, cntrl_col, cntrl.ZOrderPosition
    en = en_Font_Reserved:                      CntrlPropertyValue en, cntrl_col, cntrl.Font_Reserved
        
xt: Exit Sub

eh: Select Case Err.Number
        Case 438, 445, 1004, -2147024809, -2147467259, -2147024891
            '~~ All error indicating that the property is not applicable for the object
            wsSyncTestB.Range(cntrl_col & PropertyRow(en)).Value = "'-"
        Case 450
            Debug.Print cntrl_name & ": " & mSheetCntrlProperties.CntrlPropertyName(en) & ": Error = " & Err.Description
            Stop: Resume
        Case Else
            Debug.Print cntrl_name & ": " & mSheetCntrlProperties.CntrlPropertyName(en) & ": Error " & Err.Number & " = " & Err.Description
            Stop: Resume
    End Select
    Resume Next
End Sub

Private Sub CntrlPropertyNames()
    Dim i As enCntrlProperties
    For i = enCntrlProperties.enType To enCntrlProperties.enZOrderPosition
        wsSyncTestB.Range("A" & PropertyRow(i), "K" & PropertyRow(i)).ClearContents
        wsSyncTestB.Range("A" & PropertyRow(i)).Value = mSheetCntrlProperties.CntrlPropertyName(i)
    Next i
End Sub

Private Sub CntrlPropertyValue( _
                        ByVal p_enum As enCntrlProperties, _
                        ByVal p_col As String, _
                        ByRef p_prprty As Variant)
' -------------------------------------------------------------
'
' -------------------------------------------------------------
    Dim s As String
    
    With wsSyncTestB
        On Error Resume Next
        s = .Range(p_col & PropertyRow(p_enum)).Value
        If s <> vbNullString Then
           .Range(p_col & PropertyRow(p_enum)).Value = s & "|" & CStr(p_prprty)
        Else
            .Range(p_col & PropertyRow(p_enum)).Value = CStr(p_prprty)
        End If
    End With

xt: Exit Sub

eh: Select Case Err.Number
        Case 438:   If wsSyncTestB.ReadWrite(PropertyRow(p_enum)) = vbNullString Then wsSyncTestB.ReadWrite(PropertyRow(p_enum)) = "'-"
        Case Else:  Debug.Print "Property '" & mSheetCntrlProperties.CntrlPropertyName(p_enum) & "' Error " & Err.Number & ", " & Err.Description
    End Select
    Resume Next
End Sub

Private Function CntrlPropertyName(ByVal en As enCntrlProperties) As String
' ------------------------------------------------------------------------
' Returns the Property-Name for the provided enumerated property (en).
' ------------------------------------------------------------------------
    
    Select Case en
        Case enAccelerator:                         CntrlPropertyName = "Accelerator"
        Case enAdjustmentsApplication:              CntrlPropertyName = "Adjustments.Application"
        Case enAdjustmentsCount:                    CntrlPropertyName = "Adjustments.Count"
        Case enAdjustmentsCreator:                  CntrlPropertyName = "Adjustments.Creator"
        Case enAdjustmentsItem:                     CntrlPropertyName = "Adjustments.Item"
        Case enAdjustmentsParentName:               CntrlPropertyName = "Adjustments.Parent.Name"
        Case enAlignment:                           CntrlPropertyName = "Alignment"
        Case enAlternativeText:                     CntrlPropertyName = "AlternativeText"
        Case enAltHTML:                             CntrlPropertyName = "AltHTML"
        Case enApplication:                         CntrlPropertyName = "Application"
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
        Case enBordersSuppress:                     CntrlPropertyName = "BordersSupress"
        Case enBorderThemeColor:                    CntrlPropertyName = "Border.ThemeColor"
        Case enBorderTintAndShade:                  CntrlPropertyName = "Border.TintAndShade"
        Case enBorderWeight:                        CntrlPropertyName = "Weight"
        Case enBottomRightCell:                     CntrlPropertyName = "BottomRigthCell"
        Case enCallout:                             CntrlPropertyName = "CallOut"
        Case enCaption:                             CntrlPropertyName = "Caption"
        Case enChart:                               CntrlPropertyName = "Chart"
        Case enChild:                               CntrlPropertyName = "Child"
        Case enConnectionSiteCount:                 CntrlPropertyName = "ConnectionSiteCount"
        Case enConnector:                           CntrlPropertyName = "Connector"
        Case enConnectorFormat:                     CntrlPropertyName = "ConnectorFormat"
        Case enCount:                               CntrlPropertyName = "Count"
        Case enCreator:                             CntrlPropertyName = "Creator"
        Case enDecorative:                          CntrlPropertyName = "Decorative"
        Case enEnabled:                             CntrlPropertyName = "Enabled"
        Case enFill:                                CntrlPropertyName = "Fill"
        Case enFontApplication:                     CntrlPropertyName = "Font.Application"
        Case enFontBackground:                      CntrlPropertyName = "Font.Background"
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
        Case enFontStrikethru:                      CntrlPropertyName = "Font.Strikethru"
        Case enFontSubscript:                       CntrlPropertyName = "Font.Subscript"
        Case enFontSuperscript:                     CntrlPropertyName = "Font.Superscript"
        Case enFontThemeColor:                      CntrlPropertyName = "Font.ThemeColor"
        Case enFontThemeFont:                       CntrlPropertyName = "Font.ThemeFont"
        Case enFontTintAndShade:                    CntrlPropertyName = "Font.TintAndShade"
        Case enFontUnderline:                       CntrlPropertyName = "Font.Underline"
        Case enFontWeight:                          CntrlPropertyName = "Font.Weight"
        Case enForeColor:                           CntrlPropertyName = "ForeColor"
        Case enGlow:                                CntrlPropertyName = "Glow"
        Case enGraphicStyle:                        CntrlPropertyName = "GraphicStyle"
        Case enGroupItems:                          CntrlPropertyName = "Items"
        Case enGroupName:                           CntrlPropertyName = "GroupName"
        Case enHasChart:                            CntrlPropertyName = "Chart"
        Case enHeight:                              CntrlPropertyName = "Height"
        Case enHorizontalFlip:                      CntrlPropertyName = "HorizontalFlip"
        Case enID:                                  CntrlPropertyName = "ID"
        Case enIndex:                               CntrlPropertyName = "Index"
        Case enInterior:                            CntrlPropertyName = "Interior"
        Case enLeft:                                CntrlPropertyName = "Left"
        Case enLine:                                CntrlPropertyName = "Line"
        Case enLinkedCell:                          CntrlPropertyName = "LinkedCell"
        Case enLinkStatus:                          CntrlPropertyName = "LinkStatus"
        Case enListFillRange:                       CntrlPropertyName = "ListFillRange"
        Case enLockAspectRatio:                     CntrlPropertyName = "Ratio"
        Case enLocked:                              CntrlPropertyName = "Locked"
        Case enModel3D:                             CntrlPropertyName = "Model3D"
        Case enMouseIcon:                           CntrlPropertyName = "MouseIcon"
        Case enMousePointer:                        CntrlPropertyName = "MousPointer"
        Case enMultiSelect:                         CntrlPropertyName = "MultiSelect"
        Case enName:                                CntrlPropertyName = "Name"
        Case enNodes:                               CntrlPropertyName = "Nodes"
        Case enOLEType:                             CntrlPropertyName = "OLEType"
        Case enOnAction:                            CntrlPropertyName = "OnAction"
        Case enParent:                              CntrlPropertyName = "Parent"
        Case enParentGroup:                         CntrlPropertyName = "ParebtGroup"
        Case enPicture:                             CntrlPropertyName = "Picture"
        Case enPictureFormat:                       CntrlPropertyName = "PictureFormat"
        Case enPicturePosition:                     CntrlPropertyName = "PicturePosition"
        Case enPlacement:                           CntrlPropertyName = "Placement"
        Case enPrintObject:                         CntrlPropertyName = "PrintObject"
        Case enReflection:                          CntrlPropertyName = "Reflection"
        Case enRotation:                            CntrlPropertyName = "Rotation"
        Case enShadow:                              CntrlPropertyName = "Shadow"
        Case enShapeStyle:                          CntrlPropertyName = "ShapeStyle"
        Case enSoftEdge:                            CntrlPropertyName = "SoftEdge"
        Case enSourceName:                          CntrlPropertyName = "SourceName"
        Case enSpecialEffect:                       CntrlPropertyName = "SpecialEffect"
        Case enTakeFocusOnClick:                    CntrlPropertyName = "TakeFocusOnClick"
        Case enTextAlign:                           CntrlPropertyName = "TextAlign"
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
        Case enThreeD:                              CntrlPropertyName = "ThreeD"
        Case enTitle:                               CntrlPropertyName = "Title"
        Case enTop:                                 CntrlPropertyName = "Top"
        Case enTripleState:                         CntrlPropertyName = "TripleState"
        Case enType:                                CntrlPropertyName = "Type"
        Case enVerticalFlip:                        CntrlPropertyName = "VerticalFlip"
        Case enVertices:                            CntrlPropertyName = "Vertices"
        Case enVisible:                             CntrlPropertyName = "Visible"
        Case enWidth:                               CntrlPropertyName = "Width"
        Case enWordWrap:                            CntrlPropertyName = "WordWrap"
        Case enZOrderPosition:                      CntrlPropertyName = "ZOrderPosition"
        Case en_Font_Reserved:                      CntrlPropertyName = "_Font_Reserved"
    End Select
    
End Function

Public Sub CntrlPropertyValuesApplicable()

    Dim oob     As OLEObject
    Dim shp     As Shape
    Dim i       As Long
    Dim j       As Long
    
    Application.ScreenUpdating = False
    
    With wsSyncTestB
        If .UsedRange.Rows.CountLarge >= 2 And .UsedRange.Columns.CountLarge >= 3 Then
            Intersect(.UsedRange, .Range("2:" & .Rows.Count)).ClearContents ' Clear used values range
        End If
        CntrlPropertyNames
        
        For i = 1 To .Shapes.Count
            Set shp = .Shapes(i)
            Select Case shp.Type
                Case msoOLEControlObject
                    j = j + 1
                    Set oob = .OLEObjects(j)
                    On Error Resume Next
                    CntrlProperties cntrl:=oob _
                                      , cntrl_name:=oob.name _
                                      , cntrl_col:=col(oob.name & "oob") _
                                      , cntrl_type:=TypeName(oob) _

                    
                    On Error Resume Next
                    CntrlProperties cntrl:=oob.Object _
                                      , cntrl_name:=oob.name _
                                      , cntrl_col:=col(oob.name & "oobObject") _
                                      , cntrl_type:=TypeName(oob.Object)
                
                Case msoFormControl
                    CntrlProperties cntrl:=shp _
                                      , cntrl_name:=shp.name _
                                      , cntrl_col:=col(shp.name) _
                                      , cntrl_type:=mSheetControls.CntrlTypeString(shp.FormControlType)
                Case Else
                    Debug.Print "Shape-Type: '" & shp.Type & "' Not implemented"
            End Select
        Next i
    End With
    Application.ScreenUpdating = True
    
End Sub

Private Function PropertyRow(ByVal en As enCntrlProperties) As Long
    PropertyRow = en + 2
End Function

