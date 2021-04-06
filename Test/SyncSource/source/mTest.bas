Attribute VB_Name = "mTest"
Option Explicit

Private sObject As String

Private Enum enObjProperties
    enAAAAAAAA = 2
    enType
    enName
    enAdjustments
    enAlternativeText
    enApplication
    enAutoLoad
    enAutoShapeType
    enAutoSize
    enAutoUpdate
    enBackColor
    enBackgroundStyle
    enBackStyle
    enBlackWhiteMode
    enBorderColor
    enBorderColorIndex
    enBorderLineStyle
    enBorderParent
    enBorderThemeColor
    enBorderTintAndShade
    enBorderWeight
    enBottomRightCell
    enCallout
    enCaption
    enChart
    enChild
    enConnectionSiteCount
    enConnector
    enConnectorFormat
    enCount
    enCreator
    enDecorative
    enEnabled
    enFill
    enFont
    enForeColor
    enGlow
    enGraphicStyle
    enGroupItems
    enHasChart
    enHeight
    enHorizontalFlip
    enID
    enIndex
    enInterior
    enLeft
    enLine
    enLinkedCell
    enLinkStatus
    enListFillRange
    enLockAspectRatio
    enLocked
    enModel3D
    enMouseIcon
    enMousePointer
    enNodes
    enOLEType
    enParent
    enParentGroup
    enPicture
    enPictureFormat
    enPicturePosition
    enPlacement
    enPrintObject
    enReflection
    enRotation
    enShadow
    enShapeStyle
    enSoftEdge
    enTakeFocusOnClick
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
    enThreeD
    enTitle
    enTop
    enVerticalFlip
    enVertices
    enVisible
    enWidth
    enWordWrap
    enZOrder
    enZOrderPosition
    enZZZZZZZ
End Enum

Private Sub DisplayCtrlProperties()

    Dim oob     As OLEObject
    Dim sCol    As String
    Dim obj     As Object
    Dim shp     As Shape
    Dim sType   As String
    Dim i       As Long
    Dim j       As Long
    
    WritePropertyNames
    
    With wsSyncB
        For i = 1 To .Shapes.Count
            Set shp = .Shapes(i)
            Select Case shp.Type
                Case msoOLEControlObject
                    j = j + 1
                    Set oob = wsSyncB.OLEObjects(j)
                    sType = vbNullString
                    On Error Resume Next
                    sType = TypeName(oob.Object)
                    WriteCtrlProperties ctrl:=oob _
                                      , ctrl_name:=oob.name _
                                      , ctrl_type:=sType
                
                Case msoFormControl
                    sType = FormControlType(shp.FormControlType)
                    WriteCtrlProperties ctrl:=shp _
                                      , ctrl_name:=shp.name _
                                      , ctrl_type:=sType
                Case Else
                    Debug.Print "Not implemented"
            End Select
        Next i
    End With
    
End Sub

Private Sub WritePropertyNames()
    Dim i As enObjProperties
    For i = enAAAAAAAA + 1 To enZZZZZZZ - 1
        wsSyncB.Range("A" & i).Value = Title(i)
    Next i
End Sub

Private Sub WriteCtrlProperties(ByRef ctrl As Variant, _
                                ByVal ctrl_name As String, _
                                ByVal ctrl_type As String)
    Dim sCol    As String
    Dim en      As enObjProperties
    Dim obj     As Object
    Dim cmb     As CommandButton
    Dim cbx     As CheckBox
    Dim sName   As String
    Dim shp     As Shape
    
    On Error Resume Next
    With ctrl
        sCol = Col(ctrl_name)
        en = enType:                WriteCtrlProperty en, sCol, ctrl_type
        en = enName:                WriteCtrlProperty en, sCol, ctrl_name
        en = enOLEType:             WriteCtrlProperty en, sCol, .OLEType
        If .OLEType = xlOLELink Then
            en = enLinkStatus:      WriteCtrlProperty en, sCol, "Linked"
        Else
            en = enLinkStatus:      WriteCtrlProperty en, sCol, "Embedded"
        End If
        en = enAdjustments:         WriteCtrlProperty en, sCol, .Adjustments
        en = enAlternativeText:     WriteCtrlProperty en, sCol, .AlternativeText
        en = enApplication:         WriteCtrlProperty en, sCol, .Application
        en = enAutoLoad:            WriteCtrlProperty en, sCol, .AutoLoad
        en = enAutoShapeType:       WriteCtrlProperty en, sCol, .AutoShapeType
        en = enAutoSize:            WriteCtrlProperty en, sCol, .AutoSize
        en = enAutoUpdate:          WriteCtrlProperty en, sCol, .AutoUpdate
        en = enBackColor:           WriteCtrlProperty en, sCol, .BackColor
        en = enBackgroundStyle:     WriteCtrlProperty en, sCol, .BackgroundStyle
        en = enBackStyle:           WriteCtrlProperty en, sCol, .BackStyle
        en = enBlackWhiteMode:      WriteCtrlProperty en, sCol, .BlackWhiteMode
        en = enBorderColor:         WriteCtrlProperty en, sCol, .Border.Color
        en = enBorderColorIndex:    WriteCtrlProperty en, sCol, .Border.ColorIndex
        en = enBorderLineStyle:     WriteCtrlProperty en, sCol, .Border.LineStyle
        en = enBorderParent:        WriteCtrlProperty en, sCol, .Border.Parent
        en = enBorderThemeColor:    WriteCtrlProperty en, sCol, .Border.ThemeColor
        en = enBorderTintAndShade:  WriteCtrlProperty en, sCol, .Border.TintAndShade
        en = enBorderWeight:        WriteCtrlProperty en, sCol, .Border.Weight
        en = enBottomRightCell:     WriteCtrlProperty en, sCol, .BottomRigthCell
        en = enCallout:             WriteCtrlProperty en, sCol, .Callout
        en = enCaption:             WriteCtrlProperty en, sCol, .Caption
        en = enChart:               WriteCtrlProperty en, sCol, .Chart
        en = enChild:               WriteCtrlProperty en, sCol, .Child
        en = enConnectionSiteCount: WriteCtrlProperty en, sCol, .ConnectionSiteCount
        en = enConnector:           WriteCtrlProperty en, sCol, .Connector
        en = enConnectorFormat:     WriteCtrlProperty en, sCol, .ConnectorFormat
        en = enCount:               WriteCtrlProperty en, sCol, .Count
        en = enCreator:             WriteCtrlProperty en, sCol, .Creator
        en = enDecorative:          WriteCtrlProperty en, sCol, .Decorative
        en = enEnabled:             WriteCtrlProperty en, sCol, .Enabled
        en = enFill:                WriteCtrlProperty en, sCol, .fill
        en = enFont:                WriteCtrlProperty en, sCol, .Font
        en = enForeColor:           WriteCtrlProperty en, sCol, .ForeColor
        en = enGlow:                WriteCtrlProperty en, sCol, .Glow
        en = enGraphicStyle:        WriteCtrlProperty en, sCol, .GraphicStyle
        en = enGroupItems:          WriteCtrlProperty en, sCol, .Items
        en = enHasChart:            WriteCtrlProperty en, sCol, .Chart
        en = enHeight:              WriteCtrlProperty en, sCol, .Height
        en = enHorizontalFlip:      WriteCtrlProperty en, sCol, .HorizontalFlip
        en = enID:                  WriteCtrlProperty en, sCol, .ID
        en = enIndex:               WriteCtrlProperty en, sCol, .Index
        en = enInterior:            WriteCtrlProperty en, sCol, .Interior
        en = enLeft:                WriteCtrlProperty en, sCol, .Left
        en = enLine:                WriteCtrlProperty en, sCol, .Line
        en = enLinkedCell:          WriteCtrlProperty en, sCol, .LinkedCell
        en = enLinkStatus:          WriteCtrlProperty en, sCol, .LinkStatus
        en = enListFillRange:       WriteCtrlProperty en, sCol, .ListFillRange
        en = enLockAspectRatio:     WriteCtrlProperty en, sCol, .Ratio
        en = enLocked:              WriteCtrlProperty en, sCol, .Locked
        en = enModel3D:             WriteCtrlProperty en, sCol, .Model3D
        en = enMouseIcon:           WriteCtrlProperty en, sCol, .MouseIcon
        en = enMousePointer:        WriteCtrlProperty en, sCol, .MousePointer
        en = enNodes:               WriteCtrlProperty en, sCol, .Nodes
        en = enOLEType:             WriteCtrlProperty en, sCol, .OLEType
        en = enParent:              WriteCtrlProperty en, sCol, .Parent
        en = enParentGroup:         WriteCtrlProperty en, sCol, .ParebtGroup
        en = enPicture:             WriteCtrlProperty en, sCol, .Picture
        en = enPictureFormat:       WriteCtrlProperty en, sCol, .PictureFormat
        en = enPicturePosition:     WriteCtrlProperty en, sCol, .PicturePosition
        en = enPlacement:           WriteCtrlProperty en, sCol, .Placement
        en = enPrintObject:         WriteCtrlProperty en, sCol, .PrintObject
        en = enReflection:          WriteCtrlProperty en, sCol, .Reflection
        en = enRotation:            WriteCtrlProperty en, sCol, .Rotation
        en = enShadow:              WriteCtrlProperty en, sCol, .Shadow
        en = enShapeStyle:          WriteCtrlProperty en, sCol, .ShapeStyle
        en = enSoftEdge:            WriteCtrlProperty en, sCol, .SoftEdge
        en = enTakeFocusOnClick:    WriteCtrlProperty en, sCol, .TakeFocusOnClick
        en = enTextEffectFormatAlignment:           WriteCtrlProperty en, sCol, .TextEffectFormat.Alignment
        en = enTextEffectFormatFontBold:            WriteCtrlProperty en, sCol, .TextEffectFormat.FontBold
        en = enTextEffectFormatFontItalic:          WriteCtrlProperty en, sCol, .TextEffectFormat.FontItalic
        en = enTextEffectFormatFontName:            WriteCtrlProperty en, sCol, .TextEffectFormat.FontName
        en = enTextEffectFormatFontSize:            WriteCtrlProperty en, sCol, .TextEffectFormat.FontSize
        en = enTextEffectFormatKernedPairs:         WriteCtrlProperty en, sCol, .TextEffectFormat.KernedPairs
        en = enTextEffectFormatNormalizedHeight:    WriteCtrlProperty en, sCol, .TextEffectFormat.NormalizedHeight
        en = enTextEffectFormatPresetShape:         WriteCtrlProperty en, sCol, .TextEffectFormat.PresetShape
        en = enTextEffectFormatPresetTextEffect:    WriteCtrlProperty en, sCol, .TextEffectFormat.PresetTextEffect
        en = enTextEffectFormatRotatedChars:        WriteCtrlProperty en, sCol, .TextEffectFormat.RotatedChars
        en = enTextEffectFormattext:                WriteCtrlProperty en, sCol, .TextEffectFormat.text
        en = enTextEffectFormatTracking:            WriteCtrlProperty en, sCol, .TextEffectFormat.Tracking
        en = enTextFrame2AutoSize:                  WriteCtrlProperty en, sCol, .TextFrame2.AutoSize
        en = enTextFrame2HorizontalAnchor:          WriteCtrlProperty en, sCol, .TextFrame2.HorizontalAnchor
        en = enTextFrame2MarginBottom:              WriteCtrlProperty en, sCol, .TextFrame2.MarginBottom
        en = enTextFrame2MarginLeft:                WriteCtrlProperty en, sCol, .TextFrame2.MarginLeft
        en = enTextFrame2MarginRight:               WriteCtrlProperty en, sCol, .TextFrame2.MarginRight
        en = enTextFrame2MarginTop:                 WriteCtrlProperty en, sCol, .TextFrame2.MarginTop
        en = enTextFrame2NoTextRotation:            WriteCtrlProperty en, sCol, .TextFrame2.NoTextRotation
        en = enTextFrame2Orientation:               WriteCtrlProperty en, sCol, .TextFrame2.Orientation
        en = enTextFrame2PathFormat:                WriteCtrlProperty en, sCol, .TextFrame2.PathFormat
        en = enTextFrame2VerticalAnchor:            WriteCtrlProperty en, sCol, .TextFrame2.VerticalAnchor
        en = enTextFrame2WarpForma:                 WriteCtrlProperty en, sCol, .TextFrame2.WarpFormat
        en = enTextFrame2WordArtformat:             WriteCtrlProperty en, sCol, .TextFrame2.WordArtformat
        en = enTextFrameAutoMargins:                WriteCtrlProperty en, sCol, .TextFrame.AutoMargins
        en = enTextFrameAutoSize:                   WriteCtrlProperty en, sCol, .TextFrame.AutoSize
        en = enTextFrameHorizontalAlignment:        WriteCtrlProperty en, sCol, .TextFrame.HorizontalAlignment
        en = enTextFrameHorizontalOverflow:         WriteCtrlProperty en, sCol, .TextFrame.HorizontalOverflow
        en = enTextFrameMarginBottom:               WriteCtrlProperty en, sCol, .TextFrame.MarginBottom
        en = enTextFrameMarginLeft:                 WriteCtrlProperty en, sCol, .TextFrame.MarginLeft
        en = enTextFrameMarginRight:                WriteCtrlProperty en, sCol, .TextFrame.MarginRight
        en = enTextFrameMarginTop:                  WriteCtrlProperty en, sCol, .TextFrame.MarginTop
        en = enTextFrameOrientation:                WriteCtrlProperty en, sCol, .TextFrame.Orientation
        en = enTextFrameReadingOrder:               WriteCtrlProperty en, sCol, .TextFrame.ReadingOrder
        en = enTextFrameVerticalAlignment:          WriteCtrlProperty en, sCol, .TextFrame.VerticalAlignment
        en = enTextFrameVerticalOverflow:           WriteCtrlProperty en, sCol, .TextFrame.VerticalOverflow
        en = enThreeD:              WriteCtrlProperty en, sCol, .ThreeD
        en = enTitle:               WriteCtrlProperty en, sCol, .Title
        en = enTop:                 WriteCtrlProperty en, sCol, .Top
        en = enVerticalFlip:        WriteCtrlProperty en, sCol, .VerticalFlip
        en = enVertices:            WriteCtrlProperty en, sCol, .Vertices
        en = enVisible:             WriteCtrlProperty en, sCol, .Visible
        en = enWidth:               WriteCtrlProperty en, sCol, .Width
        en = enWordWrap:            WriteCtrlProperty en, sCol, .WordWrap
        en = enZOrder:              WriteCtrlProperty en, sCol, .ZOrder
        en = enZOrderPosition:      WriteCtrlProperty en, sCol, .ZOrderPosition
    End With
        
xt: Exit Sub

eh:
End Sub

Private Sub WriteCtrlProperty( _
                        ByVal p_enum As enObjProperties, _
                        ByVal p_col As String, _
                        ByRef p_prprty As Variant)
' -------------------------------------------------------------
'
' -------------------------------------------------------------
    With wsSyncB
        On Error Resume Next
        .Range(p_col & p_enum).Value = CStr(p_prprty)
    End With
End Sub

Private Property Get Col(Optional ByVal o_name As String) As String
    Select Case o_name
        Case "cmbSync1": Col = "C"
        Case "cmbSync2": Col = "D"
        Case "cbxSync1": Col = "E"
        Case "cbxSync2": Col = "F"
        Case "optSync1": Col = "G"
        Case "optSync2": Col = "H"
    End Select
End Property

Private Property Get Title(Optional ByVal p_enum As enObjProperties) As String
    
    Select Case p_enum
        Case enAdjustments:         Title = "Adjustments"
        Case enAlternativeText:     Title = "AlternativeText"
        Case enApplication:         Title = "Application"
        Case enAutoLoad:            Title = "AutoLoad"
        Case enAutoShapeType:       Title = "AutoShapeType"
        Case enAutoSize:            Title = "AutoSize"
        Case enAutoUpdate:          Title = "AutoUpdate"
        Case enBackColor:           Title = "BackColor"
        Case enBackgroundStyle:     Title = "BackgroundStyle"
        Case enBackStyle:           Title = "BackStyle"
        Case enBlackWhiteMode:      Title = "BlackWhitemode"
        Case enBorderColor:         Title = "Border.Color"
        Case enBorderColorIndex:    Title = "Border.ColorIndex"
        Case enBorderLineStyle:     Title = "Border.LineStyle"
        Case enBorderParent:        Title = "Border.Parent"
        Case enBorderThemeColor:    Title = "Border.ThemeColor"
        Case enBorderTintAndShade:  Title = "Border.TintAndShade"
        Case enBorderWeight:        Title = "Weight"
        Case enBottomRightCell:     Title = "BottomRigthCell"
        Case enCallout:             Title = "CallOut"
        Case enCaption:             Title = "Caption"
        Case enChart:               Title = "Chart"
        Case enChild:               Title = "Child"
        Case enConnectionSiteCount: Title = "ConnectionSiteCount"
        Case enConnector:           Title = "Connector"
        Case enConnectorFormat:     Title = "ConnectorFormat"
        Case enCount:               Title = "Count"
        Case enCreator:             Title = "Creator"
        Case enDecorative:          Title = "Decorative"
        Case enEnabled:             Title = "Enabled"
        Case enFill:                Title = "Fill"
        Case enFont:                Title = "Font"
        Case enForeColor:           Title = "ForeColor"
        Case enGlow:                Title = "Glow"
        Case enGraphicStyle:        Title = "GraphicStyle"
        Case enGroupItems:          Title = "Items"
        Case enHasChart:            Title = "Chart"
        Case enHeight:              Title = "Height"
        Case enHorizontalFlip:      Title = "HorizontalFlip"
        Case enID:                  Title = "ID"
        Case enIndex:               Title = "Index"
        Case enInterior:            Title = "Interior"
        Case enLeft:                Title = "Left"
        Case enLine:                Title = "Line"
        Case enLinkedCell:                          Title = "LinkedCell"
        Case enLinkStatus:                          Title = "LinkStatus"
        Case enListFillRange:                       Title = "ListFillRange"
        Case enLockAspectRatio:                     Title = "Ratio"
        Case enLocked:                              Title = "Locked"
        Case enModel3D:                             Title = "Model3D"
        Case enMouseIcon:                           Title = "MouseIcon"
        Case enMousePointer:                        Title = "MousPointer"
        Case enName:                                Title = "Name"
        Case enNodes:                               Title = "Nodes"
        Case enOLEType:                             Title = "OLEType"
        Case enParent:                              Title = "Parent"
        Case enParentGroup:                         Title = "ParebtGroup"
        Case enPicture:                             Title = "Picture"
        Case enPictureFormat:                       Title = "PictureFormat"
        Case enPicturePosition:                     Title = "PicturePosition"
        Case enPlacement:                           Title = "Placement"
        Case enPrintObject:                         Title = "PrintObject"
        Case enReflection:                          Title = "Reflection"
        Case enRotation:                            Title = "Rotation"
        Case enShadow:                              Title = "Shadow"
        Case enShapeStyle:                          Title = "ShapeStyle"
        Case enSoftEdge:                            Title = "SoftEdge"
        Case enTakeFocusOnClick:                    Title = "TakeFocusOnClick"
        Case enTextEffectFormatAlignment:           Title = "TextEffectFormat.Alignment"
        Case enTextEffectFormatFontBold:            Title = "TextEffectFormat.FontBold"
        Case enTextEffectFormatFontItalic:          Title = "TextEffectFormat.FontItalic"
        Case enTextEffectFormatFontName:            Title = "TextEffectFormat.FontName"
        Case enTextEffectFormatFontSize:            Title = "TextEffectFormat.FontSize"
        Case enTextEffectFormatKernedPairs:         Title = "TextEffectFormat.KernedPairs"
        Case enTextEffectFormatNormalizedHeight:    Title = "TextEffectFormat.NormalizedHeight"
        Case enTextEffectFormatPresetShape:         Title = "TextEffectFormat.PresetShape"
        Case enTextEffectFormatPresetTextEffect:    Title = "TextEffectFormat.PresetTextEffect"
        Case enTextEffectFormatRotatedChars:        Title = "TextEffectFormat.RotatedChars"
        Case enTextEffectFormattext:                Title = "TextEffectFormat.text"
        Case enTextEffectFormatTracking:            Title = "TextEffectFormat.Tracking"
        Case enTextFrame2AutoSize:                  Title = "TextFrame2.AutoSize"
        Case enTextFrame2HorizontalAnchor:          Title = "TextFrame2.HorizontalAnchor"
        Case enTextFrame2MarginBottom:              Title = "TextFrame2.MarginBottom"
        Case enTextFrame2MarginLeft:                Title = "TextFrame2.MarginLeft"
        Case enTextFrame2MarginRight:               Title = "TextFrame2.MarginRight"
        Case enTextFrame2MarginTop:                 Title = "TextFrame2.MarginTop"
        Case enTextFrame2NoTextRotation:            Title = "TextFrame2.NoTextRotation"
        Case enTextFrame2Orientation:               Title = "TextFrame2.Orientation"
        Case enTextFrame2PathFormat:                Title = "TextFrame2.PathFormat"
        Case enTextFrame2VerticalAnchor:            Title = "TextFrame2.VerticalAnchor"
        Case enTextFrame2WarpForma:                 Title = "TextFrame2.WarpFormat"
        Case enTextFrame2WordArtformat:             Title = "TextFrame2.WordArtformat"
        Case enTextFrameAutoMargins:                Title = "TextFrame.AutoMargins"
        Case enTextFrameAutoSize:                   Title = "TextFrame.AutoSize"
        Case enTextFrameHorizontalAlignment:        Title = "TextFrame.HorizontalAlignment"
        Case enTextFrameHorizontalOverflow:         Title = "TextFrame.HorizontalOverflow"
        Case enTextFrameMarginBottom:               Title = "TextFrame.MarginBottom"
        Case enTextFrameMarginLeft:                 Title = "TextFrame.MarginLeft"
        Case enTextFrameMarginRight:                Title = "TextFrame.MarginRight"
        Case enTextFrameMarginTop:                  Title = "TextFrame.MarginTop"
        Case enTextFrameOrientation:                Title = "TextFrame.Orientation"
        Case enTextFrameReadingOrder:               Title = "TextFrame.ReadingOrder"
        Case enTextFrameVerticalAlignment:          Title = "TextFrame.VerticalAlignment"
        Case enTextFrameVerticalOverflow:           Title = "TextFrame.VerticalOverflow"
        Case enThreeD:                              Title = "ThreeD"
        Case enTitle:                               Title = "Title"
        Case enTop:                                 Title = "Top"
        Case enType:                                Title = "Type"
        Case enVerticalFlip:                        Title = "VerticalFlip"
        Case enVertices:                            Title = "Vertices"
        Case enVisible:                             Title = "Visible"
        Case enWidth:                               Title = "Width"
        Case enWordWrap:                            Title = "WordWrap"
        Case enZOrder:                              Title = "ZOrder"
        Case enZOrderPosition:                      Title = "ZOrderPosition"
    End Select
    
End Property

Public Sub SheetObjects()
    
    Dim oob As OLEObject
    Dim shp As Shape
    Dim cmb As CommandButton
    Dim i   As Long
    Dim j   As Long
    
    For i = 1 To wsSyncB.Shapes.Count
        Set shp = wsSyncB.Shapes(i)
        Select Case shp.Type
            Case msoOLEControlObject
                j = j + 1
                Set oob = wsSyncB.OLEObjects(j)
                Debug.Print i & "." & j & ". ", TypeName(oob) & ":", "Name = " & oob.name & ",", "Type = " & TypeName(oob.Object)
            
            Case msoFormControl:        Debug.Print i & ". ", TypeName(shp) & ":", "Name = " & shp.name & ",", "Type = " & FormControlType(shp.FormControlType)
            Case Else:                  Debug.Print i & ". ", TypeName(shp) & ":", "Name = " & shp.name & ",", "Type = " & ShapeType(shp.Type)
        End Select
    Next i
    


End Sub

Private Function FormControlType(ByVal en As Variant) As String

    Select Case en
        Case xlButtonControl:   FormControlType = "CommandButton"
        Case xlCheckBox:        FormControlType = "CheckBox"
        Case xlDropDown:        FormControlType = "ComboBox"
        Case xlEditBox:         FormControlType = "TextBox"
        Case xlGroupBox:        FormControlType = "GroupBox"
        Case xlLabel:           FormControlType = "Label"
        Case xlListBox:         FormControlType = "ListBox"
        Case xlOptionButton:    FormControlType = "OptionButton"
        Case xlScrollBar:       FormControlType = "ScrollBar"
        Case xlSpinner:         FormControlType = "Spinner"
    End Select
End Function

Private Function ShapeType(ByVal l As MsoShapeType) As String
    Select Case l
        Case mso3DModel:            ShapeType = "3D model"
        Case msoAutoShape:          ShapeType = "AutoShape"
        Case msoCallout:            ShapeType = "Callout"
        Case msoCanvas:             ShapeType = "Canvas"
        Case msoChart:              ShapeType = "Chart"
        Case msoComment:            ShapeType = "Comment"
        Case msoContentApp:         ShapeType = "Content Office Add-in"
        Case msoDiagram:            ShapeType = "Diagram"
        Case msoEmbeddedOLEObject:  ShapeType = "Embedded OLE object"
        Case msoFormControl:        ShapeType = "Form control"
        Case msoFreeform:           ShapeType = "Freeform"
        Case msoGraphic:            ShapeType = "Graphic"
        Case msoGroup:              ShapeType = "Group"
        Case msoInk:                ShapeType = "Ink"
        Case msoInkComment:         ShapeType = "Ink comment"
        Case msoLine:               ShapeType = "Line"
        Case msoLinked3DModel:      ShapeType = "Linked 3D model"
        Case msoLinkedGraphic:      ShapeType = "Linked graphic"
        Case msoLinkedOLEObject:    ShapeType = "Linked OLE object"
        Case msoLinkedPicture:      ShapeType = "Linked picture"
        Case msoMedia:              ShapeType = "Media"
        Case msoOLEControlObject:   ShapeType = "OLE control object"
        Case msoPicture:            ShapeType = "Picture"
        Case msoPlaceholder:        ShapeType = "Placeholder"
        Case msoScriptAnchor:       ShapeType = "Script anchor"
        Case msoShapeTypeMixed:     ShapeType = "Mixed shape type"
        Case msoSlicer:             ShapeType = "Slicer"
        Case msoTable:              ShapeType = "Table"
        Case msoTextBox:            ShapeType = "Text box"
        Case msoTextEffect:         ShapeType = "Text effect"
        Case msoWebVideo:           ShapeType = "Web video"
    
    End Select
End Function
