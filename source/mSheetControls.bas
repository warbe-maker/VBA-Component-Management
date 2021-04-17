Attribute VB_Name = "mSheetControls"
Option Explicit

Public Enum enCntrlProperties
    '~~ -----------------------------------------------------
    '~~ All potentially applicable properties
    '~~ for any kind of Form-Control (Shape, OLEObject, etc.)
    '~~ -----------------------------------------------------
    enType
    enName
    enAccelerator
    enAdjustments
    enAlignment
    enAlternativeText
    enAltHTML
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
    enBordersSuppress
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
    enFontApplication
    enFontBackground
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
    enFontStrikethru
    enFontSubscript
    enFontSuperscript
    enFontThemeColor
    enFontThemeFont
    enFontTintAndShade
    enFontUnderline
    enFontWeight
    enForeColor
    enGlow
    enGraphicStyle
    enGroupItems
    enGroupName
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
    enMultiSelect
    enNodes
    enOLEType
    enOnAction
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
    enThreeD
    enTitle
    enTop
    enTripleState
    enVerticalFlip
    enVertices
    enVisible
    enWidth
    enWordWrap
    enZOrderPosition
    en_Font_Reserved
End Enum

Private Enum VarEnum
    VT_EMPTY = 0&                   '
    VT_NULL = 1&                    ' 0
    VT_I2 = 2&                      ' signed 2 bytes integer
    VT_I4 = 3&                      ' signed 4 bytes integer
    VT_R4 = 4&                      ' 4 bytes float
    VT_R8 = 5&                      ' 8 bytes float
    VT_CY = 6&                      ' currency
    VT_DATE = 7&                    ' date
    VT_BSTR = 8&                    ' BStr
    VT_DISPATCH = 9&                ' IDispatch
    VT_ERROR = 10&                  ' error value
    VT_BOOL = 11&                   ' boolean
    VT_VARIANT = 12&                ' variant
    VT_UNKNOWN = 13&                ' IUnknown
    VT_DECIMAL = 14&                ' decimal
    VT_I1 = 16&                     ' signed byte
    VT_UI1 = 17&                    ' unsigned byte
    VT_UI2 = 18&                    ' unsigned 2 bytes integer
    VT_UI4 = 19&                    ' unsigned 4 bytes integer
    VT_I8 = 20&                     ' signed 8 bytes integer
    VT_UI8 = 21&                    ' unsigned 8 bytes integer
    VT_INT = 22&                    ' integer
    VT_UINT = 23&                   ' unsigned integer
    VT_VOID = 24&                   ' 0
    VT_HRESULT = 25&                ' HRESULT
    VT_PTR = 26&                    ' pointer
    VT_SAFEARRAY = 27&              ' safearray
    VT_CARRAY = 28&                 ' carray
    VT_USERDEFINED = 29&            ' userdefined
    VT_LPSTR = 30&                  ' LPStr
    VT_LPWSTR = 31&                 ' LPWStr
    VT_RECORD = 36&                 ' Record
    VT_FILETIME = 64&               ' File Time
    VT_BLOB = 65&                   ' Blob
    VT_STREAM = 66&                 ' Stream
    VT_STORAGE = 67&                ' Storage
    VT_STREAMED_OBJECT = 68&        ' Streamed Obj
    VT_STORED_OBJECT = 69&          ' Stored Obj
    VT_BLOB_OBJECT = 70&            ' Blob Obj
    VT_CF = 71&                     ' CF
    VT_CLSID = 72&                  ' Class ID
    VT_BSTR_BLOB = &HFFF&           ' BStr Blob
    VT_VECTOR = &H1000&             ' Vector
    VT_ARRAY = &H2000&              ' Array
    VT_BYREF = &H4000&              ' ByRef
    VT_RESERVED = &H8000&           ' Reserved
    VT_ILLEGAL = &HFFFF&            ' illegal
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TTYPEDESC
    #If Win64 Then
        pTypeDesc As LongLong
    #Else
        pTypeDesc As Long
    #End If
    vt   As Integer
End Type

Private Type TPARAMDESC
    #If Win64 Then
        pPARAMDESCEX  As LongLong
    #Else
        pPARAMDESCEX  As Long
    #End If
    wParamFlags  As Integer
End Type

Private Type TELEMDESC
    tdesc  As TTYPEDESC
    pdesc  As TPARAMDESC
End Type

Private Type TYPEATTR
        aGUID As GUID
        LCID As Long
        dwReserved As Long
        memidConstructor As Long
        memidDestructor As Long
        #If Win64 Then
            lpstrSchema As LongLong
        #Else
            lpstrSchema As Long
        #End If
        cbSizeInstance As Integer
        typekind As Long
        cFuncs As Integer
        cVars As Integer
        cImplTypes As Integer
        cbSizeVft As Integer
        cbAlignment As Integer
        wTypeFlags As Integer
        wMajorVerNum As Integer
        wMinorVerNum As Integer
        tdescAlias As Long
        idldescType As Long
End Type

Private Type FUNCDESC
    memid As Long
    #If Win64 Then
        lReserved1 As LongLong
        lprgelemdescParam As LongLong
    #Else
        lReserved1 As Long
        lprgelemdescParam As Long
    #End If
    funckind As Long
    INVOKEKIND As Long
    CallConv As Long
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer
    cReserved2 As Integer
    elemdescFunc As TELEMDESC
    wFuncFlags As Integer
End Type

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)

Public Function GetObjectFunctions(ByVal TheObject As IUnknown, Optional ByVal FuncType As VbCallType) As Variant()

    #If Win64 Then
        Const vTblOffsetFac_32_64 = 2
        Dim aTYPEATTR()         As LongLong, aFUNCDESC() As LongLong, farPtr As LongLong
    #Else
        Const vTblOffsetFac_32_64 = 1
        Dim aTYPEATTR()         As Long, aFUNCDESC() As Long, farPtr As Long
    #End If
    
    Const CC_STDCALL            As Long = 4
    Const IUNK_QueryInterface   As Long = 0
    Const IDSP_GetTypeInfo      As Long = 16 * vTblOffsetFac_32_64
    Const ITYP_GetTypeAttr      As Long = 12 * vTblOffsetFac_32_64
    Const ITYP_GetFuncDesc      As Long = 20 * vTblOffsetFac_32_64
    Const ITYP_GetDocument      As Long = 48 * vTblOffsetFac_32_64
    Const ITYP_ReleaseTypeAttr  As Long = 76 * vTblOffsetFac_32_64
    Const ITYP_ReleaseFuncDesc  As Long = 80 * vTblOffsetFac_32_64
    
    Dim tTYPEATTR               As TYPEATTR
    Dim tFUNCDESC               As FUNCDESC
    Dim aGUID(0 To 11)          As Long
    Dim lFuncsCount             As Long
    Dim ITypeInfo               As IUnknown
    Dim IDispatch               As IUnknown
    Dim sName                   As String
    Dim lRequestedFuncsCount    As Long
    Dim n                       As Long
    Dim Index                   As Long


    aGUID(0) = &H20400: aGUID(2) = &HC0&: aGUID(3) = &H46000000
    Call vtblCall(ObjPtr(TheObject), IUNK_QueryInterface, vbLong, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IDispatch))
    If IDispatch Is Nothing Then MsgBox "error":   Exit Function

    Call vtblCall(ObjPtr(IDispatch), IDSP_GetTypeInfo, vbLong, CC_STDCALL, 0&, 0&, VarPtr(ITypeInfo))
    If ITypeInfo Is Nothing Then MsgBox "error": Exit Function
    
    Call vtblCall(ObjPtr(ITypeInfo), ITYP_GetTypeAttr, vbLong, CC_STDCALL, VarPtr(farPtr))
    If farPtr = 0& Then MsgBox "error": Exit Function

    Call CopyMemory(ByVal VarPtr(tTYPEATTR), ByVal farPtr, LenB(tTYPEATTR))
    ReDim aTYPEATTR(LenB(tTYPEATTR))
    Call CopyMemory(ByVal VarPtr(aTYPEATTR(0)), tTYPEATTR, UBound(aTYPEATTR))
    Call vtblCall(ObjPtr(ITypeInfo), ITYP_ReleaseTypeAttr, vbEmpty, CC_STDCALL, farPtr)
    
    If tTYPEATTR.cFuncs Then
    
        ReDim vFuncArray(tTYPEATTR.cFuncs, 6) As Variant
        
        For lFuncsCount = 0 To tTYPEATTR.cFuncs - 1
        
            Call vtblCall(ObjPtr(ITypeInfo), ITYP_GetFuncDesc, vbLong, CC_STDCALL, lFuncsCount, VarPtr(farPtr))
            If farPtr = 0 Then GoTo nextfunc
            Call CopyMemory(ByVal VarPtr(tFUNCDESC), ByVal farPtr, LenB(tFUNCDESC))
            ReDim aFUNCDESC(LenB(tFUNCDESC))
            Call CopyMemory(ByVal VarPtr(aFUNCDESC(0)), tFUNCDESC, UBound(aFUNCDESC))
            Call vtblCall(ObjPtr(ITypeInfo), ITYP_ReleaseFuncDesc, vbEmpty, CC_STDCALL, farPtr)
            Call vtblCall(ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(sName), 0, 0, 0)
            
            With tFUNCDESC
                If .INVOKEKIND And FuncType Then
                    vFuncArray(lFuncsCount, 0) = sName
                    vFuncArray(lFuncsCount, 1) = .memid
                    vFuncArray(lFuncsCount, 2) = .oVft
                    vFuncArray(lFuncsCount, 3) = Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                    vFuncArray(lFuncsCount, 4) = .cParams
                    vFuncArray(lFuncsCount, 5) = ReturnType(VarPtr(.elemdescFunc.tdesc))
                    lRequestedFuncsCount = lRequestedFuncsCount + 1
                End If
            End With
            sName = vbNullString
nextfunc:
        Next
            
        ReDim vFuncsRequestedArray(lRequestedFuncsCount, 6)
        For n = 0 To UBound(vFuncArray, 1) - 1
            If vFuncArray(n, 1) <> Empty Then
                vFuncsRequestedArray(Index, 0) = vFuncArray(n, 0)
                vFuncsRequestedArray(Index, 1) = vFuncArray(n, 1)
                vFuncsRequestedArray(Index, 2) = vFuncArray(n, 2)
                vFuncsRequestedArray(Index, 3) = vFuncArray(n, 3)
                vFuncsRequestedArray(Index, 4) = vFuncArray(n, 4)
                vFuncsRequestedArray(Index, 5) = vFuncArray(n, 5)
                Index = Index + 1
            End If
        Next n
        
        GetObjectFunctions = vFuncsRequestedArray
    End If

End Function

#If Win64 Then
    Private Function vtblCall(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    Dim vParamPtr() As LongLong
#Else
    Private Function vtblCall(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    Dim vParamPtr() As Long
#End If

    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function

    Dim pIndex          As Long
    Dim pCount          As Long
    Dim vParamType()    As Integer
    Dim vRtn            As Variant
    Dim vParams()       As Variant

    vParams() = FunctionParameters()
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If

    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, _
    vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then
        vtblCall = vRtn
    Else
        SetLastError pIndex
    End If

End Function

Private Function ReturnType(Ptr As LongPtr) As String

    Dim sName   As String
    Dim tdesc   As TTYPEDESC

    Call CopyMemory(tdesc, ByVal Ptr, Len(tdesc))

    Select Case tdesc.vt
        Case VT_NULL:       sName = "Long"
        Case VT_I2:         sName = "Integer"
        Case VT_I4:         sName = "Long"
        Case VT_R4:         sName = "Single"
        Case VT_R8:         sName = "Double"
        Case VT_CY:         sName = "CY"
        Case VT_DATE:       sName = "DATE"
        Case VT_BSTR:       sName = "BSTR"
        Case VT_DISPATCH:   sName = "IDispatch*"
        Case VT_ERROR:      sName = "SCODE"
        Case VT_BOOL:       sName = "Boolean"
        Case VT_VARIANT:    sName = "VARIANT"
        Case VT_UNKNOWN:    sName = "IUnknown*"
        Case VT_UI1:        sName = "BYTE"
        Case VT_DECIMAL:    sName = "DECIMAL"
        Case VT_I1:         sName = "Char"
        Case VT_UI2:        sName = "USHORT"
        Case VT_UI4:        sName = "ULONG"
        Case VT_I8:         sName = "__int64"
        Case VT_UI8:        sName = "unsigned __int64"
        Case VT_INT:        sName = "int"
        Case VT_UINT:       sName = "UINT"
        Case VT_HRESULT:    sName = "HRESULT"
        Case VT_VOID:       sName = "VOID"
        Case VT_LPSTR:      sName = "char*"
        Case VT_LPWSTR:     sName = "wchar_t*"
        Case Else:          sName = "ANY"
    End Select
    ReturnType = sName
    
End Function

Public Function CntrlName(ByVal shp As Shape) As String
    Const PROC = "CntrlName"
    
    On Error GoTo eh
    Dim oob     As OLEObject
    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim sName   As String
    Dim i       As Long
    Dim j       As Long
    Dim shp1    As Shape
    
    Set wb = Application.Workbooks(shp.Parent.Parent.Name)
    Set ws = wb.Worksheets(shp.Parent.Name)
    With ws
        For Each shp1 In .Shapes
            If shp1.Type = msoOLEControlObject Then j = j + 1
            If Not shp1 Is shp Then GoTo next_shp
            Select Case shp1.Type
                Case msoOLEControlObject
                    CntrlName = ws.OLEObjects(j).Name
                Case msoFormControl
                    CntrlName = shp1.Name
                Case Else
                    Debug.Print "Shape-Type: '" & shp1.Type & "' Not implemented"
            End Select
next_shp:
        Next shp1
    End With

xt: Exit Function

eh: Debug.Print Err.Description: Stop: Resume
End Function

Public Function CntrlType( _
                    ByVal shp As Shape) As String
' -------------------------------------------------
'
' -------------------------------------------------
    Const PROC = "CntrlType"
    
    On Error GoTo eh
    Dim oob     As OLEObject
    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim sName   As String
    Dim i       As Long
    Dim j       As Long
    Dim shp1    As Shape
    
    Set wb = Application.Workbooks(shp.Parent.Parent.Name)
    Set ws = wb.Worksheets(shp.Parent.Name)
    With ws
        For Each shp1 In .Shapes
            If shp1.Type = msoOLEControlObject Then
                j = j + 1
            End If
            If Not shp1 Is shp Then
                GoTo next_shp1
            End If
            Select Case shp1.Type
                Case msoOLEControlObject
                    CntrlType = TypeName(ws.OLEObjects(j).Object)
                Case msoFormControl
                    CntrlType = CntrlTypeString(shp.FormControlType)
                Case Else
                    Debug.Print "Shape-Type: '" & shp.Type & "' Not implemented"
            End Select
next_shp1:
        Next shp1
    End With

xt: Exit Function

eh: Debug.Print Err.Description: Stop: Resume
End Function

Public Function CntrlTypeString(ByVal en As XlFormControl) As String

    Select Case en
        Case xlButtonControl:   CntrlTypeString = "CommandButton"
        Case xlCheckBox:        CntrlTypeString = "CheckBox"
        Case xlDropDown:        CntrlTypeString = "ComboBox"
        Case xlEditBox:         CntrlTypeString = "TextBox"
        Case xlGroupBox:        CntrlTypeString = "GroupBox"
        Case xlLabel:           CntrlTypeString = "Label"
        Case xlListBox:         CntrlTypeString = "ListBox"
        Case xlOptionButton:    CntrlTypeString = "OptionButton"
        Case xlScrollBar:       CntrlTypeString = "ScrollBar"
        Case xlSpinner:         CntrlTypeString = "Spinner"
    End Select
End Function

Public Function ShapeTypeString(ByVal l As MsoShapeType) As String
    Select Case l
        Case mso3DModel:            ShapeTypeString = "3D model"
        Case msoAutoShape:          ShapeTypeString = "AutoShape"
        Case msoCallout:            ShapeTypeString = "Callout"
        Case msoCanvas:             ShapeTypeString = "Canvas"
        Case msoChart:              ShapeTypeString = "Chart"
        Case msoComment:            ShapeTypeString = "Comment"
        Case msoContentApp:         ShapeTypeString = "Content Office Add-in"
        Case msoDiagram:            ShapeTypeString = "Diagram"
        Case msoEmbeddedOLEObject:  ShapeTypeString = "Embedded-OLE-Object"
        Case msoFormControl:        ShapeTypeString = "Form-Control"
        Case msoFreeform:           ShapeTypeString = "Freeform"
        Case msoGraphic:            ShapeTypeString = "Graphic"
        Case msoGroup:              ShapeTypeString = "Group"
        Case msoInk:                ShapeTypeString = "Ink"
        Case msoInkComment:         ShapeTypeString = "Ink comment"
        Case msoLine:               ShapeTypeString = "Line"
        Case msoLinked3DModel:      ShapeTypeString = "Linked 3D model"
        Case msoLinkedGraphic:      ShapeTypeString = "Linked graphic"
        Case msoLinkedOLEObject:    ShapeTypeString = "Linked OLE object"
        Case msoLinkedPicture:      ShapeTypeString = "Linked picture"
        Case msoMedia:              ShapeTypeString = "Media"
        Case msoOLEControlObject:   ShapeTypeString = "OLE-Control-Object"
        Case msoPicture:            ShapeTypeString = "Picture"
        Case msoPlaceholder:        ShapeTypeString = "Placeholder"
        Case msoScriptAnchor:       ShapeTypeString = "Script-Anchor"
        Case msoShapeTypeMixed:     ShapeTypeString = "Mixed-Shape-Type"
        Case msoSlicer:             ShapeTypeString = "Slicer"
        Case msoTable:              ShapeTypeString = "Table"
        Case msoTextBox:            ShapeTypeString = "Text-Box"
        Case msoTextEffect:         ShapeTypeString = "Text-Effect"
        Case msoWebVideo:           ShapeTypeString = "Web-Video"
    End Select
End Function

Public Sub ListAscendingByName(ByVal wsh As Worksheet)
    
    Dim shp As Shape
    
    With wsh
        For Each shp In .Shapes
            Debug.Print .Name & "(" & CntrlType(shp) & ")", Tab(45), CntrlName(shp)
        Next shp
    End With
    
End Sub

Public Function ControlExists( _
                        ByRef sync_wb As Workbook, _
                        ByVal sync_sheet_name As String, _
                        ByVal sync_sheet_code_name As String, _
                        ByVal sync_sheet_control_name As String) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE when the control (sync_sheet_control_name) exists in the
' Workbook (sync_wb) in the sheet provided by its Name (sync_sheet_name) or
' its CodeName (sync_sheet_code_name).
' When this function is used to get the required info for being confirmed,
' the concerned sheet may be one of which the Name or the CodeName is about
' to be renamed - which by then will not have taken place yet.
' -------------------------------------------------------------------------
    Const PROC = "ControlExists"
    
    On Error GoTo eh
    Dim ws  As Worksheet
    Dim shp As Shape
    
    For Each ws In sync_wb.Worksheets
        If ws.Name <> sync_sheet_name And ws.CodeName <> sync_sheet_code_name _
        Then GoTo next_sheet
        For Each shp In ws.Shapes
            ControlExists = CntrlName(shp) = sync_sheet_control_name
            If ControlExists Then Exit For
        Next shp
next_sheet:
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSheetControls." & s
End Function

