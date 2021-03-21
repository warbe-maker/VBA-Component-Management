VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private wbSyncSource            As Workbook
Private wbSyncTarget            As Workbook
Private dctSourceSheetComps     As Dictionary
Private dctTargetSheetComps     As Dictionary
Private dctSourceSheets         As Dictionary
Private dctTargetSheets         As Dictionary
Private dctSourceNames          As Dictionary
Private dctTargetNames          As Dictionary
Private dctSourceSheetShapes    As Dictionary
Private dctTargetSheetShapes    As Dictionary
Private dctNewSheets            As Dictionary
Private cllConfInfo             As Collection
Private sEntry                  As String
Private lMaxLenItem             As String
Private lMaxLenTypeItem         As Long

Private Sub Class_Terminate()
    Set dctSourceSheetComps = Nothing
    Set dctTargetSheetComps = Nothing
    Set dctSourceSheets = Nothing
    Set dctTargetSheets = Nothing
    Set dctSourceNames = Nothing
    Set dctTargetNames = Nothing
    Set dctSourceSheetShapes = Nothing
    Set dctTargetSheetShapes = Nothing
    Set cllConfInfo = Nothing
End Sub

Friend Property Get ConfInfo( _
              Optional ByRef serviced_item As Variant = Nothing) As String
' -------------------------------------------------------
' Returns the collected confirmation info as string,
' completed with a summary.
' -------------------------------------------------------
    Dim v As Variant
        
    ConfInfoSummary
    
    '~~ Compile result string
    For Each v In cllConfInfo
        ConfInfo = ConfInfo & v & vbLf
    Next v
    
End Property

Friend Property Let ConfInfo( _
              Optional ByRef serviced_item As Variant = Nothing, _
                       ByVal s As String)
' ----------------------------------------------------------------
' Append a confirmation info to a Collection.
' Note: The clsLog class is used to determine the serviced item's
'       type and name string in order to have the confirmation
'       info consistent with the logged info.
' ----------------------------------------------------------------
    If cllConfInfo Is Nothing Then Set cllConfInfo = New Collection
    With cLog
        If Not TypeName(serviced_item) = "Nothing" Then
            .ServicedItem = serviced_item
        End If
        sEntry = .ServicedItem & s
    End With
    cllConfInfo.Add sEntry
End Property

Friend Property Get MaxLenItem() As Long:
    If lMaxLenItem = 0 Then
        lMaxLenItem = mCompMan.MaxLenComp(Me.Source)
        lMaxLenItem = mCompMan.MaxLenComp(Me.Target)
    End If
    MaxLenItem = lMaxLenItem
End Property

Friend Property Get MaxLenTypeItem() As Long
    MaxLenTypeItem = lMaxLenTypeItem
End Property

Friend Property Let NewSheet( _
              Optional ByVal sheet_code_name As String, _
                       ByVal sheet_name As String)
    If dctNewSheets Is Nothing Then Set dctNewSheets = New Dictionary
    If Not dctNewSheets.Exists(sheet_code_name) Then dctNewSheets.Add sheet_code_name, sheet_name
    If Not dctNewSheets.Exists(sheet_name) Then dctNewSheets.Add sheet_name, sheet_code_name
End Property

Friend Property Get NewSheetExists(ByVal sheet_name As String) As Boolean
    NewSheetExists = dctNewSheets.Exists(sheet_name)
End Property

Friend Property Get Source() As Workbook:               Set Source = wbSyncSource:                      End Property

                 
Friend Property Set Source(ByRef wb As Workbook):       Set wbSyncSource = wb:                          End Property

Friend Property Get SourceNames() As Dictionary:        Set SourceNames = dctSourceNames:               End Property

Friend Property Get SourceSheetComps() As Dictionary:   Set SourceSheetComps = dctSourceSheetComps:     End Property

Friend Property Get SourceSheets() As Dictionary:       Set SourceSheets = dctSourceSheets:             End Property

Friend Property Get SourceSheetShapes() As Dictionary:  Set SourceSheetShapes = dctSourceSheetShapes:   End Property

Friend Property Get Target() As Workbook:               Set Target = wbSyncTarget:                      End Property

Friend Property Set Target(ByRef wb As Workbook):       Set wbSyncTarget = wb:                          End Property

Friend Property Get TargetNames() As Dictionary:        Set TargetNames = dctTargetNames:               End Property

Friend Property Get TargetSheetComps() As Dictionary:   Set TargetSheetComps = dctTargetSheetComps:     End Property

Friend Property Get TargetSheets() As Dictionary:       Set TargetSheets = dctTargetSheets:             End Property

Friend Property Get TargetSheetShapes() As Dictionary:  Set TargetSheetShapes = dctTargetSheetShapes:   End Property

Public Sub CollectAllSyncItems()
    Const PROC = "CollectAllSyncItems"
    
    On Error GoTo eh
    Dim ws1     As Worksheet
    Dim ws2     As Worksheet
    Dim vbc     As VBComponent
    Dim shp     As Shape
    Dim nm      As Name
    Dim nmName  As Name
    Dim sKey    As String
    
    If dctSourceSheets Is Nothing Then Set dctSourceSheets = New Dictionary Else dctSourceSheets.RemoveAll
    If dctSourceSheetComps Is Nothing Then Set dctSourceSheetComps = New Dictionary Else dctSourceSheetComps.RemoveAll
    If dctTargetSheets Is Nothing Then Set dctTargetSheets = New Dictionary Else dctTargetSheets.RemoveAll
    If dctTargetSheetComps Is Nothing Then Set dctTargetSheetComps = New Dictionary Else dctTargetSheetComps.RemoveAll
    If dctSourceSheetShapes Is Nothing Then Set dctSourceSheetShapes = New Dictionary Else dctSourceSheetShapes.RemoveAll
    If dctTargetSheetShapes Is Nothing Then Set dctTargetSheetShapes = New Dictionary Else dctTargetSheetShapes.RemoveAll
    If dctSourceNames Is Nothing Then Set dctSourceNames = New Dictionary Else dctSourceNames.RemoveAll
    If dctTargetNames Is Nothing Then Set dctTargetNames = New Dictionary Else dctTargetNames.RemoveAll
    
    '~~ Source Worksheets by Name and by CodeName
    Stats.Count sic_sheets_total, wbSyncSource.Worksheets.Count
    For Each ws1 In wbSyncSource.Worksheets
        cLog.ServicedItem = ws1 ' Compute max length for logged item type and item name
        With ws1
            mDct.DctAdd dctSourceSheets, .Name, .Name, order_bykey, seq_ascending, , , True
    '        mDct.DctAdd dctSourceSheets, .CodeName, .Name, order_bykey, seq_ascending, , , True
            
            '~~ Source Workseet Shapes
            Stats.Count sic_shapes_total, .Shapes.Count
            For Each shp In .Shapes
                cLog.ServicedItem = shp ' Compute max length for logged item type and item name
                sKey = mSync.KeySheetShape(.Name, shp.Name)
                mDct.DctAdd dctSourceSheetShapes, sKey, .Name, order_bykey, seq_ascending, , , True
            Next shp
        End With
    Next ws1
     
    '~~ Target Worksheets by Name and by CodeName
    For Each ws2 In wbSyncTarget.Worksheets
        cLog.ServicedItem = ws2 ' Compute max length for logged item type and item name
        With ws2
            mDct.DctAdd dctTargetSheets, .Name, .Name, order_bykey, seq_ascending, , , True
'            mDct.DctAdd dctTargetSheets, .CodeName, .Name, order_bykey, seq_ascending, , , True
            
            '~~ Target Worksheet Shapes
            For Each shp In .Shapes
                cLog.ServicedItem = shp ' Compute max length for logged item type and item name
                sKey = mSync.KeySheetShape(.Name, shp.Name)
                mDct.DctAdd dctTargetSheetShapes, sKey, .Name, order_bykey, seq_ascending, , , True
            Next shp
        End With
    Next ws2
    
    '~~ Source and Target VBComponents
    Stats.Count sic_vbcomps_total, wbSyncSource.VBProject.VBComponents.Count
    For Each vbc In wbSyncSource.VBProject.VBComponents
        cLog.ServicedItem = vbc ' Compute max length for logged item type and item name
        If mCompMan.IsSheetComp(vbc) Then
            mDct.DctAdd dctSourceSheetComps, vbc.Name, vbc, order_bykey, seq_ascending, , , True
        End If
    Next vbc
    For Each vbc In wbSyncTarget.VBProject.VBComponents
        cLog.ServicedItem = vbc ' Compute max length for logged item type and item name
        If IsSheetComp(vbc) Then
            mDct.DctAdd dctTargetSheetComps, vbc.Name, vbc, order_bykey, seq_ascending, , , True
        End If
    Next vbc
    
    '~~ Source and Target Workbook Names
    Stats.Count sic_names_total, wbSyncSource.Names.Count
    For Each nm In wbSyncSource.Names
        cLog.ServicedItem = nm ' Compute max length for logged item type and item name
        Set nmName = nm
        mDct.DctAdd dctSourceNames, nm.Name, nmName, order_bykey, seq_ascending, , , True
    Next nm
    For Each nm In wbSyncTarget.Names
        cLog.ServicedItem = nm ' Compute max length for logged item type and item name
        Set nmName = nm
        mDct.DctAdd dctTargetNames, nm.Name, nmName, order_bykey, seq_ascending, , , True
    Next nm

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub ConfInfoClear()
    Set cllConfInfo = Nothing
End Sub

Private Sub ConfInfoSummary()
    Me.ConfInfo("Summary Standard-, Class Modules, UserForms") = _
                Stats.AlignedValue(sic_non_doc_mod_total) & " Total, " & _
                Stats.AlignedValue(sic_non_doc_mod_new) & " New, " & _
                Stats.AlignedValue(sic_non_doc_mod_obsolete) & " Obsolete, " & _
                Stats.AlignedValue(sic_non_doc_mods_code) & " Code-change"
    Me.ConfInfo("Summary Names") = _
                Stats.AlignedValue(sic_names_total) & " Total, " & _
                Stats.AlignedValue(sic_names_new) & " New, " & _
                Stats.AlignedValue(sic_names_obsolete) & " Obsolete"
    Me.ConfInfo("Summary References") = _
                Stats.AlignedValue(sic_refs_total) & " Total, " & _
                Stats.AlignedValue(sic_refs_new) & " New, " & _
                Stats.AlignedValue(sic_refs_obsolete) & " Obsolete"
    Me.ConfInfo("Summary Shapes") = _
                Stats.AlignedValue(sic_shapes_total) & " Total, " & _
                Stats.AlignedValue(sic_shapes_new) & " New, " & _
                Stats.AlignedValue(sic_shapes_obsolete) & " Obsolete"
    Me.ConfInfo("Summary Sheets") = _
                Stats.AlignedValue(sic_sheets_total) & " Total, " & _
                Stats.AlignedValue(sic_sheets_new) & " New, " & _
                Stats.AlignedValue(sic_sheets_obsolete) & " Obsolete, " & _
                Stats.AlignedValue(sic_sheets_code) & " Code-change, " & _
                Stats.AlignedValue(sic_sheets_codename) & " CodeName-change, " & _
                Stats.AlignedValue(sic_sheets_name) & " Name-change"
End Sub

Public Sub DisplayCollectedSourceSheets()
    Dim i   As Long
    Dim ws  As Worksheet
    Dim dct As Dictionary
    
    Debug.Print Align(s:="T/S", lngth:=10, aligned:=AlignCentered) & Align("Key", 30) & "Item"
    Set dct = Me.SourceSheets
    For i = 1 To dct.Count
        Set ws = Me.Source.Worksheets(dct.Items()(i - 1))
        Debug.Print Align("Source", 10) & Align(dct.Keys()(i - 1), 30) & ws.CodeName & "(" & ws.Name & ")"
    Next i
End Sub

Public Sub DisplayCollectedSourceSheetShapes()
    Dim v       As Variant
    Dim sSheet  As String
    Dim sShape  As String
    
    Debug.Print Align(s:="T/S", lngth:=10, aligned:=AlignCentered) & Align("Sheet", 30) & "Shape"
    For Each v In Me.SourceSheetShapes
        sSheet = mSync.KeySheetName(v)
        sShape = mSync.KeyShapeName(v)
        Debug.Print Align("Source", 10) & Align(sSheet, 30) & sShape
    Next v
End Sub

Public Sub DisplayCollectedTargetSheets()
    Dim i   As Long
    Dim ws  As Worksheet
    Dim dct As Dictionary
    
    Debug.Print Align(s:="T/S", lngth:=10, aligned:=AlignCentered) & Align("Key", 30) & "Item"
    Set dct = Me.SourceSheets
    For i = 1 To dct.Count
        Set ws = Me.Source.Worksheets(dct.Items()(i - 1))
        Debug.Print Align("Target", 10) & Align(dct.Keys()(i - 1), 30) & ws.CodeName & "(" & ws.Name & ")"
    Next i
End Sub

Public Sub DisplayCollectedTargetSheetShapes()
    Dim v       As Variant
    Dim sSheet  As String
    Dim sShape  As String
    
    Debug.Print Align(s:="T/S", lngth:=10, aligned:=AlignCentered) & Align("Sheet", 30) & "Shape"
    For Each v In Me.TargetSheetShapes
        sSheet = mSync.KeySheetName(v)
        sShape = mSync.KeyShapeName(v)
        Debug.Print Align("Source", 10) & Align(sSheet, 30) & sShape
    Next v
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsSync" & "." & sProc
End Function

Public Function SheetProjectName( _
                  Optional ByRef ws As Worksheet = Nothing, _
                  Optional ByRef wb As Workbook = Nothing, _
                  Optional ByRef vbc As VBComponent = Nothing) As String
' ----------------------------------------------------------------------
' Returns a Worksheet's full project name in the form 'codename(name)'.
' The Workshhet may either be identified by its Name or through its
' CodeName (wb and vbc).
' ----------------------------------------------------------------------
    If Not ws Is Nothing Then
        SheetProjectName = ws.CodeName & "(" & ws.Name & ")"
    ElseIf Not wb Is Nothing Then
        If Not vbc Is Nothing Then
            For Each ws In wb.Worksheets
                If ws.CodeName = vbc.Name Then
                    SheetProjectName = vbc.Name & "(" & ws.Name & ")"
                    Exit For
                End If
            Next ws
        End If
    End If
End Function

Public Function TypeAndItem(ByVal tai_type As String, ByVal tai_item As String)
    TypeAndItem = Align(tai_type, MAX_LEN_TYPE) & tai_item
End Function

