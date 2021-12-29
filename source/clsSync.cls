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

Private bAmbigous               As Boolean          ' when True sync is only done when the below is confirmed True
Private bRestrictRenameAsserted As Boolean          ' when False a sheet's CodeName a n d its Name may be changed at once
Private cllConfInfo             As Collection
Private dctChanged              As Dictionary       ' Confirm buttons and clsRaw items to display changed
Private dctNameChange           As Dictionary
Private dctNewSheets            As Dictionary
Private dctSourceNames          As Dictionary
Private dctSourceSheetOOBs      As Dictionary
Private dctSourceSheetComps     As Dictionary
Private dctSourceSheets         As Dictionary
Private dctSourceSheetControls  As Dictionary
Private dctTargetNames          As Dictionary
Private dctTargetSheetOOBs      As Dictionary
Private dctTargetSheetComps     As Dictionary
Private dctTargetSheets         As Dictionary
Private dctTargetSheetControls  As Dictionary
Private lMaxLenItem             As String
Private lMaxLenTypeItem         As Long
Private lMode                   As SyncMode
Private lSheetsNew              As Long
Private lSheetsObsolete         As Long
Private sEntry                  As String
Private wbSyncSource            As Workbook
Private wbSyncTarget            As Workbook
Private bManualSynchRequired    As Boolean

Private Sub Class_Initialize()
    Set dctNameChange = New Dictionary
    Set dctChanged = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set cllConfInfo = Nothing
    Set dctChanged = Nothing
    Set dctNameChange = Nothing
    Set dctNewSheets = Nothing
    Set dctSourceNames = Nothing
    Set dctSourceSheetComps = Nothing
    Set dctSourceSheets = Nothing
    Set dctSourceSheetControls = Nothing
    Set dctTargetNames = Nothing
    Set dctTargetSheetComps = Nothing
    Set dctTargetSheets = Nothing
    Set dctTargetSheetControls = Nothing
    Set dctTargetSheetOOBs = Nothing
    Set dctSourceSheetOOBs = Nothing
End Sub

Friend Property Let ManualSynchRequired(ByVal b As Boolean):    bManualSynchRequired = b:                           End Property
Friend Property Get ManualSynchRequired() As Boolean:           ManualSynchRequired = bManualSynchRequired:         End Property
Friend Property Get Ambigous() As Boolean:                      Ambigous = bAmbigous:                               End Property

Friend Property Let Ambigous(ByVal b As Boolean):               bAmbigous = b:                                      End Property

Public Property Get Changed() As Dictionary
    Set Changed = dctChanged
End Property

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
    With Log
        If Not TypeName(serviced_item) = "Nothing" Then
            .ServicedItem = serviced_item
        End If
        sEntry = .ServicedItem & s
    End With
    cllConfInfo.Add sEntry
End Property

Friend Property Get Mode() As SyncMode:         Mode = lMode:   End Property

Friend Property Let Mode(ByVal l As SyncMode)
    lMode = l
    If lMode = Count Then
        lSheetsNew = 0
        lSheetsObsolete = 0
    End If
End Property

Friend Property Get NameChange() As Dictionary:                 Set NameChange = dctNameChange:                     End Property

Friend Property Let NewSheet( _
              Optional ByVal sheet_code_name As String, _
                       ByVal sheet_name As String)
    If dctNewSheets Is Nothing Then Set dctNewSheets = New Dictionary
    If Not dctNewSheets.Exists(sheet_code_name) Then dctNewSheets.Add sheet_code_name, sheet_name
    If Not dctNewSheets.Exists(sheet_name) Then dctNewSheets.Add sheet_name, sheet_code_name
End Property

Friend Property Get IsNewSheet(ByVal sheet_name As String) As Boolean
    IsNewSheet = dctNewSheets.Exists(sheet_name)
End Property

Friend Property Get RestrictRenameAsserted() As Boolean:        RestrictRenameAsserted = bRestrictRenameAsserted:       End Property

Friend Property Let RestrictRenameAsserted(ByVal b As Boolean): bRestrictRenameAsserted = b:                            End Property

Friend Property Get SheetsNew() As Long:                        SheetsNew = lSheetsNew:                                 End Property

Friend Property Get SheetsObsolete() As Long:                   SheetsObsolete = lSheetsObsolete:                       End Property

Friend Property Get Source() As Workbook:                       Set Source = wbSyncSource:                              End Property

Friend Property Set Source(ByRef wb As Workbook):               Set wbSyncSource = wb:                                  End Property

Friend Property Get SourceNames() As Dictionary:                Set SourceNames = dctSourceNames:                       End Property

Friend Property Get SourceSheetComps() As Dictionary:           Set SourceSheetComps = dctSourceSheetComps:             End Property

Friend Property Get SourceSheets() As Dictionary:               Set SourceSheets = dctSourceSheets:                     End Property

Friend Property Get SourceSheetControls() As Dictionary:        Set SourceSheetControls = dctSourceSheetControls:       End Property

Friend Property Get SourceSheetOOBs() As Dictionary:            Set SourceSheetOOBs = dctSourceSheetOOBs:               End Property

Friend Property Get TargetSheetOOBs() As Dictionary:            Set TargetSheetOOBs = dctTargetSheetOOBs:               End Property

Friend Property Get Target() As Workbook:                       Set Target = wbSyncTarget:                              End Property

Friend Property Set Target(ByRef wb As Workbook):               Set wbSyncTarget = wb:                                  End Property

Friend Property Get TargetNames() As Dictionary:                Set TargetNames = dctTargetNames:                       End Property

Friend Property Get TargetSheetComps() As Dictionary:           Set TargetSheetComps = dctTargetSheetComps:             End Property

Friend Property Get TargetSheets() As Dictionary:               Set TargetSheets = dctTargetSheets:                     End Property

Friend Property Get TargetSheetControls() As Dictionary:        Set TargetSheetControls = dctTargetSheetControls:       End Property

Public Sub ChangedClear()
    If dctChanged Is Nothing Then Set dctChanged = New Dictionary Else dctChanged.RemoveAll
End Sub

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
    Dim oob     As OLEObject
    
    If dctSourceSheets Is Nothing Then Set dctSourceSheets = New Dictionary Else dctSourceSheets.RemoveAll
    If dctSourceSheetComps Is Nothing Then Set dctSourceSheetComps = New Dictionary Else dctSourceSheetComps.RemoveAll
    If dctTargetSheets Is Nothing Then Set dctTargetSheets = New Dictionary Else dctTargetSheets.RemoveAll
    If dctTargetSheetComps Is Nothing Then Set dctTargetSheetComps = New Dictionary Else dctTargetSheetComps.RemoveAll
    If dctSourceSheetControls Is Nothing Then Set dctSourceSheetControls = New Dictionary Else dctSourceSheetControls.RemoveAll
    If dctTargetSheetControls Is Nothing Then Set dctTargetSheetControls = New Dictionary Else dctTargetSheetControls.RemoveAll
    If dctSourceNames Is Nothing Then Set dctSourceNames = New Dictionary Else dctSourceNames.RemoveAll
    If dctTargetNames Is Nothing Then Set dctTargetNames = New Dictionary Else dctTargetNames.RemoveAll
    If dctSourceSheetOOBs Is Nothing Then Set dctSourceSheetOOBs = New Dictionary Else dctSourceSheetOOBs.RemoveAll
    If dctTargetSheetOOBs Is Nothing Then Set dctTargetSheetOOBs = New Dictionary Else dctTargetSheetOOBs.RemoveAll
    
    '~~ Source Worksheets by Name and by CodeName  and the sheet's FormControls and OLEObjects
    Stats.Count sic_sheets_total, wbSyncSource.Worksheets.Count
    For Each ws1 In wbSyncSource.Worksheets
        Log.ServicedItem = ws1 ' Compute max length for logged item type and item name
        With ws1
            mDct.DctAdd dctSourceSheets, .Name, .Name, order_bykey, seq_ascending, , , True
            Stats.Count sic_sheet_controls_total, .Shapes.Count
            CollectSheetControls ws1, dctSourceSheetControls
        End With
    Next ws1
    
    '~~ Target Worksheets by Name and by CodeName and the sheet's FormControls and OLEObjects
    For Each ws2 In wbSyncTarget.Worksheets
        Log.ServicedItem = ws2 ' Compute max length for logged item type and item name
        With ws2
            mDct.DctAdd dctTargetSheets, .Name, .Name, order_bykey, seq_ascending, , , True
            Stats.Count sic_sheet_controls_total, .Shapes.Count
            CollectSheetControls ws2, dctTargetSheetControls
        End With
    Next ws2
    
    '~~ Source and Target VBComponents
    Stats.Count sic_comps_total, wbSyncSource.VBProject.VBComponents.Count
    For Each vbc In wbSyncSource.VBProject.VBComponents
        Log.ServicedItem = vbc ' Compute max length for logged item type and item name
        If mComp.IsSheetDocMod(vbc) Then
            mDct.DctAdd dctSourceSheetComps, vbc.Name, vbc, order_bykey, seq_ascending, , , True
        End If
    Next vbc
    For Each vbc In wbSyncTarget.VBProject.VBComponents
        Log.ServicedItem = vbc ' Compute max length for logged item type and item name
        If mComp.IsSheetDocMod(vbc) Then
            mDct.DctAdd dctTargetSheetComps, vbc.Name, vbc, order_bykey, seq_ascending, , , True
        End If
    Next vbc
    
    '~~ Source and Target Workbook Names
    Stats.Count sic_names_total, wbSyncSource.Names.Count
    For Each nm In wbSyncSource.Names
        Log.ServicedItem = nm ' Compute max length for logged item type and item name
        Set nmName = nm
        mDct.DctAdd dctSourceNames, nm.Name, nmName, order_bykey, seq_ascending, , , True
    Next nm
    For Each nm In wbSyncTarget.Names
        Log.ServicedItem = nm ' Compute max length for logged item type and item name
        Set nmName = nm
        mDct.DctAdd dctTargetNames, nm.Name, nmName, order_bykey, seq_ascending, , , True
    Next nm

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CollectSheetControls( _
                           ByRef ws As Worksheet, _
                           ByRef dct As Dictionary)
' -------------------------------------------------------
'
' -------------------------------------------------------
    Const PROC = "CollectSheetControls"
    
    On Error GoTo eh
    Dim shp     As Shape
    Dim sName   As String
    Dim sKey    As String
    
    For Each shp In ws.Shapes
        Log.ServicedItem = shp ' Compute max length for logged item type and item name
        sName = mSheetControls.CntrlName(shp)
        sKey = mSync.KeySheetControl(sheet_name:=ws.Name, control_name:=sName)
'        Debug.Print "Key='" & sKey & "'", Tab(50), "Item='" & sName & "'"
        mDct.DctAdd dct, sKey, sName, order_bykey, seq_ascending, , , True
    Next shp

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
                Stats.AlignedValue(sic_sheet_controls_total) & " Total, " & _
                Stats.AlignedValue(sic_sheet_controls_new) & " New, " & _
                Stats.AlignedValue(sic_sheet_controls_obsolete) & " Obsolete"
    Me.ConfInfo("Summary Sheets") = _
                Stats.AlignedValue(sic_sheets_total) & " Total, " & _
                Stats.AlignedValue(sic_sheets_new) & " New, " & _
                Stats.AlignedValue(sic_sheets_obsolete) & " Obsolete, " & _
                Stats.AlignedValue(sic_sheets_code) & " Code-change, " & _
                Stats.AlignedValue(sic_sheets_codename) & " CodeName-change, " & _
                Stats.AlignedValue(sic_sheets_name) & " Name-change"
End Sub

Public Sub CountSheetsNew():                    lSheetsNew = lSheetsNew + 1:            End Sub

Public Sub CountSheetsObsolete():               lSheetsObsolete = lSheetsObsolete + 1:  End Sub

Public Sub DisplayCollectedSourceSheets()
    Dim i   As Long
    Dim ws  As Worksheet
    Dim dct As Dictionary
    
    Debug.Print Align(align_s:="T/S", align_lngth:=10, align_mode:=AlignCentered) & Align("Key", 30) & "Item"
    Set dct = Me.SourceSheets
    For i = 1 To dct.Count
        Set ws = Me.Source.Worksheets(dct.Items()(i - 1))
        Debug.Print Align("Source", 10) & Align(dct.Keys()(i - 1), 30) & ws.CodeName & "(" & ws.Name & ")"
    Next i
End Sub

Public Sub DisplayCollectedSourceSheetControls()
    Dim v           As Variant
    Dim sSheet      As String
    Dim sControl    As String
    
    Debug.Print Align(align_s:="T/S", align_lngth:=10, align_mode:=AlignCentered) & Align("Sheet", 30) & "Shape"
    For Each v In Me.SourceSheetControls
        sSheet = mSync.KeySheetName(v)
        sControl = mSync.KeyControlName(v)
        Debug.Print Align("Source", 10) & Align(sSheet, 30) & sControl
    Next v
End Sub

Public Sub DisplayCollectedTargetSheets()
    Dim i   As Long
    Dim ws  As Worksheet
    Dim dct As Dictionary
    
    Debug.Print Align(align_s:="T/S", align_lngth:=10, align_mode:=AlignCentered) & Align("Key", 30) & "Item"
    Set dct = Me.SourceSheets
    For i = 1 To dct.Count
        Set ws = Me.Source.Worksheets(dct.Items()(i - 1))
        Debug.Print Align("Target", 10) & Align(dct.Keys()(i - 1), 30) & ws.CodeName & "(" & ws.Name & ")"
    Next i
End Sub

Public Sub DisplayCollectedTargetSheetControls()
    Dim v           As Variant
    Dim sSheet      As String
    Dim sControl    As String
    
    Debug.Print Align(align_s:="T/S", align_lngth:=10, align_mode:=AlignCentered) & Align("Sheet", 30) & "Shape"
    For Each v In Me.TargetSheetControls
        sSheet = mSync.KeySheetName(v)
        sControl = mSync.KeyControlName(v)
        Debug.Print Align("Source", 10) & Align(sSheet, 30) & sControl
    Next v
End Sub

Public Sub DisplayCollectedTargetSheetOOBs()
    Dim v           As Variant
    Dim sSheet      As String
    Dim sControl    As String
    
    Debug.Print Align(align_s:="T/S", align_lngth:=10, align_mode:=AlignCentered) & Align("Sheet", 30) & "Shape"
    For Each v In Me.TargetSheetOOBs
        sSheet = mSync.KeySheetName(v)
        sControl = mSync.KeyControlName(v)
        Debug.Print Align("Target", 10) & Align(sSheet, 30) & sControl
    Next v
End Sub

Public Sub DisplayCollectedSourceSheetOOBs()
    Dim v           As Variant
    Dim sSheet      As String
    Dim sControl    As String
    
    Debug.Print Align(align_s:="T/S", align_lngth:=10, align_mode:=AlignCentered) & Align("Sheet", 30) & "Shape"
    For Each v In Me.SourceSheetOOBs
        sSheet = mSync.KeySheetName(v)
        sControl = mSync.KeyControlName(v)
        Debug.Print Align("Source", 10) & Align(sSheet, 30) & sControl
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

