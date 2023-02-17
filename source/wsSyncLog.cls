VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsSyncLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Property Get CompsDone() As Boolean:         CompsDone = DoneYN(SummaryDone("VBComponent")):            End Property

Public Property Let CompsDone(ByVal b As Boolean):  SummaryDone("VBComponent") = DoneYN(b):                    End Property

Public Property Get NamesAdded() As Long:           NamesAdded = Me.Range("celNamesAdded").Value:               End Property

Public Property Get NamesChanged() As Long:         NamesChanged = Me.Range("celNamesChanged").Value:           End Property

Public Property Get NamesDone() As Boolean:         NamesDone = DoneYN(SummaryDone("Name")):                    End Property

Public Property Let NamesDone(ByVal b As Boolean):  SummaryDone("Name") = DoneYN(b):                            End Property

Public Property Get NamesRemoved() As Long:         NamesRemoved = Me.Range("celNamesRemoved").Value:           End Property

Public Property Get RefsDone() As Boolean:          RefsDone = DoneYN(SummaryDone("Reference")):                End Property

Public Property Let RefsDone(ByVal b As Boolean):   SummaryDone("Reference") = DoneYN(b):                       End Property

Public Property Get ShapesDone() As Boolean:        ShapesDone = DoneYN(SummaryDone("Shape")):                  End Property

Public Property Let ShapesDone(ByVal b As Boolean): SummaryDone("Shape") = DoneYN(b):                           End Property

Public Property Get SheetsDone() As Boolean:        SheetsDone = DoneYN(SummaryDone("Worksheet")):              End Property

Public Property Let SheetsDone(ByVal b As Boolean): SummaryDone("Worksheet") = DoneYN(b):                       End Property

Public Property Get SummaryDone(Optional ByVal ds_type As String) As Variant
    Dim rng As Range
    For Each rng In Me.Range("rngSummaryType").Cells
        If rng.Value = ds_type Then
            If rng.Offset(0, 1).Value = "done" Then SummaryDone = True
            Exit For
        End If
    Next rng
End Property

Public Property Let SummaryDone(Optional ByVal ds_type As String, _
                                         ByVal ds_done As Variant)
    Dim rng As Range
    For Each rng In Me.Range("rngSummaryType").Cells
        If rng.Value = ds_type Then
            If VarType(ds_done) = vbBoolean Then
                If ds_done Then
                    rng.Offset(0, 1).Value = "done"
                Else
                    rng.Offset(0, 1).Value = "due"
                End If
            Else
                rng.Offset(0, 1).Value = ds_done
            End If
            Exit For
        End If
    Next rng
    
End Property

Public Sub Clear()
    Dim cel As Range
    Me.Range("rngDone").ClearContents
    For Each cel In Me.Range("rngSummaryStatus").Cells
        cel.Value = "'-"
    Next cel
    Me.Range("rngShapePropertiesUpdated").ClearContents
End Sub

Public Sub Done(ByVal d_prefix As String, _
                ByVal d_type As String, _
                ByVal d_item As String, _
                ByVal d_action As String, _
       Optional ByVal d_details As String = vbNullString, _
       Optional ByVal d_log_item As Variant = Nothing)
' ----------------------------------------------------------------------------
' Write a done synchronization.
' ----------------------------------------------------------------------------
    Const PROC = "Done"
    
    mBasic.BoP ErrSrc(PROC)
    With FirstFreeInCol("colSyncItemType")
        .Value = d_type
        .Offset(0, -1).Value = d_prefix
        .Offset(0, 1).Value = d_item
        .Offset(0, 2).Value = d_action
        .Offset(0, 3).Value = d_details
    End With
    
'    '~~ Write to log file
'    If Not d_log_item Is Nothing Then mService.Log.ServicedItem = d_log_item
'    mService.Log.Entry = d_details
    
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Function DoneYN(ByVal v As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns True when (v) = "yes" and False when (v) = "no"
' Returns "yes" when (v) = True and "no" when (v) = False.
' ----------------------------------------------------------------------------
    If VarType(v) = vbBoolean Then
        If v Then DoneYN = "yes" Else DoneYN = "no"
    Else
        If v = "yes" Then DoneYN = True Else DoneYN = False
    End If
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wsSyncLog" & "." & sProc
End Function

Private Function FirstFreeInCol(ByVal ff_rng As String) As Range
' ----------------------------------------------------------------------------
' Returns the first unused cell in column (ff_col)
' ----------------------------------------------------------------------------
    Const PROC = "FirstFreeInCol"
    
    Dim rng As Range
    
    mBasic.BoP ErrSrc(PROC)
    With Me
        For Each rng In Intersect(.Range(ff_rng).EntireColumn, .Range("rngDone").EntireRow).Cells
            If rng.Value = vbNullString Then
                Set FirstFreeInCol = rng
                GoTo xt
            End If
        Next rng
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

End Function

Public Sub SummaryDoneShapeProperties(ByVal sp_property As String)
    With FirstFreeInCol("rngShapePropertiesUpdated")
        .Value = sp_property
        .EntireColumn.AutoFit
    End With
End Sub

Private Sub Worksheet_Activate()
    Me.Range("rngLogDetails").EntireColumn.AutoFit
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
