VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aDone()         As String
Private lItemsDone      As Long
Private lItemsSkipped   As Long
Private lItemsTotal     As Long
Private sOperation      As String
Private sDone           As String
Private sServiceName    As String
Private sWbkServiced    As String
Private sWbkServicing   As String
Private bFigures        As Boolean
Private bDoneItems      As Boolean

Private Sub Class_Initialize()
    sServiceName = mCompManClient.ServiceName(Services.CurrentService)
    sWbkServiced = Services.ServicedWbk.Name
    sWbkServicing = ThisWorkbook.Name
End Sub

Public Property Let DoneItems(ByVal d_b As Boolean):    bDoneItems = d_b:               End Property

Public Property Let Figures(ByVal f_b As Boolean):      bFigures = f_b:                 End Property

Public Property Let ItemDone(ByVal i_name As String)
    If ArrayIsAllocated(aDone) Then
        ReDim Preserve aDone(UBound(aDone) + 1)
    Else
        ReDim aDone(0)
    End If
    aDone(UBound(aDone)) = i_name
    lItemsDone = UBound(aDone) + 1
    Dsply
    
End Property

Public Property Get ItemsSkipped() As Long:             ItemsSkipped = lItemsSkipped:   End Property
Public Property Get ItemsDone() As Long:                ItemsDone = lItemsDone:         End Property

Public Property Let ItemsTotal(ByVal i As Long)
    lItemsTotal = i
    Dsply
End Property

Public Property Let Operation(ByVal o As String):       sOperation = o:                 End Property

Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = Not IsError(UBound(a_v))
End Function

Public Sub Dsply()
    mCompManClient.Progress p_service_name:=sServiceName _
                          , p_serviced_wbk_name:=sWbkServiced _
                          , p_by_servicing_wbk_name:=sWbkServicing _
                          , p_progress_figures:=bFigures _
                          , p_service_op:=sOperation _
                          , p_no_comps_total:=lItemsTotal _
                          , p_no_comps_serviced:=lItemsDone _
                          , p_no_comps_skipped:=lItemsSkipped _
                          , p_service_info:=ItemsDoneInfo
End Sub

Private Function ItemsDoneInfo() As String
    If bDoneItems Then
        If ArrayIsAllocated(aDone) _
        Then ItemsDoneInfo = "(" & Join(aDone, ", ") & ")"
    End If
End Function

Public Sub ItemSkipped()
    lItemsSkipped = lItemsSkipped + 1
'    If lItemsSkipped - (Int(lItemsSkipped / 10) * 10) = 0 Then
        Dsply ' optimizes performance
'    End If
End Sub

