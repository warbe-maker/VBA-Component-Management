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

Public Property Let ItemsTotal(ByVal i As Long)
    lItemsTotal = i
    Dsply
End Property

Public Property Let Operation(ByVal o As String):   sOperation = o:     End Property

Public Sub Dsply()
    mCompManClient.Progress p_service_name:=sServiceName _
                          , p_serviced_wbk_name:=sWbkServiced _
                          , p_by_servicing_wbk_name:=sWbkServicing _
                          , p_progress_figures:=True _
                          , p_service_op:=sOperation _
                          , p_no_comps_total:=lItemsTotal _
                          , p_no_comps_serviced:=lItemsDone _
                          , p_no_comps_skipped:=lItemsSkipped _
                          , p_service_info:=ItemsDone
End Sub

Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = Not IsError(UBound(a_v))
End Function

Public Sub ItemSkipped()
    lItemsSkipped = lItemsSkipped + 1
    If lItemsSkipped - (Int(lItemsSkipped / 10) * 10) = 0 _
    Then Dsply ' optimizes performance
End Sub

Private Function ItemsDone() As String

    If ArrayIsAllocated(aDone) _
    Then ItemsDone = "(" & Join(aDone, ", ") & ")"
    
End Function

Private Sub Class_Initialize()
    sServiceName = mCompManClient.ServiceName(Services.CurrentService)
    sWbkServiced = Services.ServicedWbk.Name
    sWbkServicing = ThisWorkbook.Name
End Sub
