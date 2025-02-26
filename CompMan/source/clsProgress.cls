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

Private aDone           As Variant
Private lItemsDone      As Long
Private lItemsSkipped   As Long
Private lItemsTotal     As Long
Private sOperation      As String
Private sServiceName    As String
Private sWbkServiced    As String
Private sWbkServicing   As String
Private bFigures        As Boolean
Private bDoneItemsInfo  As Boolean

Private Sub Class_Initialize()
    sServiceName = Serviced.ServiceName
    sWbkServiced = Serviced.Wrkbk.Name
    sWbkServicing = ThisWorkbook.Name
End Sub

Public Property Let DoneItemsInfo(ByVal d_b As Boolean):            bDoneItemsInfo = d_b:           End Property

Public Property Let Figures(ByVal f_b As Boolean):                  bFigures = f_b:                 End Property

Public Property Let ItemDone(Optional ByVal i_sequ_no As Long = 0, _
                                      ByVal i_name As String)

    mBasic.Arry(aDone) = i_name
    lItemsDone = UBound(aDone) + 1
    Dsply i_sequ_no
    
End Property

Public Property Get ItemsDone() As Long:                            ItemsDone = lItemsDone:         End Property

Public Property Get ItemsSkipped() As Long:                         ItemsSkipped = lItemsSkipped:   End Property

Public Property Let ItemsTotal(ByVal i As Long)
    lItemsTotal = i
    Dsply
End Property

Public Property Let Operation(ByVal o As String)
    If o <> sOperation Then ClearDone
    sOperation = o
End Property

Friend Property Let ServiceName(ByVal s_service_name As String):    sServiceName = s_service_name:  End Property

Private Function ArrayErase(ByRef c_arr As Variant)
    If IsArray(c_arr) Then Erase c_arr
End Function

Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = UBound(a_v) >= LBound(a_v)
End Function

Public Sub ClearDone():
    ArrayErase aDone
    lItemsDone = 0
    lItemsSkipped = 0
    lItemsTotal = 0
End Sub

Public Sub Dsply(Optional ByVal d_sequ_no As Long)
    mCompManClient.Progress p_service_name:=sServiceName _
                          , p_serviced_wbk_name:=sWbkServiced _
                          , p_by_servicing_wbk_name:=sWbkServicing _
                          , p_progress_figures:=bFigures _
                          , p_service_op:=sOperation _
                          , p_no_comps_total:=lItemsTotal _
                          , p_no_comps_serviced:=lItemsDone _
                          , p_no_comps_skipped:=lItemsSkipped _
                          , p_service_info:=ItemsDoneInfo _
                          , p_sequ_no:=d_sequ_no
End Sub

Private Function ItemsDoneInfo() As String
    If bDoneItemsInfo Then
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

