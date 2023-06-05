VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dctChanged          As Dictionary
Private dctAll              As Dictionary
Private lMaxLenServicedItem As Long
Private lMaxLenServicedType As Long

Private Sub Class_Initialize()
    CollectAll
    Set dctChanged = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dctChanged = Nothing
End Sub

Public Property Get All() As Dictionary:            Set All = dctAll:                           End Property

Public Property Get MaxLenServicedItem() As Long:   MaxLenServicedItem = lMaxLenServicedItem:   End Property

Public Property Get MaxLenServicedType() As Long:   MaxLenServicedType = lMaxLenServicedType:   End Property

Public Function CollectAll(Optional ByRef c_stats As clsStats = Nothing, _
                           Optional ByVal c_wbk As Workbook = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all VBComponent objects as key and the
' corresponding clsComp objects as item.
' Note: The mDct.DctAdd service orders the items ascending by key (possible
'       because the key object has a Name property.
' ----------------------------------------------------------------------------
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    
    If c_wbk Is Nothing Then Set c_wbk = ActiveWorkbook
    Set dctAll = New Dictionary
    
    With c_wbk.VBProject
        For Each vbc In .VBComponents
            If Not c_stats Is Nothing Then c_stats.Count mCompMan.sic_comps
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = c_wbk
                .CompName = vbc.Name
                Set .VBComp = vbc
                lMaxLenServicedItem = Max(lMaxLenServicedItem, Len(.CompName))
                lMaxLenServicedType = Max(lMaxLenServicedType, Len(.TypeString))
            End With
            dctAll.Add vbc.Name, Comp
        Next vbc
    End With
    mDct.KeySort dctAll
    
End Function

Private Function Changed(Optional ByRef c_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary all components of which the code has been modified.
' The Dictionary has the VBComponent objects as key and the corresponding
' clsComp objects as item. The function uses the All Dictionary and removes
' all items where the components not indicates it has changed.
' ----------------------------------------------------------------------------
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    
    Set dct = dctAll(c_stats)
    For Each v In dct
        Set Comp = dct(v)
        If Not Comp.Changed Then
            dct.Remove v
        End If
    Next v
    If Not c_stats Is Nothing Then c_stats.Count mCompMan.sic_comps_changed, dct.Count
    Set Changed = dct
    
End Function

Private Function Outdated(Optional ByRef o_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary of all VBComponents in the mService.Serviced Workbook's
' VBProject indicated Used Common Components of which an update is due because
' the correspondingthe Raw's code has changed.
' The Dictionary has the VBComponent objects as key and the corresponding
' clsComp objects as item. The function uses the All Dictionary and removes
' all items not identified as a Used Common Component of which the correponding
' Raw Common Component has changed.
' ----------------------------------------------------------------------------
    Const PROC = "Outdated"
    
    On Error GoTo eh
    Dim Comp    As clsComp
    Dim dct     As Dictionary
    Dim v       As Variant
    
    Set dct = dctAll
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            If Not .KindOfComp = mCompMan.enCommCompUsed Then
                dct.Remove v
            ElseIf Not .Outdated Then
                dct.Remove v
            End If
        End With
    Next v
    If Not o_stats Is Nothing Then o_stats.Count mCompMan.sic_used_comm_vbc_outdated, dct.Count
    
xt: Set Outdated = dct
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComps" & "." & sProc
End Function

