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

Private dctChanged  As Dictionary

Private Sub Class_Initialize()
    Set dctChanged = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dctChanged = Nothing
End Sub

Public Function All(Optional ByRef a_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all VBComponent objects as key and the
' corresponding clsComp objects as item.
' Note: The mDct.DctAdd service orders the items ascending by key (possible
'       because the key object has a Name property.
' ----------------------------------------------------------------------------
    Dim wbk     As Workbook
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    
    Set wbk = mService.WbkServiced
    With wbk.VBProject
        For Each vbc In .VBComponents
            If Not a_stats Is Nothing Then a_stats.Count sic_comps
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wbk
                .CompName = vbc.Name
                Set .VBComp = vbc
            End With
            mDct.DctAdd add_dct:=All, add_key:=vbc, add_item:=Comp, add_order:=order_bykey
        Next vbc
    End With
    
End Function

Public Function Changed(Optional ByRef c_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary all components of which the code has been modified.
' The Dictionary has the VBComponent objects as key and the corresponding
' clsComp objects as item. The function uses the All Dictionary and removes
' all items where the components not indicates it has changed.
' ----------------------------------------------------------------------------
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    
    Set dct = Me.All(c_stats)
    For Each v In dct
        Set Comp = dct(v)
        If Not Comp.Changed Then
            dct.Remove v
        End If
    Next v
    If Not c_stats Is Nothing Then c_stats.Count sic_comps_changed, dct.Count
    Set Changed = dct
    
End Function

Public Function Outdated(Optional ByRef o_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary of all VBComponents in the mService.WbkServiced Workbook's
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
    
    Set dct = Me.All
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            If Not .KindOfComp = enCommCompUsed Then
                dct.Remove v
            ElseIf Not .Outdated Then
                dct.Remove v
            End If
        End With
    Next v
    If Not o_stats Is Nothing Then o_stats.Count sic_used_comm_vbc_Outdated, dct.Count
    
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

