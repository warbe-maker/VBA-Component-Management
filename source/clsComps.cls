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
Private RawComp     As clsRaw

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
    Dim wb      As Workbook
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    
    Set wb = mService.Serviced
    With wb.VBProject
        For Each vbc In .VBComponents
            If Not a_stats Is Nothing Then a_stats.Count sic_comps
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wb
                .CompName = vbc.name
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
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    Dim dct     As Dictionary
    Dim v       As Variant
    
    Set dct = Me.All
    For Each v In dct
        Set Comp = dct(v)
        If Not Comp.KindOfComp = enCommCompUsed Then
            dct.Remove v
        Else
            If Not Comp.Outdated Then
                dct.Remove v
            End If
        End If
    Next v
    Set Outdated = dct
    If Not o_stats Is Nothing Then o_stats.Count sic_used_comm_comp_Outdated, dct.Count
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function CommonUsed(Optional ByRef c_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all componentws in the mService.Serviced Workbook's
' VBProject which are indicated a Used Common Component.
' ----------------------------------------------------------------------------
    Const PROC = "CommonUsed"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lRemaining  As Long
    Dim dctComps    As Dictionary
    Dim sMsg        As String
    Dim Log         As New clsLog
    Dim Comp        As clsComp
    Dim lUsed       As Long
    Dim lComps      As Long
    Dim Comps       As New clsComps
    
    Log.Service() = PROC
    Set dctComps = Me.All(c_stats)
    lRemaining = dctComps.Count
    
    For Each v In dctComps
        lComps = lComps + 1
        Set Comp = dctComps(v)
        If Not Comp.KindOfComp = enCommCompUsed Then
            dctComps.Remove v
        End If
        
        lRemaining = lRemaining - 1
        sMsg = Log.Service
        sMsg = sMsg & lUsed & " of " & lComps
        sMsg = sMsg & " " & String(lComps, ".")
        If Len(sMsg) > 255 Then
            sMsg = Left(sMsg, 251) & " ..."
        End If
        Application.StatusBar = sMsg
    
    Next v
    Set CommonUsed = dctComps
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Property Get Sorted() As Dictionary
' ----------------------------------------------------
' Return all components in ascendig order by type.name
' ----------------------------------------------------
    Dim dct As Dictionary
    Dim vbc As VBComponent
    
    For Each vbc In mService.Serviced.VBProject.VBComponents
        Log.ServicedItem = vbc
        mDct.DctAdd dct, vbc.Type & vbc.name, vbc, order_bykey, seq_ascending, , , True
    Next vbc
    Set Sorted = dct
    Set dct = Nothing
End Property

'Public Sub CollectAllChanged()
'' --------------------------------------------------------------
'' Collect all changed VBComponents.
'' I.e. all VBComponents of which a temporary Export-File differs
'' from the regular (last) Export-File or which do not have such
'' a regular Export-File because they never have been exported
'' or the export folder has changed its name and/or location.
'' ---------------------------------------------------------------
'    Const PROC = "CollectAllChanged"
'
'    On Error GoTo eh
'    Dim vbc         As VBComponent
'    Dim Comp        As clsComp
'    Dim v           As Variant
'    Dim dct         As Dictionary
'    Dim sChanged    As String
'    Dim CompChanged As clsComp
'
'    Stats.Count sic_comps_total, mService.Serviced.VBProject.VBComponents.Count
'
'    Set dct = Sorted
'    For Each v In dct
'        Set vbc = dct(v)
'        mCompMan.DsplyProgress p_result:=sChanged & " " & vbc.name _
'                    , p_total:=Stats.Total(sic_comps_total) _
'                    , p_done:=Stats.Total(sic_comps)
'
'        Set Comp = New clsComp
'        With Comp
'            Set .Wrkbk = mService.Serviced
'            .CompName = vbc.name
'            Log.ServicedItem = vbc
'            Stats.Count sic_comps
'            If .Changed Then
'                Stats.Count sic_comps_changed
'                Set CompChanged = New clsComp
'                Set CompChanged = Comp
'                mDct.DctAdd dctChanged, vbc.name, CompChanged, order_bykey, seq_ascending, , , True
'                If Stats.Total(sic_comps_changed) = 1 _
'                Then sChanged = vbc.name _
'                Else sChanged = sChanged & ", " & vbc.name
'            Else
'                Log.Entry = "Code un-changed"
'            End If
'        End With
'next_v:
'        Set Comp = Nothing
'        Set RawComp = Nothing
'    Next v
'
'xt: Set dct = Nothing
'    Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComps" & "." & sProc
End Function

