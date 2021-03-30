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
Private RawComp        As clsRaw

Private Sub Class_Initialize()
    Set dctChanged = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dctChanged = Nothing
End Sub

Friend Property Get AllChanged() As Dictionary
    Me.CollectAllChanged
    Set AllChanged = dctChanged
End Property

Private Property Get Sorted() As Dictionary
' ----------------------------------------------------
' Return all components in ascendig order by type.name
' ----------------------------------------------------
    Dim dct As Dictionary
    Dim vbc As VBComponent
    
    For Each vbc In mService.Serviced.VBProject.VBComponents
        Log.ServicedItem = vbc
        mDct.DctAdd dct, vbc.Type & vbc.Name, vbc, order_bykey, seq_ascending, , , True
    Next vbc
    Set Sorted = dct
    Set dct = Nothing
End Property

Public Sub CollectAllChanged()
' --------------------------------------------------------------
' Collect all changed VBComponents.
' I.e. all VBComponents of which a temporary Export-File differs
' from the regular (last) Export-File or which do not have such
' a regular Export-File because they never have been exported
' or the export folder has changed its name and/or location.
' ---------------------------------------------------------------
    Const PROC = "CollectAllChanged"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim Comp       As clsComp
    Dim v           As Variant
    Dim dct         As Dictionary
    Dim sChanged    As String
    Dim CompChanged    As clsComp
    
    Stats.Count sic_comps_total, mService.Serviced.VBProject.VBComponents.Count
    
    Set dct = Sorted
    For Each v In dct
        Set vbc = dct(v)
        DsplyProgress p_result:=sChanged & " " & vbc.Name _
                    , p_total:=Stats.Total(sic_comps_total) _
                    , p_done:=Stats.Total(sic_comps)
        
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = mService.Serviced
            .CompName = vbc.Name
            Log.ServicedItem = vbc
            Stats.Count sic_comps
            If .Changed Then
                Stats.Count sic_comps_changed
                Set CompChanged = New clsComp
                Set CompChanged = Comp
                mDct.DctAdd dctChanged, vbc.Type & vbc.Name, CompChanged, order_bykey, seq_ascending, , , True
                If Stats.Total(sic_comps_changed) = 1 _
                Then sChanged = vbc.Name _
                Else sChanged = sChanged & ", " & vbc.Name
            Else
                Log.Entry = "Code un-changed"
            End If
        End With
next_v:
        Set Comp = Nothing
        Set RawComp = Nothing
    Next v
    
xt: Set dct = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComps" & "." & sProc
End Function

