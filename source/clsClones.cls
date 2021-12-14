VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClones"
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

Friend Property Get RawChanged() As Dictionary
    Me.CollectAllChanged
    Set RawChanged = dctChanged
End Property

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

Public Sub CollectAllChanged()
' -----------------------------------------------------
' Collect all changed Clone-Components.
' I.e. all VBComponents of which the name is registered
' as a Raw-Component hosted in another Workbook.
' ------------------------------------------------------
    Const PROC = "CollectAllChanged"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim CloneComp   As clsComp
    Dim sChanged    As String
    Dim v           As Variant
    Dim RawComp     As clsRaw
    
    Stats.Clear
    Stats.Count sic_comps_total, mService.Serviced.VBProject.VBComponents.Count
        
    For Each v In Sorted
        Set vbc = Sorted(v)
        Stats.Count sic_comps
        Set CloneComp = New clsComp
        With CloneComp
            Set .Wrkbk = mService.Serviced
            .CompName = vbc.name
            
            DsplyProgress p_result:=sChanged & " " & vbc.name _
                        , p_total:=Stats.Total(sic_comps_total) _
                        , p_done:=Stats.Total(sic_comps)
            
            Log.ServicedItem = vbc
            If Not .KindOfComp = enRawClone Then GoTo next_vbc
            Stats.Count sic_clone_comps
            '~~ This is a Raw-Component's Clone-Component
            Set RawComp = New clsRaw
            RawComp.HostWrkbkFullName = mRawsHosted.HostFullName(comp_name:=.CompName)
            RawComp.CompName = .CompName
            RawComp.ExpFileExt = .ExpFileExt
            RawComp.CloneExpFileFullName = .ExpFileFullName
            RawComp.TypeString = .TypeString
            If RawComp.Changed(CloneComp) Then
                Stats.Count sic_raw_changed
                mDct.DctAdd dctChanged, vbc.name, RawComp, order_bykey, seq_ascending, , , True
                If Stats.Total(sic_clone_changed) = 1 _
                Then sChanged = vbc.name _
                Else sChanged = sChanged & ", " & vbc.name
            Else
                Log.Entry = "Corresponding Raw's code un-changed."
            End If
        End With
next_vbc:
        Set CloneComp = Nothing
        Set RawComp = Nothing
    Next v
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsClone" & "." & sProc
End Function
