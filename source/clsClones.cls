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
Private wbServiced  As Workbook
Private cRaw        As clsRaw

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

Friend Property Set Serviced(ByRef wb As Workbook):     Set wbServiced = wb:        End Property

Private Property Get Sorted() As Dictionary
' ----------------------------------------------------
' Return all components in ascendig order by type.name
' ----------------------------------------------------
    Dim dct As Dictionary
    Dim vbc As VBComponent
    
    For Each vbc In wbServiced.VBProject.VBComponents
        Log.ServicedItem = vbc
        mDct.DctAdd dct, vbc.Type & vbc.Name, vbc, order_bykey, seq_ascending, , , True
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
    Dim cComp       As clsComp
    Dim sChanged    As String
    Dim v           As Variant
    
    Stats.Clear
    Stats.Count sic_comps_total, wbServiced.VBProject.VBComponents.Count
        
    For Each v In Sorted
        Set vbc = Sorted(v)
        Stats.Count sic_comps
        Set cComp = New clsComp
        With cComp
            Set .Wrkbk = wbServiced
            .CompName = vbc.Name
            
            DsplyProgress p_result:=sChanged & " " & vbc.Name _
                        , p_total:=Stats.Total(sic_comps_total) _
                        , p_done:=Stats.Total(sic_comps)
            
            Log.ServicedItem = vbc
            If Not .KindOfComp = enRawClone Then GoTo next_vbc
            Stats.Count sic_clone_comps
            '~~ This is a Raw-Component's Clone-Component
            Set cRaw = New clsRaw
            cRaw.HostFullName = mRawsHosted.HostFullName(comp_name:=.CompName)
            cRaw.CompName = .CompName
            cRaw.ExpFileExt = .ExpFileExt
            cRaw.CloneExpFileFullName = .ExpFileFullName
            cRaw.TypeString = .TypeString
            If cRaw.Changed Then
                Set cRaw.CloneVbc = vbc
                Stats.Count sic_raw_changed
                mDct.DctAdd dctChanged, vbc.Name, cRaw, order_bykey, seq_ascending, , , True
                If Stats.Total(sic_clone_changed) = 1 _
                Then sChanged = vbc.Name _
                Else sChanged = sChanged & ", " & vbc.Name
            Else
                Log.Entry = "Corresponding Raw's code un-changed."
            End If
        End With
next_vbc:
        Set cComp = Nothing
        Set cRaw = Nothing
    Next v
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsClone" & "." & sProc
End Function
