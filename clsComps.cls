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
Private wbServiced  As Workbook
Private cRaw        As clsRaw

Private Sub Class_Initialize()
    Set dctChanged = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dctChanged = Nothing
End Sub

Friend Property Get Changed() As Dictionary
    Me.CollectAllChanged
    Set Changed = dctChanged
End Property

Friend Property Set Serviced(ByRef wb As Workbook)
    Set wbServiced = wb
End Property

Private Property Get Sorted() As Dictionary
' ----------------------------------------------------
' Return all components in ascendig order by type.name
' ----------------------------------------------------
    Dim dct As Dictionary
    Dim vbc As VBComponent
    
    For Each vbc In wbServiced.VBProject.VBComponents
        cLog.ServicedItem = vbc
        mDct.DctAdd dct, vbc.Type & vbc.Name, vbc, order_bykey, seq_ascending, , , True
    Next vbc
    Set Sorted = dct
    Set dct = Nothing
End Property

Public Sub CollectAllChanged()
' -------------------------------------------------------------------
' Collect all changed Clone-Components. I.e. all VBComponents of
' which the name is registered as a Raw-Component hosted in another
' Workbook.
' -------------------------------------------------------------------
    Const PROC = "CollectAllChanged"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim cComp       As clsComp
    Dim v           As Variant
    Dim dct         As Dictionary
    
    Stats.Count sic_comps_total, wbServiced.VBProject.VBComponents.Count
    
    Set dct = Sorted
    For Each v In dct
        Set vbc = dct(v)
        Set cComp = New clsComp
        With cComp
            Set .Wrkbk = wbServiced
            .CompName = vbc.Name
            cLog.ServicedItem = vbc
            Stats.Count sic_comps
            If .Changed Then
                Stats.Count sic_comps_changed
                mDct.DctAdd dctChanged, vbc.Type & vbc.Name, cComp, order_bykey, seq_ascending, , , True
            Else
                cLog.Entry = "Code un-changed"
            End If
        End With
next_v:
        Set cComp = Nothing
        Set cRaw = Nothing
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

