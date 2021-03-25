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

Friend Property Get Changed() As Dictionary:            Set Changed = dctChanged:   End Property

Friend Property Set Serviced(ByRef wb As Workbook):     Set wbServiced = wb:        End Property

Public Sub CollectAllChanged()
' -------------------------------------------------------------------
' Collect all changed Clone-Components. I.e. all VBComponents of
' which the name is registered as a Raw-Component hosted in another
' Workbook.
' -------------------------------------------------------------------
    Const PROC = "CollectAllChanged"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim cComp   As clsComp
            
    Stats.Count sic_comps_total, wbServiced.VBProject.VBComponents.Count
    
    '~~ Ensure max lenght for the log entries
    For Each vbc In wbServiced.VBProject.VBComponents
        cLog.ServicedItem = vbc
    Next vbc
    
    For Each vbc In wbServiced.VBProject.VBComponents
        Set cComp = New clsComp
        With cComp
            Set .Wrkbk = wbServiced
            .CompName = vbc.Name
            cLog.ServicedItem = vbc
            If Not .KindOfComp = enRawClone Then GoTo next_vbc
            Stats.Count sic_clone_comps
            '~~ This is a Raw-Component's Clone-Component
            Set cRaw = New clsRaw
            cRaw.HostFullName = mHostedRaws.HostFullName(comp_name:=.CompName)
            cRaw.CompName = .CompName
            cRaw.ExpFileExtension = .ExpFileExtension
            cRaw.CloneExpFileFullName = .ExpFileFullName
            cRaw.TypeString = .TypeString
            If cRaw.Changed Then
                Stats.Count sic_raw_changed
                mDct.DctAdd dctChanged, vbc.Name, cComp, order_bykey, seq_ascending, , , True
            Else
                cLog.Entry = "Corresponding Raw's code un-changed."
            End If
        End With
next_vbc:
        Set cComp = Nothing
        Set cRaw = Nothing
    Next vbc
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsClone" & "." & sProc
End Function
