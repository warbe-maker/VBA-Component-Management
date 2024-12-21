VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBEMenuCbbOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' To be used by the caller via: Set CmdBarButton = myButton
' ----------------------------------------------------------------------------
Public WithEvents CmdReleaseComp    As CommandBarButton
Attribute CmdReleaseComp.VB_VarHelpID = -1
Public WithEvents CmdRelease        As CommandBarButton
Attribute CmdRelease.VB_VarHelpID = -1

Private Sub CmdRelease_Click(ByVal m_item As Office.CommandBarButton, _
                             ByRef CancelDefault As Boolean)
    mRelease.ReleaseService
        
End Sub

Private Sub CmdReleaseComp_Click(ByVal m_item As Office.CommandBarButton, _
                                 ByRef CancelDefault As Boolean)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cbb     As Office.CommandBarButton
    Dim sComp   As String
    
    mCompMan.ServicedWrkbk = ActiveWorkbook
    mEnvironment.Provide False
    If CommonPending Is Nothing Then
        Set CommonPending = New clsCommonPending
    End If
    
    Set cbb = m_item
    sComp = Split(cbb.Caption, " ")(1)
    CommonPending.ReleaseComp sComp
    CommCompsPendingRelease.Remove sComp
    mCompManMenuVBE.Setup
    
End Sub
