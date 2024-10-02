VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBEMenuCbbVBEVB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cbeRelease       As CommandBarEvents
Attribute cbeRelease.VB_VarHelpID = -1
Private WithEvents cbeReleaseComp   As CommandBarEvents
Attribute cbeReleaseComp.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Set cbeRelease = Nothing
    Set cbeReleaseComp = Nothing
End Sub

Private Sub cbeRelease_Click(ByVal m_item As Object, _
                             ByRef handled As Boolean, _
                             ByRef CancelDefault As Boolean)
    mRelease.ReleaseService
        
End Sub

Private Sub cbeReleaseComp_Click(ByVal m_item As Object, _
                                 ByRef handled As Boolean, _
                                 ByRef CancelDefault As Boolean)
    Dim cbb     As CommandBarButton
    Dim sComp As String
    
    Set cbb = m_item
    sComp = Split(cbb.Caption, " ")(1)
    CommonPending.ReleaseComp sComp
    CommCompsPendingRelease.Remove 1
    mCompManMenuVBE.Setup
    
End Sub

Friend Property Set MenuItemRelease(ByRef m_cbe As CommandBarEvents)
    Set cbeRelease = m_cbe
End Property

Friend Property Set MenuItemReleaseComp(ByRef m_cbe As CommandBarEvents)
    Set cbeReleaseComp = m_cbe
End Property

