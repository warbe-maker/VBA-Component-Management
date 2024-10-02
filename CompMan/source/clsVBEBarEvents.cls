VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBEBarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents oCBControlEvents As CommandBarEvents
Attribute oCBControlEvents.VB_VarHelpID = -1

Private Sub oCBControlEvents_Click(ByVal cbCommandBarControl As Object, _
                                   ByRef handled As Boolean, _
                                   ByRef CancelDefault As Boolean)

    On Error Resume Next
    'Run the routine given by the commandbar control's OnAction property
    Application.Run cbCommandBarControl.OnAction

    handled = True
    CancelDefault = True

End Sub

