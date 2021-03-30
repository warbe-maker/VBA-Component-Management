VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private sCurrentStatus  As String

Public Property Let CompManAddInPausedStatus(ByVal s As String):        Me.Range("rngCompManPauseStatus").Value = s:        End Property

Public Property Get CurrentStatus() As String:                          CurrentStatus = Me.Range("rngCurrentStatus").Value: End Property

Public Property Let CurrentStatus(ByVal s As String)
    With Me
        If s = vbNullString Then
            sCurrentStatus = vbNullString
            .Range("rngCurrentStatus").ClearContents
        Else
            sCurrentStatus = sCurrentStatus & vbLf & s & vbLf
            .Range("rngCurrentStatus").Value = sCurrentStatus
        End If
    End With
End Property

Public Property Let RenewStepAction(Optional ByVal rs_step As Long, ByVal s As String)
    Intersect(Me.Range("rowRenewStep" & rs_step), Me.Range("colRenewStepExecuted")).Value = s
End Property

Public Property Let RenewStepLogSubTitle(ByVal s As String):            Me.Range("rngRenewLogSubTitle").Value = s:          End Property

Public Property Let RenewStepLogTitle(ByVal s As String):               Me.Range("rngRenewLogTitle").Value = s:             End Property

Public Property Let RenewStepNo(Optional rn_step As Long, ByVal rn_step_value As String)
        Intersect(Me.Range("rowRenewStep" & rn_step), Me.Range("colRenewStep")).Value = rn_step_value
End Property

Public Property Let RenewStepResult(Optional ByVal rs_step As Long, ByVal s As String)
    Intersect(Me.Range("rowRenewStep" & rs_step), Me.Range("colRenewStepResult")).Value = s
End Property

Public Sub LogRenewStep( _
               Optional ByVal rn_action As String = vbNullString, _
               Optional ByVal rn_result As String = vbNullString)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    Application.EnableEvents = False
    Dim l As Long: l = mMe.RenewStep
    
    With wsAddIn
        .RenewStepNo(l) = l
        If rn_result <> vbNullString Then
           .RenewStepResult(l) = rn_result
        End If
        If rn_action <> vbNullString Then
            .RenewStepAction(l) = rn_action
        End If
    End With
    Application.EnableEvents = False
    DoEvents
    Application.Wait Now() + 0.00000001
    
End Sub

Public Sub RenewStepLogClear()
    Me.Range("rngRenewLog").ClearContents
End Sub

Private Sub Worksheet_Activate()
End Sub
