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
Private RenewStep       As Long
Private RenewAction     As String
Private RenewResult     As String
Private sProgressMsg    As String

Private Sub ProgressMsg(ByVal prgrs_step As Long, ByVal prgrs_result As String, ByVal prgrs_action As String)
    If sProgressMsg = vbNullString Then
        sProgressMsg = mBasic.Align(RenewStep, 2, AlignRight) & " " & RenewResult & " " & Split(RenewAction, vbLf)(0)
        If UBound(Split(RenewAction, vbLf)) > 0 Then
            sProgressMsg = sProgressMsg & vbLf & "          " & Split(RenewAction, vbLf)(1)
        End If
    Else
        sProgressMsg = sProgressMsg & vbLf & mBasic.Align(RenewStep, 2, AlignRight) & " " & RenewResult & " " & Split(RenewAction, vbLf)(0)
        If UBound(Split(RenewAction, vbLf)) > 0 Then
            sProgressMsg = sProgressMsg & vbLf & "   " & Split(RenewAction, vbLf)(1)
        End If
    End If
End Sub

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

Public Sub LogRenewStep( _
               Optional ByVal rn_action As String = vbNullString, _
               Optional ByVal rn_result As String = vbNullString, _
               Optional ByVal rn_last_step As Boolean = False)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    Dim msg As String
    Dim l   As Long: l = mMe.RenewStep
    
    RenewStep = l
    If rn_result <> vbNullString Then RenewResult = rn_result
    If rn_action <> vbNullString Then RenewAction = rn_action
    If RenewResult <> vbNullString Then ProgressMsg RenewStep, RenewResult, RenewAction
    
    If RenewStep > 1 And RenewResult <> vbNullString Then
        msg = sProgressMsg
        If Not rn_last_step Then
            mMsg.Monitor mntr_title:="Log of the steps to Setup/Renew the 'CompMan-Addin'" _
                       , mntr_msg:=msg _
                       , mntr_msg_monospaced:=True _
                       , mntr_width_min:=600
        Else
            mMsg.Monitor mntr_title:="Log of the steps to Setup/Renew the 'CompMan-Addin'" _
                       , mntr_msg:=msg _
                       , mntr_msg_monospaced:=True _
                       , mntr_footer:="Renew process finished! Close window." _
                       , mntr_width_min:=600
        End If
        sProgressMsg = vbNullString
    End If
    RenewResult = vbNullString
    
End Sub

Private Sub Worksheet_Activate()
End Sub