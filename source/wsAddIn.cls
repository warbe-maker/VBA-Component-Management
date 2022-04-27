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

Public Property Let CompManAddInPausedStatus(ByVal s As String):        Me.Range("rngCompManPauseStatus").Value = s:            End Property

Private Property Get MonitorAnchor() As Range:                          Set MonitorAnchor = Me.Range("celMonitorAnchor"):       End Property

Public Property Get CurrentStatus() As String:                          CurrentStatus = Me.Range("rngCurrentStatus").Value:     End Property

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

Public Property Let RenewInfo(ByVal s As String):                       Me.Range("rngRenewInfo").Value = s:                 End Property

Public Sub LogRenewStep( _
               Optional ByVal rn_action As String = vbNullString, _
               Optional ByVal rn_result As String = vbNullString, _
               Optional ByVal rn_last_step As Boolean = False, _
               Optional ByVal rn_renew_failed As Boolean = False)
' ---------------------------------------------------------------
'
' ---------------------------------------------------------------
    Dim Title   As String:      Title = "Establish or renew CompMan as Addin"
    Dim Header  As TypeMsgText
    Dim Step    As TypeMsgText
    Dim Footer  As TypeMsgText
    Dim l       As Long:        l = mMe.RenewStep
    
    RenewStep = l
    If rn_result <> vbNullString Then RenewResult = rn_result
    If rn_action <> vbNullString Then RenewAction = rn_action
    If RenewResult <> vbNullString Then ProgressMsg ' RenewStep, RenewResult, RenewAction
    
    If RenewStep > 0 And RenewResult <> vbNullString Then
        Step.Text = sProgressMsg
        Step.MonoSpaced = True
        If Not rn_last_step Then
            mMsg.Monitor mon_title:=Title _
                   , mon_text:=Step _
                   , mon_steps_displayed:=12 _
                   , mon_width_min:=70 _
                   , mon_width_max:=80 _
                   , mon_pos:=MonitorAnchor
        Else
            mMsg.Monitor Title, Step
            If mMe.RenewTerminatedByUser Then
                Footer.Text = "Renew process terminated by user! Close window."
                mMsg.MonitorFooter Title, Footer
            Else
                If Not rn_renew_failed Then
                    With Footer
                        .Text = "Renew process  " & mBasic.Spaced("successfully") & "  finished! Close window."
                        .FontBold = True
                        .FontColor = rgbBlue
                    End With
                    mMsg.MonitorFooter Title, Footer
                Else
                    With Footer
                        .Text = "Renew process  " & mBasic.Spaced("failed") & "  Close window and " & vbLf & _
                                "the only way to solve this problem is:" & vbLf & _
                                "1. Close all application and restart the computer" & vbLf & _
                                "2. Remove the CompMan.xlma from its folder" & vbLf & _
                                "3. Reopen CompMan.xlsb and repeat this Renew (then a setup in fact)."
                        .FontColor = rgbRed
                    End With
                    mMsg.MonitorFooter Title, Footer
                End If
            End If
        End If
        sProgressMsg = vbNullString
    End If
    RenewResult = vbNullString
    
End Sub

Private Sub ProgressMsg()
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

Private Sub Worksheet_Activate()
    With Me
        Application.EnableEvents = False
        .Range("rngDisplayed").Select
        ActiveWindow.Zoom = True
        .Range("A1").Select
        .Range("celSelect").Select
        Application.EnableEvents = True
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
