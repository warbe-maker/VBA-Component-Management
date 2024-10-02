Attribute VB_Name = "mVBEMenu"
Option Explicit

Private Const BUTTEN_CAPTION1 = "Zeilennummern &setzen"
Private Const BUTTEN_CAPTION2 = "&Zeilennummern &löschen"

Private objcmbButton As clscmbButtonEvent

'=================================================================
' CommandButtons zum Menü Extras im VB-Editor hinzufügen
'=================================================================

Public Sub prcCreateButton()
    Dim cmbButton As CommandBarButton
    
    On Error GoTo err_exit
    Call prcDeleteButton
    
    Set objcmbButton = New clscmbButtonEvent
    Set cmbButton = Application.VBE.CommandBars(1).Controls("E&xtras"). _
        Controls.Add(Type:=msoControlButton, Temporary:=True)
    With cmbButton
        .BeginGroup = True
        .Caption = BUTTEN_CAPTION1
        .FaceId = 462
        .Style = msoButtonIconAndCaption
    End With
    
    Set objcmbButton.prpButtonEvents1 = _
        Application.VBE.Events.CommandBarEvents(cmbButton)
    
    Set cmbButton = Application.VBE.CommandBars(1).Controls("E&xtras"). _
        Controls.Add(Type:=msoControlButton, Temporary:=True)
    
    With cmbButton
        .Caption = BUTTEN_CAPTION2
        .FaceId = 464
        .Style = msoButtonIconAndCaption
    End With
    Set objcmbButton.prpButtonEvents2 = _
        Application.VBE.Events.CommandBarEvents(cmbButton)
    Set cmbButton = Nothing
    Exit Sub
    
err_exit:
    MsgBox "Fehler " & CStr(Err.Number) & vbLf & vbLf & _
        Err.Description, vbCritical, "Fehlermeldung"
End Sub


'=================================================================
' CommandButtons zum Menü Extras im VB-Editor entfernen
'=================================================================

Public Sub prcDeleteButton()
    Dim cmbButton As CommandBarControl
    On Error GoTo err_exit
    For Each cmbButton In Application.VBE.CommandBars(1). _
            Controls("E&xtras").Controls
        If cmbButton.Caption = BUTTEN_CAPTION1 Or _
            cmbButton.Caption = BUTTEN_CAPTION2 Then cmbButton.Delete
    Next
    Set objcmbButton = Nothing
    Exit Sub
    
err_exit:
    MsgBox "Fehler " & CStr(Err.Number) & vbLf & vbLf & _
        Err.Description, vbCritical, "Fehlermeldung"
End Sub

