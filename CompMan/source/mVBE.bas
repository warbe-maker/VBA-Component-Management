Attribute VB_Name = "mVBE"
Option Explicit

Private Const MENU_NAME As String = "CompMan"
Private Const ITEM_NAME As String = "Release pending modification "

Private cbp  As CommandBarPopup
Private cbc1 As CommandBarControl
Private cbc2 As CommandBarControl

Public cllCommandBarEvents As New Collection 'collection to store menu item click event handlers

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mVBE." & sProc
End Function

Public Sub MenuAdd()
    Const PROC = "MenuAdd"
    
    Dim v As Variant
    Dim cbe As clsVBEBarEvents
    Dim cbb As CommandBarButton
    
    MenuRemove
    Set cbp = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbp.Caption = MENU_NAME
    cbp.Tag = "CustomMenu"
    cbp.Visible = True
    
    ' Add items to the custom menu
    Set cbb = cbp.Controls.Add(msoControlButton, , , 1)
    With cbb
        .Caption = "Custom Item 1"
        .OnAction = "CustomItem1_Click"
        .FaceId = 29
        Debug.Print ErrSrc(PROC) & ": " & "OnAction=" & .OnAction
    End With
    
    Set cbe = New clsVBEBarEvents 'Create a new instance of our button event-handling class
    Set cbe.oCBControlEvents = Application.VBE.Events.CommandBarEvents(cbb) 'Tell the class to hook into the events for this button
    cllCommandBarEvents.Add cbe 'And add the event handler to our collection of handlers
    Set cbe = Nothing
    
    Set cbb = cbp.Controls.Add(msoControlButton, , , 1)
    With cbb
        .Caption = "Custom Item 2"
        .OnAction = "CustomItem2_Click"
        Debug.Print ErrSrc(PROC) & ": " & "OnAction=" & .OnAction
    End With
    
    Set cbe = New clsVBEBarEvents 'Create a new instance of our button event-handling class
    Set cbe.oCBControlEvents = Application.VBE.Events.CommandBarEvents(cbb) 'Tell the class to hook into the events for this button
    cllCommandBarEvents.Add cbe 'And add the event handler to our collection of handlers
    Set cbe = Nothing
    
End Sub

Public Sub CustomItem1_Click()
    MsgBox "Custom Item 1 clicked!"
End Sub

Public Sub CustomItem2_Click()
    MsgBox "Custom Item 2 clicked!"
End Sub

Sub MenuRemove()
    On Error Resume Next
    Application.VBE.CommandBars(1).Controls(MENU_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub MenuItemAdd(ByVal m_caption As String)

End Sub

