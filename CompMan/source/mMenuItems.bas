Attribute VB_Name = "mMenuItems"
Option Explicit
' ----------------------------------------------------------------------------
' Standard module mMenuVBECommandBarEvents: Add/remove buttons to a VBE menu.
' ----------------------------------------------------------------------------
Private Const BTTN_RELEASE      As String = "Release to public ..."
Private Const BTTN_RELEASE_COMP As String = "Release " ' generic caption

Private cbbVBEVBA    As clsVBEMenuCbbVBEVB6
Private cbbOffice    As clsVBEMenuCbbOffice

Public Sub AddWithCommandBarEvents(ByVal a_cbp As CommandBarPopup)
' ----------------------------------------------------------------------------
' Add CommandButtons to the CommandBarPopup (a_cbp).
' ----------------------------------------------------------------------------
    Const PROC = "AddMenuItemsWithCommandBarEvents"
    
    On Error GoTo eh
    Dim cbb As CommandBarButton
              
    '~~ Button 1
    Set cbb = a_cbp.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With cbb
        .BeginGroup = True
        .Caption = BTTN_RELEASE
        .FaceId = 462
        .Style = msoButtonIconAndCaption
    End With
    Set cbbOffice = New clsVBEMenuCbbOffice
    Set cbbOffice.CmdBarBtn = cbb
    cllCommandBarEvents.Add cbbOffice 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbOffice = Nothing
    
    '~~ Button 2
    Set cbb = a_cbp.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With cbb
        .Caption = BTTN_RELEASE_COMP & "XXXXX"
        .FaceId = 464
        .Style = msoButtonIconAndCaption
    End With
    Set cbbOffice = New clsVBEMenuCbbOffice
    Set cbbOffice.CmdBarBtn = cbb
    cllCommandBarEvents.Add cbbOffice 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbOffice = Nothing

xt: Exit Sub
    
eh:

End Sub

Public Sub AddWithClickEvents(ByVal a_cbp As CommandBarPopup)
' ----------------------------------------------------------------------------
' Add CommandButtons to the CommandBarPopup (a_cbp).
' ----------------------------------------------------------------------------
    Const PROC = "AddMenuItemsWithCommandBarEvents"
    
    On Error GoTo eh
    Dim cbb As CommandBarButton
              
    '~~ Button 1
    Set cbb = a_cbp.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With cbb
        .BeginGroup = True
        .Caption = BTTN_RELEASE
        .FaceId = 462
        .Style = msoButtonIconAndCaption
    End With
    Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
    Set cbbVBEVBA.MenuItemRelease = Application.VBE.Events.CommandBarEvents(cbb)
    cllCommandBarEvents.Add cbbVBEVBA 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbVBEVBA = Nothing
    
    '~~ Button 2
    Set cbb = a_cbp.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With cbb
        .Caption = BTTN_RELEASE_COMP & "XXXXX"
        .FaceId = 464
        .Style = msoButtonIconAndCaption
    End With
    Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
    Set cbbVBEVBA.MenuItemReleaseComp = Application.VBE.Events.CommandBarEvents(cbb)
    cllCommandBarEvents.Add cbbVBEVBA 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbVBEVBA = Nothing

xt: Exit Sub
    
eh:

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMenuVBECommandBarEvents." & sProc
End Function

Public Sub Release_Click()
    MsgBox "Release clicked!"
End Sub

Public Sub ReleaseComp_Click()
    MsgBox "ReleaseComp clicked!"
End Sub

