Attribute VB_Name = "mMenuVBEItems"
Option Explicit
' ----------------------------------------------------------------------------
' Standard module mMenuVBECommandBarEvents: Add/remove buttons to a VBE menu.
' ----------------------------------------------------------------------------
Private Const BTTN_RELEASE      As String = "Release to public ..."
Private Const BTTN_RELEASE_COMP As String = "Release " ' generic caption

Private cbbVBEVBA    As clsVBEMenuCbbVBEVB6
Private cbbOffice    As clsVBEMenuCbbOffice

Private Sub AddButton(ByVal a_cmb_popup As CommandBarPopup, _
                      ByVal a_caption As String, _
                      ByVal a_face_id As Long, _
                      ByRef a_cmb_button As CommandBarButton, _
             Optional ByVal a_begin_group As Boolean = False)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "AddButton"
    
    On Error GoTo eh
    Set a_cmb_button = a_cmb_popup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With a_cmb_button
        .BeginGroup = a_begin_group
        .Caption = a_caption
        .FaceId = a_face_id
        .Style = msoButtonIconAndCaption
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AddMenuItemsWithVBECommandBarEvents(ByVal a_cmb_popup As CommandBarPopup)
' ----------------------------------------------------------------------------
' Add CommandButtons to the CommandBarPopup (a_cmb_popup).
' ----------------------------------------------------------------------------
    Const PROC = "AddMenuItemsWithVBECommandBarEvents"
    
    Dim cbb As CommandBarButton
              
    AddButton a_cmb_popup:=cbb _
            , a_caption:=BTTN_RELEASE _
            , a_face_id:=462 _
            , a_cmb_button:=cbb _
            , a_begin_group:=True
    Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
    Set cbbVBEVBA.MenuItemRelease = Application.VBE.Events.CommandBarEvents(cbb)
    cllCommandBarEvents.Add cbbVBEVBA 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbVBEVBA = Nothing
    
    AddButton a_cmb_popup:=cbb _
            , a_caption:=BTTN_RELEASE_COMP & "XXXXX" _
            , a_face_id:=464 _
            , a_cmb_button:=cbb
    Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
    Set cbbVBEVBA.MenuItemReleaseComp = Application.VBE.Events.CommandBarEvents(cbb)
    cllCommandBarEvents.Add cbbVBEVBA 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbVBEVBA = Nothing

End Sub

Public Sub AddMenuItemsWithOfficeEvents(ByVal a_cmb_popup As CommandBarPopup)
' ----------------------------------------------------------------------------
' Add CommandButtons to the CommandBarPopup (a_cmb_popup).
' ----------------------------------------------------------------------------
    Const PROC = "AddMenuItemsWithCommandBarEvents"
    
    Dim cbb As CommandBarButton
              
    AddButton a_cmb_popup:=cbb _
            , a_caption:=BTTN_RELEASE _
            , a_face_id:=462 _
            , a_cmb_button:=cbb _
            , a_begin_group:=True
    Set cbbOffice = New clsVBEMenuCbbOffice
    Set cbbOffice.CmdBarBtn = cbb
    cllCommandBarEvents.Add cbbOffice 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbOffice = Nothing
    
    AddButton a_cmb_popup:=cbb _
            , a_caption:=BTTN_RELEASE_COMP & "XXXXX" _
            , a_face_id:=464 _
            , a_cmb_button:=cbb
    Set cbbOffice = New clsVBEMenuCbbOffice
    Set cbbOffice.CmdBarBtn = cbb
    cllCommandBarEvents.Add cbbOffice 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbOffice = Nothing

xt: Exit Sub
    
eh:

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMenuVBECommandBarEvents." & sProc
End Function

Public Sub ReleaseComp_Click()
    MsgBox "ReleaseComp clicked!"
End Sub

Public Sub Release_Click()
    MsgBox "Release clicked!"
End Sub

