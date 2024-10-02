Attribute VB_Name = "mCompManMenuVBEItems"
Option Explicit
' ----------------------------------------------------------------------------
' Standard module mMenuVBECommandBarEvents: Add/remove buttons to a VBE menu.
' ----------------------------------------------------------------------------
Private Const BTTN_RELEASE      As String = "Release pending modifications by this Workbook ..."
Private Const BTTN_RELEASE_COMP As String = "Release <comp> modified by this Workbook" ' generic caption

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

Public Sub AddMenuItemsWithVBECommandBarEvents(ByRef a_cmb_popup As CommandBarPopup)
' ----------------------------------------------------------------------------
' Adds CommandButtons to the CompMan menu in the VBE (a_cmb_popup).
' ----------------------------------------------------------------------------
    Const PROC = "AddMenuItemsWithVBECommandBarEvents"
    
    Dim cbb     As CommandBarButton
    Dim i       As Long
    Dim lFaceId As Long
    Dim v       As Variant
    
    '~~ Invoke pending release dialog
    AddButton a_cmb_popup:=a_cmb_popup _
            , a_caption:=BTTN_RELEASE _
            , a_face_id:=806 _
            , a_cmb_button:=cbb _
            , a_begin_group:=True
    Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
    Set cbbVBEVBA.MenuItemRelease = Application.VBE.Events.CommandBarEvents(cbb)
    cllCommandBarEvents.Add cbbVBEVBA 'And the event handler to a collection of handlers
    Set cbb = Nothing
        
    '~~ Direkt release without dialog
    For Each v In CommonPending.Components
        Select Case Serviced.Wrkbk.VBProject.VBComponents(v).Type
            Case vbext_ct_ClassModule:  lFaceId = 229
            Case vbext_ct_MSForm:       lFaceId = 230
            Case vbext_ct_StdModule:    lFaceId = 231
        End Select
        AddButton a_cmb_popup:=a_cmb_popup _
                , a_caption:=Replace(BTTN_RELEASE_COMP, "<comp>", v) _
                , a_face_id:=lFaceId _
                , a_cmb_button:=cbb
        Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
        Set cbbVBEVBA.MenuItemReleaseComp = Application.VBE.Events.CommandBarEvents(cbb)
        cllCommandBarEvents.Add cbbVBEVBA 'And add the event handler to our collection of handlers
        Set cbb = Nothing
    Next v
    
End Sub

Public Sub AddMenuItemsWithOfficeEvents(ByVal a_cmb_popup As CommandBarPopup)
' ----------------------------------------------------------------------------
' Add CommandButtons to the CommandBarPopup (a_cmb_popup).
' ----------------------------------------------------------------------------
    Const PROC = "AddMenuItemsWithCommandBarEvents"
    
    Dim cbb As CommandBarButton
              
    AddButton a_cmb_popup:=cbb _
            , a_caption:=BTTN_RELEASE _
            , a_face_id:=806 _
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

