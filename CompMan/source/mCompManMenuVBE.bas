Attribute VB_Name = "mCompManMenuVBE"
Option Explicit
' ----------------------------------------------------------------------------
' Standard module mMenuVBE: Prepare CompMan menu in VBE.
' ----------------------------------------------------------------------------
Public cllCommandBarEvents      As New Collection 'collection to store menu item click event handlers
Public Const MENU_NAME          As String = "CompMan"

Private Const BTTN_RELEASE      As String = "Release pending modifications by this Workbook ..."
Private Const BTTN_RELEASE_COMP As String = "Release <comp> modified by Workbook <wbk>" ' generic caption

Private cbp         As CommandBarPopup
Private cbbVBEVBA   As clsVBEMenuCbbVBEVB6
Private cbbOffice   As clsVBEMenuCbbOffice


Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMenuVBE." & sProc
End Function

Public Sub MenuCreate()
    
    Set cbp = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With cbp
        .Caption = mCompManMenuVBE.MENU_NAME
        .Tag = "CustomMenu"
        .Visible = True
    End With
    
End Sub

Public Sub Setup()
' ----------------------------------------------------------------------------
' Provides a CompMan menu in the VBE.
' ----------------------------------------------------------------------------
    Const PROC = "Setup"
    
    MenuRemove
    If CommonPending.ReadyForRelease.Count = 0 Then Exit Sub
    MenuCreate
    
    On Error GoTo om
    MenuItemsAddWithOfficeEvents
    GoTo xt
    
om: '~~ The "CommandBarEvents" method as alternative in case the "OfficeEvents" method failed
    MenuRemove
    MenuCreate
    MenuItemsAddWithVBECommandBarEvents

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub MenuRemove()
    On Error Resume Next
    Application.VBE.CommandBars(1).Controls(MENU_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub MenuAddButton(ByVal a_caption As String, _
                      ByVal a_face_id As Long, _
                      ByRef a_cmb_button As CommandBarButton, _
             Optional ByVal a_begin_group As Boolean = False)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "MenuAddButton"
    
    On Error GoTo eh
    Set a_cmb_button = cbp.Controls.Add(Type:=msoControlButton, Temporary:=True)
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

Public Sub MenuItemsAddWithVBECommandBarEvents()
' ----------------------------------------------------------------------------
' Adds CommandButtons to the CompMan menu in the VBE (a_cmb_popup).
' ----------------------------------------------------------------------------
    Dim cbb     As CommandBarButton
    Dim lFaceId As Long
    Dim v       As Variant
    
    '~~ Invoke the pending release dialog (component by component with display modification option)
    MenuAddButton a_caption:=BTTN_RELEASE _
                , a_face_id:=806 _
                , a_cmb_button:=cbb _
                , a_begin_group:=True
    Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
    Set cbbVBEVBA.MenuItemRelease = Application.VBE.Events.CommandBarEvents(cbb)
    cllCommandBarEvents.Add cbbVBEVBA 'And the event handler to a collection of handlers
    Set cbb = Nothing
        
    '~~ Direkt release without dialog.
    '~~ Note: This works only when the to-be-released component is a component in the current serviced Workbook
    For Each v In CommonPending.ReadyForRelease
        Select Case FSo.GetExtensionName(CommonPending.LastModExpFile(v))
            Case "cls":  lFaceId = 229
            Case "frm":  lFaceId = 230
            Case "bas":  lFaceId = 231
        End Select
        MenuAddButton a_caption:=Replace(Replace(BTTN_RELEASE_COMP, "<comp>", v), "<wbk>", CommonPending.LastModInWrkbkName(v)) _
                    , a_face_id:=lFaceId _
                    , a_cmb_button:=cbb
        Set cbbVBEVBA = New clsVBEMenuCbbVBEVB6
        Set cbbVBEVBA.MenuItemReleaseComp = Application.VBE.Events.CommandBarEvents(cbb)
        cllCommandBarEvents.Add cbbVBEVBA 'And add the event handler to our collection of handlers
        Set cbb = Nothing
    Next v
    
End Sub

Public Sub MenuItemsAddWithOfficeEvents()
' ----------------------------------------------------------------------------
' Add CommandButtons to the CommandBarPopup (a_cmb_popup).
' ----------------------------------------------------------------------------
    Dim cbb     As CommandBarButton
    Dim lFaceId As Long
    Dim v       As Variant
              
    '~~ Invoke the pending release dialog (component by component with display modification option)
    MenuAddButton a_caption:=BTTN_RELEASE _
                , a_face_id:=806 _
                , a_cmb_button:=cbb _
                , a_begin_group:=True
    Set cbbOffice = New clsVBEMenuCbbOffice
    Set cbbOffice.CmdRelease = cbb
    cllCommandBarEvents.Add cbbOffice 'And add the event handler to our collection of handlers
    Set cbb = Nothing
    Set cbbOffice = Nothing
    
    '~~ Direkt release without dialog
    For Each v In CommonPending.ReadyForRelease
        Select Case FSo.GetExtensionName(CommonPending.LastModExpFile(v))
            Case "cls":  lFaceId = 229
            Case "frm":  lFaceId = 230
            Case "bas":  lFaceId = 231
        End Select
        MenuAddButton a_caption:=Replace(Replace(BTTN_RELEASE_COMP, "<comp>", v), "<wbk>", CommonPending.LastModInWrkbkName(v)) _
                    , a_face_id:=lFaceId _
                    , a_cmb_button:=cbb
        Set cbbOffice = New clsVBEMenuCbbOffice
        Set cbbOffice.CmdReleaseComp = cbb
        cllCommandBarEvents.Add cbbOffice 'And add the event handler to our collection of handlers
        Set cbb = Nothing
        Set cbbOffice = Nothing
    Next v

xt: Exit Sub
    
eh:

End Sub

Public Sub ReleaseComp_Click()
    MsgBox "ReleaseComp clicked!"
End Sub

Public Sub Release_Click()
    MsgBox "Release clicked!"
End Sub


