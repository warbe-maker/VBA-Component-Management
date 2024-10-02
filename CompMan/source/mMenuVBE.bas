Attribute VB_Name = "mMenuVBE"
Option Explicit
' ----------------------------------------------------------------------------
' Standard module mMenuVBE: Prepare CompMan menu in VBE.
' ----------------------------------------------------------------------------
Public cllCommandBarEvents As New Collection 'collection to store menu item click event handlers

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMenuVBE." & sProc
End Function

Public Sub MenuCreate(ByRef m_cbp As CommandBarPopup)
    
    Set m_cbp = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With m_cbp
        .Caption = mCompManMenu.MENU_NAME
        .Tag = "CustomMenu"
        .Visible = True
    End With
    
End Sub

Public Sub Setup()
' ----------------------------------------------------------------------------
' Provides a CompMan menu in the VBE.
' ----------------------------------------------------------------------------
    Const PROC = "Setup"
    
    Dim v       As Variant
    Dim cbp     As CommandBarPopup
    
    MenuRemove
    MenuCreate cbp
    
    On Error GoTo om
    mMenuVBEItems.AddMenuItemsWithVBECommandBarEvents cbp
    GoTo xt
    
om: '~~ The "Office" method (alternative to the above when it raised an error
    MenuRemove
    MenuCreate cbp
    mMenuVBEItems.AddMenuItemsWithOfficeEvents cbp

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SetupTest()
    mMenuVBE.Setup
End Sub

Public Sub MenuRemove()
    On Error Resume Next
    Application.VBE.CommandBars(1).Controls(MENU_NAME).Delete
    On Error GoTo -1
End Sub

Private Sub MenuItemAdd(ByVal m_caption As String)

End Sub

