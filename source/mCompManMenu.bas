Attribute VB_Name = "mCompManMenu"
Option Explicit
' ------------------------------------------------------------------------------
' Standard  Module mMenu: Provides the "CompMan" menu with all its appropriate
' ======================= items.
'
' ------------------------------------------------------------------------------
Private Const MENU_NAME                         As String = "CompMan"
Private Const MENU_ITEM_RELEASE_SERVICE         As String = "Release Common Component changes"
Private Const MENU_ITEM_HELP_RELEASE_SERVICE    As String = "Help Release"
Private Const MENU_ITEM_HELP_SERVICED           As String = "Help Serviced"
Private Const MENU_ITEM_HELP_CONFIGURE          As String = "Help Configure"
Private Const MENU_ITEM_ABOUT                   As String = "About"
Private Const README_SERVICED                   As String = "#enabling-the-services-serviced-or-not-serviced"
Private Const README_RELEASE                    As String = "#the-release-service"
Private Const MENU_ITEM_TAG                     As String = "COMPMAN_MENU_ITEM"

Private objPopUp                                As CommandBarPopup
Private dctPendingRelease                       As New Dictionary
Private CmdBarItem                              As CommandBarControl
Private cllEventHandlers                        As New Collection

Private Sub DeleteMenuItems()
' ------------------------------------------------------------------------------
' Delete all controls that have a tag of MENU_ITEM_TAG.
' ------------------------------------------------------------------------------
    Dim Ctrl As Office.CommandBarControl
    Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=MENU_ITEM_TAG)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=MENU_ITEM_TAG)
    Loop
End Sub

Public Sub Procedure_One()
    MsgBox "Procedure One"
End Sub

Public Sub Procedure_Two()
    MsgBox "Procedure Two"
End Sub

Public Property Get ThisMenuBar() As CommandBar
    Set ThisMenuBar = Application.CommandBars("Worksheet Menu Bar")
'    Set ThisMenuBar = Application.VBE.CommandBars("Menüleiste") ' unfortunately not possible
End Property

Private Sub Explore()
    Dim cmb     As CommandBar
    Dim cbp1    As Variant
    Dim cbp2    As Variant
    
    For Each cmb In Application.VBE.CommandBars
        For Each cbp1 In cmb.Controls
            Debug.Print "   " & TypeName(cbp1) & ": '" & cbp1.Caption & "' Id: " & cbp1.id
            If TypeOf cbp1 Is CommandBarPopup Then
                For Each cbp2 In cbp1.Controls
                    Debug.Print "      " & TypeName(cbp2) & ": '" & cbp2.Caption & "' Id: " & cbp2.id
                Next cbp2
            End If
        Next cbp1
    Next cmb
xt:
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMenu." & sProc
End Function

Private Function CompMan() As CommandBarPopup
' ------------------------------------------------------------------------------
' Returns the CompMan CommandBarPopup.
' ------------------------------------------------------------------------------
    Dim cmb     As CommandBar
    Dim cbp     As CommandBarPopup
    Dim cbr     As CommandBar
    
    Set cbr = ThisMenuBar
    For Each cbp In cbr.Controls
        If TypeOf cbp Is CommandBarPopup Then
            If cbp.Caption Like "*" & MENU_NAME & "*" _
            Or cbp.Tag = MENU_ITEM_TAG Then
                Set CompMan = cbp
                GoTo xt
            End If
        End If
    Next cbp
xt:
End Function

Private Function Exists(ByRef e_bar As CommandBar, _
                        ByRef e_cbp As CommandBarPopup) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the CompMan menu exists.
' ------------------------------------------------------------------------------
    
    Set e_bar = ThisMenuBar
    Set e_cbp = CompMan
    Exists = Not e_bar Is Nothing And Not e_cbp Is Nothing
    If Not Exists And Not e_cbp Is Nothing Then
          Exists = e_cbp.Tag = MENU_ITEM_TAG
    End If
    
End Function

Public Sub Remove()
' ------------------------------------------------------------------------------
' Removes the CompMan menu when the last open CompMan instance is closed.
' ------------------------------------------------------------------------------
    Dim cbr As CommandBar
    Dim cbp As CommandBarPopup
    
    Select Case True
        Case mAddin.IsOpen And mMe.IsDevInstnc
            '~~ When the dev instance is closed and there's still the Addin instance open
            '~~ the CompMan menu is not removed
        Case mMe.IsAddinInstnc And mWbk.IsOpen(mCompManClient.COMPMAN_DEVLP)
            '~~ When the Add-in instance is closed and the dev instance is still open
            '~~ this means that the Add-in is about to be renewed. In case the renewal
            '~~ fails the CompMan menu still needs to remain avialable.
        Case Else
            While Exists(cbr, cbp)
                '~~ The CompMan menu exists in the Add-Ins menu as sub-menu
                cbr.Controls(cbp.Index).Delete
            Wend
    End Select

End Sub

Public Sub HelpConfigure()
    Debug.Print "HelpConfigure performed"
    mBasic.README r_url:=GITHUB_REPO_URL _
                , r_bookmark:=README_CONFIG_CHANGES
End Sub

Public Sub HelpServiced()
    Debug.Print "HelpServiced performed"
    mBasic.README r_url:=GITHUB_REPO_URL _
                , r_bookmark:=README_SERVICED
End Sub

Public Sub RemoveItem(ByVal r_item As String)
' ------------------------------------------------------------------------------
' Removes the item (r_item)
' ------------------------------------------------------------------------------
    Dim cbr As CommandBar
    Dim cbp As CommandBarPopup
    
    If Exists(cbr, cbp) Then
        On Error Resume Next
        cbp.Controls(r_item).Delete
    End If
   
End Sub

Public Sub Setup()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Setup"
    
    On Error GoTo eh
    Dim cbr             As CommandBar
    Dim cbp             As CommandBarPopup
    Dim cbpComoManMenu  As CommandBarPopup
    
    If CommComps Is Nothing Then Set CommComps = New clsCommComps
    
    While Exists(cbr, cbp)
        '~~ The CompMan menu exists in the Add-Ins menu as sub-menu
        cbr.Controls(cbp.Index).Delete
    Wend
    If CommComps.PendingReleases.Count <> 0 Then
        '~~ When there are no releases pending the CompMan menu may not exist
        Set cbpComoManMenu = ThisMenuBar.Controls.Add(Type:=msoControlPopup)
        With cbpComoManMenu
            .Caption = MENU_NAME
            .Tag = MENU_ITEM_TAG
        End With
        mCompManMenu.SetupItem MENU_ITEM_RELEASE_SERVICE & " (" & CommComps.PendingReleases.Count & " pending) ...", "ReleaseService"
        mCompManMenu.SetupItem MENU_ITEM_HELP_RELEASE_SERVICE, "HelpRelease"
    End If
    '~~ Other, pending releases independant items
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SetupItem(ByVal s_caption As String, _
                      ByVal s_action As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "SetupItem"
    
    On Error GoTo eh
    Dim cbp As CommandBarPopup
    Dim cbc As CommandBarControl
    
    Set cbp = CompMan
    If Not cbp Is Nothing Then
'        Set MenuEvent = New CVBECommandHandler
        Set cbc = cbp.Controls.Add(msoControlButton)
        With cbc
            .Caption = s_caption
            .OnAction = "'" & ThisWorkbook.Name & "'!" & s_action
            .Style = msoButtonCaption
            .Enabled = True
            .Tag = MENU_ITEM_TAG
        End With
    
'        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(cbc)
'        cllEventHandlers.Add MenuEvent
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

