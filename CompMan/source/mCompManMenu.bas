Attribute VB_Name = "mCompManMenu"
'Option Explicit
'' ------------------------------------------------------------------------------
'' Standard  Module mMenu: Provides the "CompMan" menu with all its appropriate
'' ======================= items.
''
'' ------------------------------------------------------------------------------
'Public Const MENU_NAME                         As String = "CompMan"
'
'Private Const MENU_ITEM_RELEASE_SERVICE         As String = "Release Common Component changes"
'Private Const MENU_ITEM_HELP_RELEASE_SERVICE    As String = "Help Release"
'Private Const README_SERVICED                   As String = "#enabling-the-services-serviced-or-not-serviced"
'Private Const MENU_ITEM_TAG                     As String = "COMPMAN_MENU_ITEM"
'
'Private cbr                                     As CommandBar
'Private cbp                                     As CommandBarPopup
'Private cbc                                     As CommandBarControl
'
'Public Sub Procedure_One()
'    MsgBox "Procedure One"
'End Sub
'
'Public Sub Procedure_Two()
'    MsgBox "Procedure Two"
'End Sub
'
'Public Property Get ThisMenuBar() As CommandBar
'    Set ThisMenuBar = Application.CommandBars("Worksheet Menu Bar")
'End Property
'
'Private Function ErrSrc(ByVal sProc As String) As String
'    ErrSrc = "mMenu." & sProc
'End Function
'
'Private Function CompManCommandBarPopup() As CommandBarPopup
'' ------------------------------------------------------------------------------
'' Returns the CompMan CommandBarPopup.
'' ------------------------------------------------------------------------------
'
'    Set cbr = ThisMenuBar
'    For Each cbp In cbr.Controls
'        If TypeOf cbp Is CommandBarPopup Then
'            If cbp.Caption Like "*" & MENU_NAME & "*" _
'            Or cbp.Tag = MENU_ITEM_TAG Then
'                Set CompManCommandBarPopup = cbp
'                Exit Function
'            End If
'        End If
'    Next cbp
'
'End Function
'
'Private Function Exists() As Boolean
'' ------------------------------------------------------------------------------
'' Returns TRUE when the Add-in menu contains any CommandBarPopUp items with a
'' tag = MENU_ITEM_TAG.
'' ------------------------------------------------------------------------------
'
'    Set cbr = ThisMenuBar
'    Set cbp = CompManCommandBarPopup
'    Exists = Not cbr Is Nothing And Not cbp Is Nothing
'    If Not Exists And Not cbp Is Nothing Then
'          Exists = cbp.Tag = MENU_ITEM_TAG
'    End If
'
'End Function
'
'Public Sub Remove()
'' ------------------------------------------------------------------------------
'' Removes the CompMan menu when the last open CompMan instance is closed,
'' provided it does not contain other but pending release items.
'' ------------------------------------------------------------------------------
'
'    Select Case True
'        Case mAddin.IsOpen And mMe.IsDevInstnc
'            '~~ When the dev instance is closed and there's still the Addin instance open
'            '~~ the CompMan menu is not removed
'        Case mMe.IsAddinInstnc And mWbk.IsOpen(mCompManClient.COMPMAN_DEVLP)
'            '~~ When the Add-in instance is closed and the dev instance is still open
'            '~~ this means that the Add-in is about to be renewed. In case the renewal
'            '~~ fails the CompMan menu still needs to remain avialable.
'        Case Else
'            Clear
'    End Select
'
'End Sub
'
'Public Sub HelpConfigure()
'    Const PROC = "HelpConfigure"
'
'    Debug.Print ErrSrc(PROC) & ": " & "HelpConfigure performed"
'    mBasic.README r_base_url:=GITHUB_REPO_URL _
'                , r_bookmark:=README_CONFIG_CHANGES
'End Sub
'
'Public Sub HelpServiced()
'    Const PROC = "HelpServiced"
'
'    Debug.Print ErrSrc(PROC) & ": " & "HelpServiced performed"
'    mBasic.README r_base_url:=GITHUB_REPO_URL _
'                , r_bookmark:=README_SERVICED
'End Sub
'
'Public Sub RemoveItem(ByVal r_item As String)
'' ------------------------------------------------------------------------------
'' Removes the item (r_item)
'' ------------------------------------------------------------------------------
'
'    If Exists() Then
'        On Error Resume Next
'        cbp.Controls(r_item).Delete
'    End If
'
'End Sub
'
'Public Sub Clear()
'' ------------------------------------------------------------------------------
'' Removes any items from the Add-Ins menu.
'' ------------------------------------------------------------------------------
'    While Exists()
'        cbr.Controls(cbp.Index).Delete
'    Wend
'
'End Sub
'
'Private Sub SetupItem(ByVal s_caption As String, _
'                      ByVal s_action As String)
'' ------------------------------------------------------------------------------
''
'' ------------------------------------------------------------------------------
'    Const PROC = "SetupItem"
'
'    On Error GoTo eh
'
'    Set cbp = CompManCommandBarPopup
'    If Not cbp Is Nothing Then
'        Set cbc = cbp.Controls.Add(msoControlButton)
'        With cbc
'            .Caption = s_caption
'            On Error Resume Next
'            .OnAction = "'" & ThisWorkbook.Name & "'!" & s_action
'            .Style = msoButtonCaption
'            .Enabled = True
'            .Tag = MENU_ITEM_TAG
'            On Error GoTo eh
'        End With
'    End If
'
'xt: Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub
'
