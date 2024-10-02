Attribute VB_Name = "mMenu"
Option Explicit
' ------------------------------------------------------------------------------
' Standard  Module mMenu: Provides the "CompMan" menu with all its appropriate
' ======================= items.
'
' ------------------------------------------------------------------------------
Private Const MENU_ITEM_RELEASE_SERVICE         As String = "CompMan: Release new Component version"
Private Const MENU_ITEM_HELP_RELEASE_SERVICE    As String = "CompMan: Help Release"
Private Const MENU_ITEM_HELP_SERVICED           As String = "CompMan: Help Serviced"
Private Const MENU_ITEM_HELP_CONFIGURE          As String = "CompMan: Help Configure"
Private Const README_SERVICED                   As String = "#enabling-the-services-serviced-or-not-serviced"
Private Const README_RELEASE                    As String = "#the-release-service"

Private objPopUp                                As CommandBarPopup
Private dctReleaseCandidates                    As New Dictionary

Private Function CommandBar() As CommandBar
    Dim cmb As CommandBar
    Dim cbp As CommandBarPopup
    
    For Each cmb In Application.CommandBars
        For Each cbp In cmb.Controls
            If TypeOf cbp Is CommandBarPopup Then
                Debug.Print "Menu-Bar '" & cmb.Name & "': " & TypeName(cbp) & ": " & cbp.Caption
            End If
        Next cbp
    Next cmb
End Function

Public Sub HelpConfigure()
    Debug.Print "HelpConfigure performed"
    mBasic.README r_url:=GITHUB_REPO_URL _
                , r_bookmark:=README_CONFIG_CHANGES
End Sub

Public Sub HelpRelease()
    Debug.Print "HelpRelease performed"
End Sub

Public Sub HelpServiced()
    Debug.Print "HelpServiced performed"
    mBasic.README r_url:=GITHUB_REPO_URL _
                , r_bookmark:=README_SERVICED
End Sub

Public Sub ReleaseService(ByRef r_candidates As Dictionary)
' ------------------------------------------------------------------------------
' Collects all Common Components of the Serviced Workbook - when opened or
' when saved - of which a more up-to-date version is available then
' available in the Common Components folder and presents them one by one for
' - being released
' - displaying the difference
' . being skipped
' ------------------------------------------------------------------------------
    Dim v   As Variant
    
    For Each v In r_candidates
    Next v
    
End Sub

Public Function ReleaseCandidates(ByVal r_wbk As Workbook) As Dictionary
    Dim dct     As Dictionary
    Dim v       As Variant
    Dim Comp    As clsComp
    
    If CommComps Is Nothing Then Set CommComps = New clsCommComps
    Set dct = CommComps.Components
    For Each v In dct
        Set Comp = New clsComp
        With Comp
            .Wrkbk = r_wbk
            .CompName = v
            If .LastModifiedAtDatTime > .Raw.LastModifiedAtDatTime Then
                dct.Add v, Comp
            End If
        End With
        Set Comp = Nothing
    Next v
    Set ReleaseCandidates = dct
    Set dct = Nothing
    
End Function

Public Sub Remove()
    RemoveItem MENU_ITEM_RELEASE_SERVICE
    RemoveItem MENU_ITEM_HELP_RELEASE_SERVICE
    RemoveItem MENU_ITEM_HELP_SERVICED
    RemoveItem MENU_ITEM_HELP_CONFIGURE
End Sub

Public Sub RemoveItem(ByVal r_item As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls(r_item).Delete
      On Error GoTo 0
   End With
End Sub

Public Sub Setup()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Remove
    SetupItem MENU_ITEM_HELP_SERVICED, "HelpServiced"
    SetupItem MENU_ITEM_HELP_CONFIGURE, "HelpConfigure"
    
End Sub

Public Sub SetupRelease()
    mMenu.SetupItem MENU_ITEM_RELEASE_SERVICE, "ReleaseService"
    mMenu.SetupItem MENU_ITEM_HELP_RELEASE_SERVICE, "HelpRelease"
End Sub

Public Sub RemoveRelease()
    mMenu.RemoveItem MENU_ITEM_RELEASE_SERVICE
    mMenu.RemoveItem MENU_ITEM_HELP_RELEASE_SERVICE
End Sub

Private Sub SetupItem(ByVal s_name As String, _
                      ByVal s_action As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    With Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlButton)
       .Caption = s_name
       .OnAction = s_action
       .Style = msoButtonCaption
    End With

End Sub

