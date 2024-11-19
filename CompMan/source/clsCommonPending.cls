VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------------
' Class Module clsCommonPending: Management of pending release Common Components of
' ============================== which the Export-File is temporarily maintained in a
' dedicated "PendingReleases" folder within the Common-Components folder until the
' Common Component is releasesd, i.e. moved into the Common Component folder from
' where it is then available for VB-projects using the comoponent. The class module
' maintains a "PendingReleases.dat" Private Profile file in the CommonPending folder
' with the release relevant properties:
' - Modified in Workbook (LastModInWrkbkFullName, LastModInWrkbkFullName)
' - Current Export-File pending release (LastModExpFileOrigin)
' - Last modification date-time (LastModAt)
' - Last modified on computer (LastModOn)
' - Current Export-file name in the "PendingReleases" folder (ExportFileName)
'
' Public methods:
' ---------------
' Manage          Manages registration, conflicting modification, and removal of
'                 obsolete registrations
' Register        Registers an Exportfile as pending release by copying it into the
'                 "PendingReleases" folder and updating the corresponding properties
'                 in the CommonPending.dat Private Profile file.
' Remove          Removes a released component from the Private Profile file
'                 "PendingReleases.dat". Whne the last one registered had been
'                 removed also removes the "PendingReleases" folder.
' Release         Called by Manage. Moves the pending release Export File from the
'                 "PendingReleases" folder into the Common Components folder by
'                 updating the corresponding Properties in the CommComps.dat Private
'                 Profile file.
' -----------------------------------------------------------------------------------
' Components      Returns a Collection with the component name of all registerd
'                 pending release components, i.e. sections in the Private Profile
'                 file "PendingReleases.dat")
' ConflictResolve Displays a dialog providing the following choices:
'                 - Display the modifications already made in the other Workbook
'                 - Continue with former (pending release) modifications. I.e. the
'                   pending release Export File is re-imported to the serviced
'                   Workbook in order to enable continuation with ongoing modifi-
'                   cations.
'                 - Display the - now conflicting - modifications made in this
'                   Workbook
'                 - Ignore the former modifications. I.e. this modifications become the
'                   new pending release ones.
' -----------------------------------------------------------------------------------
Private PPFile          As clsPrivProf
Private sComp           As String
Private dctComponents   As Dictionary

Private Sub Class_Initialize()
' -----------------------------------------------------------------------------------
' Provide the "PendingReleases.dat" Private Profile file's name
' -----------------------------------------------------------------------------------
    
    Set PPFile = New clsPrivProf
    With PPFile
        .FileName = mEnvironment.CommCompsPendingPrivProfFileFullName
        .HskpngNames .FileName _
                   , VALUE_NAME_LAST_MOD_AT & "<Modified_AtDateTime" _
                   , VALUE_NAME_LAST_MOD_BY & "<Modified_ByUser" _
                   , VALUE_NAME_LAST_MOD_EXP_FILE_ORIG & "<Modified_ExportFileFullName" _
                   , VALUE_NAME_LAST_MOD_IN & "<Modified_InWorkbookFullName" _
                   , VALUE_NAME_LAST_MOD_ON & "<Modified_OnMachine"
        .FileHeader = FileHeader
        .FileFooter = mEnvironment.PrivateProfileFileFooter
    End With
    
End Sub

Private Sub Class_Terminate()
    Set PPFile = Nothing
End Sub

Friend Property Let CompName(ByVal s As String)
    PPFile.Section = s
End Property

Public Property Get Components() As Dictionary
    Dim dct As Dictionary
    Dim v   As Variant
    
    Set dct = PPFile.SectionNames
    For Each v In dct
        If v Like "_*" Then dct.Remove v
        If v Like "@*" Then dct.Remove v
    Next v
    Set dctComponents = mBasic.KeySort(dct)
    Set Components = dctComponents
    Set dct = Nothing
    
End Property

Friend Property Get LastModAt(Optional ByVal m_comp As String = vbNullString) As String
    LastModAt = Value(m_comp, VALUE_NAME_LAST_MOD_AT)
End Property

Friend Property Let LastModAt(Optional ByVal m_comp As String = vbNullString, _
                                       ByVal p_mod_last As String)
    Value(m_comp, VALUE_NAME_LAST_MOD_AT) = p_mod_last
End Property

Friend Property Get LastModBy(Optional ByVal m_comp As String = vbNullString) As String
    LastModBy = Value(m_comp, VALUE_NAME_LAST_MOD_BY)
End Property

Friend Property Let LastModBy(Optional ByVal m_comp As String = vbNullString, _
                                       ByVal p_mod_at As String)
    Value(m_comp, VALUE_NAME_LAST_MOD_BY) = p_mod_at
End Property

Friend Property Get LastModExpFileOrigin(Optional ByVal m_comp As String = vbNullString) As String
    LastModExpFileOrigin = Value(m_comp, VALUE_NAME_LAST_MOD_EXP_FILE_ORIG)
End Property

Friend Property Let LastModExpFileOrigin(Optional ByVal m_comp As String = vbNullString, _
                                                  ByVal p_mod_exp_file As String)
    Value(m_comp, VALUE_NAME_LAST_MOD_EXP_FILE_ORIG) = p_mod_exp_file
End Property

Friend Property Get LastModInWrkbkFullName(Optional ByVal m_comp As String = vbNullString) As String
    LastModInWrkbkFullName = Value(m_comp, VALUE_NAME_LAST_MOD_IN)
End Property

Friend Property Let LastModInWrkbkFullName(Optional ByVal m_comp As String = vbNullString, _
                                              ByVal p_mod_by As String)
    Value(m_comp, VALUE_NAME_LAST_MOD_IN) = p_mod_by
End Property

Friend Property Let LastModKindOfComp(Optional ByVal comp_name As String = vbNullString, _
                                               ByVal comp_reg_state As enKindOfComp)
    Value(comp_name, VALUE_NAME_REG_STAT_COMP) = mCompMan.RegState(comp_reg_state)
End Property

Friend Property Get LastModKindOfComp(Optional ByVal comp_name As String) As enKindOfComp
    LastModKindOfComp = mCompMan.RegState(Value(comp_name, VALUE_NAME_REG_STAT_COMP))
End Property

Friend Property Get LastModOn(Optional ByVal m_comp As String = vbNullString) As String
    LastModOn = Value(m_comp, VALUE_NAME_LAST_MOD_ON)
End Property

Friend Property Let LastModOn(Optional ByVal m_comp As String = vbNullString, _
                                          ByVal p_mod_at As String)
    Value(m_comp, VALUE_NAME_LAST_MOD_ON) = p_mod_at
End Property

Public Property Get Releases() As Collection:                           Set Releases = Components:                                       End Property

Private Property Get Value(Optional ByVal v_section_name As String = vbNullString, _
                           Optional ByVal v_value_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    If v_section_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No component-name provided!"
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    Value = PPFile.Value(v_value_name, v_section_name)

End Property

Private Property Let Value(Optional ByVal v_section_name As String = vbNullString, _
                           Optional ByVal v_value_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    If v_section_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No component-name provided!"
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    PPFile.Value(v_value_name, v_section_name) = v_value
                                    
End Property

Private Function Aligned(ByVal a_comp As String)
    Static l As Long: If l = 0 Then l = Serviced.MaxLenItem
    Aligned = mBasic.Align(a_comp, enAlignLeft, l, ".", " ")
End Function

Private Function BtoS(ByVal b As Boolean) As String
    If b Then BtoS = "True" Else BtoS = "False"
End Function

Private Function BttnDsplyDiffTheseVersusFormerModifications(ByVal b_comp As String) As String
    BttnDsplyDiffTheseVersusFormerModifications = "Display the modifications made  in " & vbLf & _
                                                  "this serviced ActiveWorkbook (""" & vbLf & _
                                                  Serviced.Wrkbk.Name & """) " & vbLf & _
                                                  "versus those already made in Workbook """ & vbLf & _
                                                  LastModInWrkbkName(b_comp) & vbLf & """ on computer """ & LastModOn(b_comp) & """"
End Function

Private Function BttnDsplyFormerModifications(ByVal b_comp As String) As String
    BttnDsplyFormerModifications = "Display the modifications aleady made in " & vbLf & _
                                   """ " & LastModInWrkbkName(b_comp) & """" & vbLf & _
                                   "on computer" & vbLf & _
                                   """" & LastModOn(b_comp) & """"
End Function

Private Function BttnDsplyTheseModifications() As String
    BttnDsplyTheseModifications = "Display the modifications" & vbLf & "made in this Workbook"
End Function

Private Function BttnIgnore() As String
    BttnIgnore = "Ignore/suspend this issue"
End Function

Private Function BttnIgnoreTheseModifications() As String
    BttnIgnoreTheseModifications = "Ignore the modifications" & vbLf & "made in Workbook" & vbLf & Services.ServicedWbk.Name
End Function

Private Function BttnUndoThisModifications() As String
    BttnUndoThisModifications = "Undo these modifications"
End Function

Private Function CommCompPendingModified(ByVal c_comp As clsComp) As Boolean

    With c_comp
        If mFact.CommCompHasModificationPendingRelease(.CompName) Then
            CommCompPendingModified = .CodeCrrent.Meets(.CodePnding) = False
        End If
    End With
        
End Function

Public Function CommonComponentsPendingReadyForRelease() As Collection
' ------------------------------------------------------------------------------
' Returns a collection of all pending release Common Components which are ready
' for being release. The fact that only the PendingReleases folder and the
' PendingReleases.dat file, both in the Common-Components folder, are involved
' in the release process, all pending components which are a component of the
' current serviced Workbook are regarded ready for release.
' However, when the logged in user differs from the user who registered the
' pending release the release has to be explicitely confirmed.
' ------------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    For Each v In Components
        If Serviced.CompExists(v) Then cll.Add v
    Next v
    Set CommonComponentsPendingReadyForRelease = cll
    Set cll = Nothing
    
End Function

Public Sub ConflictResolve(ByVal c_comp As clsComp)
' ------------------------------------------------------------------------------
' Displays a dialog providing various difference dsiplay options and finally
' only allow discarding those made in the serviced Workbook.
' Note: The option undo requires a public code version. In case of a pending
'       release there may not be one and thus the conflict cannot be resolved
'       with undo.
' ------------------------------------------------------------------------------
    Const PROC = "ConflictResolve"
    
    On Error GoTo eh
    Dim Msg         As mMsg.udtMsg
    Dim cllBttns    As Collection
    
    sComp = c_comp.CompName
    Set cllBttns = mMsg.Buttons(BttnDsplyFormerModifications(sComp) _
                              , BttnDsplyDiffTheseVersusFormerModifications(sComp))
    With Msg
        With .Section(1).Text
            .Text = "The component  " & mBasic.Spaced(c_comp.CompName) & " is already registered pending release " & _
                         "by another Workbook." & vbLf
            If Not c_comp.IsCommCompPublic Then
                .Text = .Text & "This issue cannot be resolved because there's yet no public code version " & _
                        "(the pending release will relase the component to public for the vry first time."
            End If
        End With
        With .Section(2)
            .Label.Text = "Concurrent modification:"
            .Label.FontBold = True
            .Label.FontColor = rgbRed
            .Text.MonoSpaced = True
            .Text.Text = "Component   :  " & mBasic.Spaced(c_comp.CompName) & "  has already been modified" & vbLf & _
                         "in Workbook : " & LastModInWrkbkFullName(sComp) & vbLf & _
                         "at date(utc): " & LastModAt(sComp)
        End With
        With .Section(3)
            .Label.Text = "Option Display this most recently made modifications:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Displays the difference between the current valid/released code versus these most resently made."
        End With
        With .Section(4)
            .Label.Text = "Option Display the already made modifications:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Displays the difference between the current valid/released code versus those currently regigsterd waiting for rease."
        End With
        With .Section(5)
            .Label.Text = "Option Display the difference between the most recently made versus the already made modifications:"
            .Label.FontColor = rgbBlue
            .Text.Text = "This difference will enable to integrate (merge) the already made modifications with those most recently made in " & _
                         "order not to loose the others when finally the option ignore them is used."
        End With
        If c_comp.IsCommCompPublic Then
            With .Section(6)
                .Label.Text = BttnUndoThisModifications & ":"
                .Label.FontColor = rgbBlue
                .Text.Text = "This is the only choice! The Common Component may only be modified is this Workbook when it is based on an up-to-date " & _
                             "version and unless there's not already a modification pending release made in another Workbook."
            End With
            Set cllBttns = mMsg.Buttons(cllBttns, vbLf, BttnUndoThisModifications)
        Else
            With .Section(6)
                .Label.Text = BttnIgnore & ":"
                .Label.FontColor = rgbBlue
                .Text.Text = "This is in fact the only only choice! The pending release Common Component should be released to public asap. " & _
                             "The component had been recognized as a used Common Component due to a known pendig release. However, the component " & _
                             "is yet not ready for being used since this requires an available public code version. In case the component's code " & _
                             "is modified this dialog will popup again when the Workbook is saved!"
            End With
            Set cllBttns = mMsg.Buttons(cllBttns, vbLf, BttnIgnore)

        End If
    End With
    Do
        Select Case mMsg.Dsply(dsply_title:="Severe conflict with an already made code modification in another Workbook!" _
                             , dsply_msg:=Msg _
                             , dsply_Label_spec:="R130" _
                             , dsply_width_min:=40 _
                             , dsply_buttons:=cllBttns)
            Case BttnDsplyTheseModifications
                DsplyDiffTheseModificationsVersusCurrentPublic c_comp
            Case BttnDsplyFormerModifications(sComp)
                DsplyDiffFormerModificationsVersusCurrentPublic c_comp
            Case BttnDsplyDiffTheseVersusFormerModifications(sComp)
                DsplyDiffTheseVersusFormerModifications c_comp
            Case BttnUndoThisModifications
                Stop
                mUpdate.ByReImport c_comp, CommonPublic.LastModExpFile(sComp)
                Serviced.Wrkbk.VBProject.VBComponents(sComp).Export c_comp.ExpFileFullName
                CommonServiced.LastModAt(sComp) = CommonPublic.LastModAt(sComp)
                Services.Log(sComp) = "Serviced Common Component's code modifications discarded by re-intantiating the " & _
                                      "current public version (update by re-import)"
                Exit Do
            Case BttnIgnore
                '~~ The component is exported but not registered pending since another one is already pending
                '~~ This dialog will popup again only when the component's code is modified
                Serviced.Wrkbk.VBProject.VBComponents(sComp).Export c_comp.ExpFileFullName
                c_comp.SetServicedProperties
        End Select
    Loop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Conflicts(ByVal c_comp As clsComp)
' -----------------------------------------------------------------------------------
' Returns TRUE when the current code of the serviced Common Component differs from
' the code change which is already pending release but the pending release
' modification has been made in another Workbook (i.e. the Workbook or the computer
' of the serviced Common Component (c_comp) is not identical with those of the
' pending release.
' Note: The conflict check may be made before or after the export of a code change.
' -----------------------------------------------------------------------------------
    Dim sIn     As String
    Dim sOn     As String
    Dim sComp   As String
    
    With c_comp
        sComp = .CompName
        Select Case True
            Case Not .IsCommCompPending
                '~~ Not modification is yet pending releaseYet
            Case .CodeCrrent.Meets(.CodePnding)
                '~~ The current pending release code modification is identical with the
                '~~ code of the service Common Component
            Case Else
                '~~ When the code differs and the properties (in Workbook or on Computer)
                '~~ differs there is a conflicting code modification
                With CommonPending
                    sIn = .LastModInWrkbkFullName(sComp)
                    sOn = .LastModOn(sComp)
                End With
                With CommonServiced
                    Conflicts = sIn <> .LastModInWrkbkFullName(sComp) _
                             Or sOn <> .LastModOn(sComp)
                End With
        End Select
    End With
                    
End Function

Private Sub DsplyDiffFormerModificationsVersusCurrentPublic(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    With d_comp
        .CodePublic.DsplyDiffs d_this_file_name:="CurrentPublicCommonComponent" _
                             , d_this_file_title:="Current public code of the Common Component  " & mBasic.Spaced(.CompName) _
                             , d_versus_code:=.CodePnding _
                             , d_versus_file_name:="CurrentPendingReleaseModifications" _
                             , d_versus_file_title:="Already pending release modifications made in Workbook " & LastModInWrkbkFullName(.CompName) & " (" & LastModOn(.CompName) & ")"
    End With
    
End Sub

Public Sub DsplyDiffTheseModificationsVersusCurrentPublic(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim sComp As String
    sComp = d_comp.CompName
    With d_comp
        .CodeCrrent.Source = .ExpFileFullName
        .CodeCrrent.DsplyDiffs d_this_file_name:="ModificationsInThisWorkbook" _
                              , d_this_file_title:="Current exported last modifications of the Common Component  " & mBasic.Spaced(sComp) _
                              , d_versus_code:=.CodePublic _
                              , d_versus_file_name:="CurrentPublicCommonComponent" _
                              , d_versus_file_title:="Current public code of Common Component  " & mBasic.Spaced(d_comp.CompName)
    End With
    
End Sub

Private Sub DsplyDiffTheseVersusFormerModifications(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim sComp As String
    
    With d_comp
        sComp = .CompName
        .CodeCrrent.Source = .ExpFileFullName
        .CodePublic.Source = CommonPublic.LastModExpFile(.CompName)
        .CodeCrrent.DsplyDiffs d_this_file_name:="ModificationsInThisWorkbook" _
                                 , d_this_file_title:="Current exported last modifications of the Common Component  " & mBasic.Spaced(sComp) _
                                 , d_versus_code:=d_comp.CodePnding _
                                 , d_versus_file_name:="CurrentPendingReleaseModifications" _
                                 , d_versus_file_title:="Modifications already made for Common Component  " & mBasic.Spaced(sComp) & "   and pending release."
    End With
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommonPending" & "." & sProc
End Function

Public Function Exists(ByVal e_comp As String, _
              Optional ByRef e_mod_at_datetime As String, _
              Optional ByRef e_mod_export_file_name As String, _
              Optional ByRef e_mod_in_wbk_full_name As String, _
              Optional ByRef e_mod_in_wbk_name As String, _
              Optional ByRef e_mod_on_machine As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the component (e_comp) is pending release.
' ------------------------------------------------------------------------------
    With PPFile
        Exists = .Exists(.FileName, e_comp)
    End With
    If Exists Then
        e_mod_at_datetime = LastModAt(e_comp)
        e_mod_export_file_name = LastModExpFileOrigin(e_comp)
        e_mod_in_wbk_full_name = LastModInWrkbkFullName(e_comp)
        e_mod_on_machine = LastModOn(e_comp)
    End If

End Function

Public Function ExportFiles() As Dictionary
' ----------------------------------------------------------------------------
' Returns all export files in the Common-Components\Pending folder as
' Dictionary with the corresponding component name as key abd the file's full
' name as item.
' ----------------------------------------------------------------------------
    
    Dim dct     As New Dictionary
    Dim fld     As Folder
    Dim fle     As File
    Dim sName   As String
    Dim sExt    As String
    
    If FSo.FolderExists(mEnvironment.CommCompsPendingPath) Then
        Set fld = FSo.GetFolder(mEnvironment.CommCompsPendingPath)
        With dct
            For Each fle In fld.Files
                sName = FSo.GetBaseName(fle.Path)
                sExt = FSo.GetExtensionName(fle.Path)
                Select Case sExt
                    Case "bas", "cls", "frm"
                        If Not .Exists(sName) _
                        Then .Add sName, fle.Path
                End Select
            Next fle
        End With
    End If
    Set ExportFiles = mBasic.KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function FileHeader() As String
    FileHeader = "Common Components which had been modified and now are pending release to the public Common-Components folder." & vbLf & _
                 "The values are transferred from the CommComps.dat of the serviced Workbook in which the code had been modified." & vbLf & _
                 "- LastModAt           : Date/time of the last modification date/time (the export file's creation date repectively)" & vbLf & _
                 "- LastModBy           : User, which had made the last modification" & vbLf & _
                 "- LastModExpFileOrigin: Indicates the 'origin'! of the export file (may point to an export file not available on or not accessable by the used compunter)" & vbLf & _
                 "- LastModIn           : The Workbook/VB-Project in which the last code modification had been made (may point to a Workbook om another computer)" & vbLf & _
                 "- LastModOn           : The computer on which the last modification had been made in the above Workbook."

End Function

Public Sub Info()

    Dim s As String
    Select Case CommonComponentsPendingReadyForRelease.Count
        Case 0:    s = vbNullString
        Case 1:    s = "     See the VBE CompMan menu for 1 Common Component is pending release!"
        Case Else: s = "     See the VBE CompMan menu for " & CommonPending.Components.Count & " Common Components pending release!"
    End Select
    On Error Resume Next ' may be too long to be displayed
    Application.StatusBar = Application.StatusBar & s

End Sub

Public Function LastModExpFile(ByVal e_comp As String) As String
' ----------------------------------------------------------------------------
' Provides the full name of a component's (e_comp) last modification Export-
' File. Note that this file is identical with the LastModExpFileOrigin which
' just points to the origin file.
' ----------------------------------------------------------------------------
    LastModExpFile = mEnvironment.CommCompsPendingPath & "\" & FSo.GetFileName(LastModExpFileOrigin(e_comp))

End Function

Public Function LastModInWrkbkName(ByVal l_comp As String) As String
    LastModInWrkbkName = FSo.GetFileName(Value(l_comp, VALUE_NAME_LAST_MOD_IN))
End Function

Public Sub Manage(ByVal m_comp As clsComp)
' ------------------------------------------------------------------------------
' Manages Common Component's pending release issues:
' - Remove obsolete "PendingReleases",
' - Establish/register/re-register "PendingReleases"
' - Provide choices for a concurrent/conflicting Common Component modification.
'
' The service is called regularilly mExport.ChangedComponents when a used or
'  hosted Common Component's code modification is exported and thus is due for
'  being registered as pending release.
'
' The service is also called for housekeeping and to cover manual user
' interventions by:
' - clsCommonPublic.HskpngCompsUsedOrPrivate when the Common Component is a used one
' - clsComp.Outdated when the seviced Workbooks used or hosted Common
'   Component's code has changed and the last modified date-time therefore is
'   greater than the current public Common Component's last modified date-time
'   and pending release is thus considered.
' - clsCommonPublic.CompsServicedHosted when the Common Component claimed hosted is
'   not identical with the current public code in the Common Components folder
'   and thus is due for being registered as pending release.
' ------------------------------------------------------------------------------
    Const PROC = "Manage"
    
    On Error GoTo eh
    Dim bModifiedInAnotherWorkbookAndOrAnotherMachine   As Boolean
    Dim bPendingIsOutdated                              As Boolean
    Dim bCommCompHasBeenModifiedInTheServicedWorkbook   As Boolean
    Dim bCommCompHasModificationPendingRelease          As Boolean
    Dim bCommCompPendingModified                        As Boolean
    Dim sComp                                           As String
    Dim sModAtDateTimeUTC                               As String
    Dim sModInWbkFullName                               As String
    Dim sPendingModAtDateTimeUTC                        As String
    Dim sPendingModInWbkFullName                        As String
    Dim sPendingModOnMachine                            As String
    
'    mBasic.BoP ErrSrc(PROC), m_comp.CompName
    With m_comp
        sComp = .CompName
        If Not .IsCommComp Then GoTo xt
        
        '~~ Gather facts for decision
        bCommCompHasBeenModifiedInTheServicedWorkbook = mFact.CommCompHasBeenModifiedinTheServicedWorkbook(m_comp, sModAtDateTimeUTC, sModInWbkFullName)
        bCommCompHasModificationPendingRelease = mFact.CommCompHasModificationPendingRelease(sComp, sPendingModAtDateTimeUTC, , sPendingModInWbkFullName, , sPendingModOnMachine)
        If bCommCompHasModificationPendingRelease Then
            bCommCompPendingModified = CommCompPendingModified(m_comp)
            bModifiedInAnotherWorkbookAndOrAnotherMachine = sPendingModInWbkFullName <> .WrkbkFullName _
                                        Or sPendingModOnMachine <> mEnvironment.ThisComputersName
            bPendingIsOutdated = sPendingModAtDateTimeUTC <> .ServicedLastModAt
        End If
        
        Debug.Print ErrSrc(PROC) & ": " & Aligned(sComp) & " " & BtoS(bCommCompHasBeenModifiedInTheServicedWorkbook) _
                                                         & " " & BtoS(bCommCompHasModificationPendingRelease) _
                                                         & " " & BtoS(bCommCompPendingModified) _
                                                         & " " & BtoS(bModifiedInAnotherWorkbookAndOrAnotherMachine) _
                                                         & " " & BtoS(bPendingIsOutdated)
        Select Case True
            Case bCommCompHasBeenModifiedInTheServicedWorkbook = False _
             And bCommCompHasModificationPendingRelease = False _
             And bCommCompPendingModified = False _
             And bModifiedInAnotherWorkbookAndOrAnotherMachine = False _
             And bPendingIsOutdated = False
             
            Case bCommCompHasBeenModifiedInTheServicedWorkbook = True _
             And bCommCompHasModificationPendingRelease = True _
             And bCommCompPendingModified = False _
             And bModifiedInAnotherWorkbookAndOrAnotherMachine = False _
             And bPendingIsOutdated = False
                Services.Log(sComp) = "Serviced Common Component's code modification still pending release"
                         
            '~~ True, False
            Case bCommCompHasBeenModifiedInTheServicedWorkbook = True _
             And bCommCompHasModificationPendingRelease = False
                '~~ The code had been modified, i.e. it differs from the current public Common Component,
                '~~ and has yet not been registered pending release.
                Register m_comp
                Services.Log(m_comp) = "Serviced Common Component's code modification registered pending release"
                            
            '~~ False, True, False, True
            Case bCommCompHasBeenModifiedInTheServicedWorkbook = False _
             And bCommCompHasModificationPendingRelease = True _
             And bCommCompPendingModified = False _
             And bModifiedInAnotherWorkbookAndOrAnotherMachine = True
                '~~ 1. The Common Component's code is identical with the current public Common Components code,
                '~~    i.e. it has not been modified since the last update
                '~~ 2. The Common Component has a modification pending release
                '~~ 3. The pending code differs from the serviced Common Component's code, i.e. it must have been
                '~~    modified in another Workbook and/or on another machine
                '~~ 4. The pending release has another Workbook and or another machine registered
                '~~ Conclusion: This is a normal situation which does not require any interaction
                Services.Log(sComp) = "Pending Release Common Component (modified in another Workbook and/or another machine)"
            
            '~~ False, True, False, False
            Case bCommCompHasBeenModifiedInTheServicedWorkbook = False _
             And bCommCompHasModificationPendingRelease = True _
             And bCommCompPendingModified = False _
             And bModifiedInAnotherWorkbookAndOrAnotherMachine = False
                '~~ 1. The Common Component's code is identical with the current public Common Components code,
                '~~    i.e. it has not been modified since the last update
                '~~ 2. The Common Component has a modification pending release
                '~~ 3. The pending code is identical with the code of the serviced Common Component's code
                '~~ 4. The pending release has the serviced Workbook on the current machine registered
                '~~ Conclusion: This should never be the case, because the pending registration is removed once
                '~~             the pending modification is released
                Remove sComp
                Services.Log(sComp) = "Obsolete Pending Release Common Component removed!"
                            
            '~~ True, True, True, False, True
            Case bCommCompHasBeenModifiedInTheServicedWorkbook = True _
             And bCommCompHasModificationPendingRelease = True _
             And bCommCompPendingModified = True _
             And bModifiedInAnotherWorkbookAndOrAnotherMachine = False _
             And bPendingIsOutdated = True
                '~~ 1. The Common Component's code differs from the current public Common Components code,
                '~~    i.e. it has been modified since the last update
                '~~ 2. The Common Component's code also differs from the current pending release code
                '~~ 3. The pending release concerns the same Workbook on the same machine
                '~~ 4. The registered penfing last modification datetime is not identical with this components last mod datetime
                '~~ Conclusion: This should have already been registered along with the most recent Export!
                Stop
                Register m_comp ' re-register a consequtive modification
                Services.Log(sComp) = "Serviced Common Component's code modification re-registered pending release."
                            
            Case Else
                Debug.Print ErrSrc(PROC) & ": " & Aligned(sComp) & " " & BtoS(bCommCompHasBeenModifiedInTheServicedWorkbook) _
                                                                 & " " & BtoS(bCommCompHasModificationPendingRelease) _
                                                                 & " " & BtoS(bCommCompPendingModified) _
                                                                 & " " & BtoS(bModifiedInAnotherWorkbookAndOrAnotherMachine) _
                                                                 & " " & BtoS(bPendingIsOutdated) _
                                                                 & " " & "not covered!"
        End Select
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Register(ByVal r_comp As clsComp)
' -----------------------------------------------------------------------------------
' Registers a Common Component's (r_comp) Export-File as "pending release" by copying
' it into the "PendingReleases" folder and updating the corresponding properties in
' the "PendingReleases.dat" Private Profile file.
' Note: The PendingReleases folder exists only while there is at least one pending
'       release component.
' -----------------------------------------------------------------------------------
    Const PROC = "Register"
    
    Dim sSource As String
    Dim sTarget As String
    Dim sFolder As String
    
    sFolder = mEnvironment.CommCompsPendingPath
    With r_comp
        CompName = .CompName
        sSource = .ExpFileFullName
        sTarget = sFolder & "\" & FSo.GetFileName(.ExpFileFullName)
        With FSo
            If Not .FolderExists(sFolder) Then .CreateFolder sFolder
            If .FileExists(sSource) Then
                .CopyFile sSource, sTarget, True
            Else
                Stop
            End If
        End With
        .SetPendingEqualServiced
        Services.Log(.CompName) = "Modified Common Component: Release registered pending"
    End With

    mCompManMenuVBE.Setup
    
End Sub

Public Sub ReleaseComp(ByVal r_comp As String, _
              Optional ByVal r_skip_service_initiation As Boolean = False)
' -----------------------------------------------------------------------------------
' 1. Move a component (r_comp) from the PendingRelease folder into the
'    Common-Components folder
' 2. Update the corresponding values in the CommComps.dat
' 3. Remove the corresponding values from the "PendingReleases.dat" file - which by
'    the way removes the section
' -----------------------------------------------------------------------------------
    Const PROC = "ReleaseComp"
    
    On Error GoTo eh
    Dim sName   As String
    Dim sSource As String
    Dim sTarget As String
        
    If Not r_skip_service_initiation _
    Then mCompMan.ServiceInitiate s_serviced_wbk:=mCompMan.ServicedWrkbk _
                                , s_service:="Release Common Components" _
                                , s_do_housekeeping:=False
    
    CompName = r_comp
    With Prgrss
        .Operation = "Release"
        .ItemsTotal = Components.Count
    End With
    
    '~~ Move export file from PendingReleases folder to the Common-Components folder
    With FSo
        sName = .GetFileName(LastModExpFileOrigin(r_comp))
        sSource = mEnvironment.CommCompsPendingPath & "\" & sName
        sTarget = mEnvironment.CommCompsPath & "\" & sName ' may yet not exist when first time release !
        If .FileExists(sTarget) Then .DeleteFile sTarget
        .MoveFile sSource, sTarget
    End With
    
    '~~ Update the public Common Component properties with the pending infos
    With New clsComp
        .CompName = r_comp
        .SetPublicEqualPending
    End With
    If Not CommonPublic.All.Exists(r_comp) Then
        '~~ This is obviously the very first release to public
        CommonPublic.Add r_comp
    End If
    
    If FSo.FileExists(mEnvironment.CommCompsPendingPrivProfFileFullName) _
    Then Remove r_comp ' in case the file still exists
    Prgrss.ItemDone = r_comp

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Remove(ByVal r_comp As String)
' -----------------------------------------------------------------------------------
' Removes the component/section from the Private Profile file.
' When the removed section was the last one the Private Profile file the file will no
' longer exist after the removal and thus the "PendingReleases" folder also has
' become obsolete and will be removed too.
' -----------------------------------------------------------------------------------
    
    PPFile.SectionRemove r_comp
    If Not FSo.FileExists(PPFile.FileName) And FSo.FolderExists(mEnvironment.CommCompsPendingPath) _
    Then FSo.DeleteFolder mEnvironment.CommCompsPendingPath
    
End Sub

