VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsCommComps: Management of Common Components
' ========================== a) in the dedicated Common-Components folder
'                            b) in the CommComps.dat file
' Public Methods:
' ----------------
' BttnInconsistencyExport          .
' BttnInconsistencySkip            .
' BttnInconsistencyUpdate          .
' CommCompsDatFileFullName         .
'
' Public Properties:
' ------------------
' CompName                 w Essential throughout all services/methods
' Hosted                   w Provides the hosted components, provided as a
'                            comma delimited string in a Dictionary
'                            (dctHosted) and performs the housekeeping for
'                            the Common-Components folder's CommComps.dat
'                            file
' LastMod...             r/w Properties are provided whenever a used/hosted
'                            Common Component's code has changed, along with
'                            the clsComp.ReleaseComp method.
' PendingReleaseRegister   w Modifications made for a used/hosted Common
'                            Component are registered 'pending release'.
'
' ---------------------------------------------------------------------------

' Housekeeping syntax: (allows the on-the-fly-change of value names as well as their removal)
' "<current-name>:<old-name>" = rename in all sections
' "<current>"                 = no action
' ":<remove-name>"            = remove in all sections
Private Const EXP_FILE_EXTENTION                        As String = "EXP_FILE_EXTENTION"
Private Const LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN   As String = "LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN"
Private Const LAST_MODIFIED_AT_DATE_TIME                As String = "LAST_MODIFIED_AT_DATE_TIME"
Private Const LAST_MODIFIED_ON_MACHINE                  As String = "LAST_MODIFIED_ON_MACHINE"
Private Const LAST_MODIFIED_IN_WBK_FULL_NAME            As String = "LAST_MODIFIED_IN_WBK_FULL_NAME"
Private Const LAST_MODIFIED_IN_WBK_NAME                 As String = "LAST_MODIFIED_IN_WBK_NAME"
Private Const PENDING_RELEASE_MODIFIED_ON_MACHINE       As String = "PENDING_RELEASE_MODIFIED_ON_MACHINE"
Private Const PENDING_RELEASE_MODIFIED_IN_WBK_NAME      As String = "PENDING_RELEASE_MODIFIED_IN_WBK_NAME"
Private Const PENDING_RELEASE_MODIFIED_IN_WBK_FULL_NAME As String = "PENDING_RELEASE_MODIFIED_IN_WBK_FULL_NAME"
Private Const PENDING_RELEASE_MODIFIED_LAST_DATE_TIME   As String = "PENDING_RELEASE_MODIFIED_LAST_DATE_TIME"
Private Const PENDING_RELEASE_MODIFIED_EXP_FILE         As String = "PENDING_RELEASE_EXP_FILE"

Private PPCommComps                                     As clsPP
Private dctHosted                                       As Dictionary
Private sFolderPath                                     As String
Private sCompName                                       As String
Private sCommCompsDatFile                               As String
Private CodePending                                     As New clsCode

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
    
    Dim dct As New Dictionary
    
    mBasic.BoP ErrSrc(PROC)
'    Set dctHosted = New Dictionary
    Set PPCommComps = New clsPP
    
    sFolderPath = wsConfig.FolderCommonComponentsPath

    '~~> Full name of the PPCommComps-File
    sCommCompsDatFile = sFolderPath & "\" & "CommComps.dat"
    PPCommComps.FileFullName = sCommCompsDatFile
        
    '~~> Value Names with the syntax (allows the on-the-fly-change of value names as well as their removal):
    '~~> "<current>"                 = no action
    '~~> "<current>:<old>[:<old>].." = rename any <old> to <new> in all sections
    '~~> ":<remove>"                 = remove in all sections
    With dct
        .Add "Obsolete_1", ":RawRevisionNumber"
        .Add EXP_FILE_EXTENTION, "ExportFileExtention"
        .Add LAST_MODIFIED_AT_DATE_TIME, "Last_Modified_AtDateTime:LastModifiedAtDateTime:RevisionNumber"
        .Add LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN, "LastModExpFileFullNameOrigin:LastModExpFileFullName"
        .Add LAST_MODIFIED_IN_WBK_FULL_NAME, "Last_Modified_InWbkFullName:LastModifiedInWbkFullName:LastModWbkFullName"
        .Add LAST_MODIFIED_IN_WBK_NAME, "Last_Modified_InWbkName:LastModifiedInWbkName"
        .Add LAST_MODIFIED_ON_MACHINE, "Last_Modified_OnMachine:LastModifiedOnMachine"
        .Add PENDING_RELEASE_MODIFIED_IN_WBK_NAME, "PendingRelease_Modified_InWorkbookName"
        .Add PENDING_RELEASE_MODIFIED_IN_WBK_FULL_NAME, "PendingRelease_Modified_InWorkbookFullName"
        .Add PENDING_RELEASE_MODIFIED_EXP_FILE, "PendingRelease_Modified_ExportFileFullName:PendingRelease_Modified_ExportFile"
        .Add PENDING_RELEASE_MODIFIED_LAST_DATE_TIME, "PendingRelease_Modified_AtDateTime"
        .Add PENDING_RELEASE_MODIFIED_ON_MACHINE, "PendingRelease_Modified_OnMachine"
    End With
    PPCommComps.Names = dct
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)

End Sub

Private Sub Class_Terminate()
    Set PPCommComps = Nothing
End Sub

Friend Property Get BttnDsplyDiff() As String: BttnDsplyDiff = "Display difference" & vbLf & "(of Export-Files)":   End Property

Friend Property Get BttnDsplyThatDiff() As String
    BttnDsplyThatDiff = "Display the modifications" & vbLf & "already made in Workbook" & vbLf & PendingReleaseModInWbkName
End Property

Friend Property Get BttnDsplyThisModifications() As String
    BttnDsplyThisModifications = "Display the modifications" & vbLf & "made in this Workbook"
End Property

Friend Property Get BttnDsplyThisThatDiff() As String
    BttnDsplyThisThatDiff = "Display modification difference" & vbLf & "this Workbook versus" & vbLf & "in " & PendingReleaseModInWbkName
End Property

Private Property Get BttnIgnoreThatChanges() As String
    BttnIgnoreThatChanges = "Ignore the modifications" & vbLf & "made in Workbook" & vbLf & PendingReleaseModInWbkName & ")"
End Property

Friend Property Get BttnIgnoreThisChanges() As String
    BttnIgnoreThisChanges = "Ignore the modifications" & vbLf & "made in Workbook" & vbLf & Services.ServicedWbk.Name
End Property

Public Property Let CompName(ByVal c_name As String): sCompName = c_name: End Property

Friend Property Let Hosted(ByVal s As String)
' ---------------------------------------------------------------------------
' Provides the hosted components, provided as a comma delimited string in a
' Dictionary (dctHosted) and performs the housekeeping for Common Components
' in the Common-Components folder's CommComps.dat file.
' ---------------------------------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    Dim dct     As New Dictionary
    
    Set dctHosted = Nothing: Set dctHosted = New Dictionary
    For Each v In Split(s, ",")
        sComp = Trim$(v)
        If Not dct.Exists(sComp) Then
            dctHosted.Add sComp, sComp
        End If
    Next v
    
End Property

Friend Property Get LastModAtDateTime(Optional ByVal l_comp_name As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModAtDateTime = Value(LAST_MODIFIED_AT_DATE_TIME, l_comp_name)
End Property

Friend Property Let ExportFileExtention(Optional ByVal l_comp_name As String = vbNullString, _
                                                 ByVal l_ext As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(EXP_FILE_EXTENTION, l_comp_name) = l_ext
End Property
                                                 
Friend Property Get ExportFileExtention(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    ExportFileExtention = Value(EXP_FILE_EXTENTION, l_comp_name)
End Property

Friend Property Let LastModAtDateTime(Optional ByVal l_comp_name As String = vbNullString, _
                                               ByVal l_comp_rev_no As String)
' ------------------------------------------------------------------------------
' Returns a revision number in the form yy-mm-dd.00. Plus one when an existing
' revision number is provided (comp_rev-no) or the current date with .01.
' ------------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    If l_comp_rev_no = vbNullString Then l_comp_rev_no = CompMan.UTC
    Value(LAST_MODIFIED_AT_DATE_TIME, l_comp_name) = l_comp_rev_no
End Property

Friend Property Get LastModExpFile(Optional ByVal l_comp_name As String = vbNullString) As File
' ---------------------------------------------------------------------------
' Returns a component's (comp_name) Export-File in the Common-Components
' folder as file object.
' ---------------------------------------------------------------------------
    Const PROC = "LastModExpFile Get"
    
    On Error GoTo eh
    Dim fl As File
    
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    With fso
        For Each fl In .GetFolder(sFolderPath).Files
            If .GetBaseName(fl.Path) = l_comp_name Then
                Set LastModExpFile = fl
                Exit For
            End If
        Next fl
    End With
    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let LastModExpFile(Optional ByVal l_comp_name As String = vbNullString, _
                                             ByVal l_comp_exp_file As File)
' ---------------------------------------------------------------------------
' Copies the Raw Export File from its host Workbook location to the Common
' Components Folder from where it is used as the source for the update of
' outdated Used Common Components.
' ---------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    l_comp_exp_file.Copy Destination:=sFolderPath & l_comp_exp_file.Name, OverWriteFiles:=True
    Me.LastModExpFileFullNameOrigin(l_comp_name) = l_comp_exp_file.Path
End Property

Friend Property Get LastModExpFileFullName(Optional ByVal l_comp_name As String = vbNullString) As String
' ---------------------------------------------------------------------------
' Returns the Common-Component folder's Export-File Full Name.
' ---------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    With fso
        LastModExpFileFullName = sFolderPath & "\" & l_comp_name & "." & ExportFileExtention(l_comp_name)
    End With
End Property

Friend Property Get LastModExpFileFullNameOrigin(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModExpFileFullNameOrigin = Value(LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN, l_comp_name)
End Property

Friend Property Let LastModExpFileFullNameOrigin(Optional ByVal l_comp_name As String = vbNullString, _
                                                          ByVal l_exp_file_full_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN, l_comp_name) = l_exp_file_full_name
End Property

Friend Property Let LastModInWbk(Optional ByVal l_comp_name As String = vbNullString, _
                                          ByVal l_wbk As Workbook)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name) = l_wbk.FullName
    Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name) = l_wbk.Name
    
End Property

Friend Property Get LastModInWbkBaseName(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModInWbkBaseName = fso.GetBaseName(Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name))
End Property

Friend Property Get LastModInWbkFullName(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModInWbkFullName = Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name)
End Property

Friend Property Let LastModInWbkFullName(Optional ByVal l_comp_name As String = vbNullString, _
                                                  ByVal l_full_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name) = l_full_name
End Property

Friend Property Get LastModInWbkName(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModInWbkName = Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name)
    If LastModInWbkName = vbNullString Then
        LastModInWbkName = fso.GetFileName(Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name))
    End If
End Property

Friend Property Let LastModInWbkName(Optional ByVal l_comp_name As String = vbNullString, _
                                              ByVal l_wbk_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name) = l_wbk_name
End Property
Friend Property Get LastModOnMachine(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModOnMachine = Value(LAST_MODIFIED_ON_MACHINE, l_comp_name)
End Property

Friend Property Let LastModOnMachine(Optional ByVal l_comp_name As String = vbNullString, _
                                              ByVal l_machine_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_ON_MACHINE, l_comp_name) = l_machine_name
End Property

Public Property Let PendingReleaseRegister(ByVal c_comp As clsComp)
' ------------------------------------------------------------------------------
' Registers a Common Component's pending release, done with each export for
' a modified used/hosted Common Component.
' ------------------------------------------------------------------------------
    Const PROC = "PendingReleaseRegister/Let"
    
    On Error GoTo eh
'    Stop
    With c_comp
        CompName = .CompName
        PendingReleaseModInWbkName = .Wrkbk.Name
        PendingReleaseModInWbkFullName = .Wrkbk.FullName
        PendingReleaseModExpFileFullName = .ExpFileFullName
        PendingReleaseModAtDateTime = .LastModAtDateTime
        PendingReleaseModOnMachine = Environ("COMPUTERNAME")
        ExportFileExtention = fso.GetExtensionName(.ExpFileFullName)
        Debug.Print "Pending release registered: " & sCompName
    End With
    
xt: Exit Property
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get PendingReleaseModAtDateTime(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModAtDateTime = Value(PENDING_RELEASE_MODIFIED_LAST_DATE_TIME, p_comp_name)
End Property

Friend Property Let PendingReleaseModAtDateTime(Optional ByVal p_comp_name As String = vbNullString, _
                                                         ByVal p_mod_last As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(PENDING_RELEASE_MODIFIED_LAST_DATE_TIME, p_comp_name) = p_mod_last
End Property

Friend Property Get PendingReleaseModExpFileFullName(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModExpFileFullName = Value(PENDING_RELEASE_MODIFIED_EXP_FILE, p_comp_name)
End Property

Friend Property Let PendingReleaseModExpFileFullName(Optional ByVal p_comp_name As String = vbNullString, _
                                                              ByVal p_mod_exp_file As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(PENDING_RELEASE_MODIFIED_EXP_FILE, p_comp_name) = p_mod_exp_file
End Property

Friend Property Get PendingReleaseModInWbkFullName(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModInWbkFullName = Value(PENDING_RELEASE_MODIFIED_IN_WBK_FULL_NAME, p_comp_name)
End Property

Friend Property Let PendingReleaseModInWbkFullName(Optional ByVal p_comp_name As String = vbNullString, _
                                                            ByVal p_mod_by As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(PENDING_RELEASE_MODIFIED_IN_WBK_FULL_NAME, p_comp_name) = p_mod_by
End Property

Friend Property Get PendingReleaseModInWbkName(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModInWbkName = Value(PENDING_RELEASE_MODIFIED_IN_WBK_NAME, p_comp_name)
End Property

Friend Property Let PendingReleaseModInWbkName(Optional ByVal p_comp_name As String = vbNullString, _
                                                        ByVal p_mod_in_wbk_name As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(PENDING_RELEASE_MODIFIED_IN_WBK_NAME, p_comp_name) = p_mod_in_wbk_name
End Property

Friend Property Get PendingReleaseModOnMachine(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModOnMachine = Value(PENDING_RELEASE_MODIFIED_ON_MACHINE, p_comp_name)
End Property

Friend Property Let PendingReleaseModOnMachine(Optional ByVal p_comp_name As String = vbNullString, _
                                                 ByVal p_mod_at As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(PENDING_RELEASE_MODIFIED_ON_MACHINE, p_comp_name) = p_mod_at
End Property

Private Property Get Value(Optional ByVal v_value_name As String = vbNullString, _
                           Optional ByVal v_comp_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    On Error GoTo eh
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No value-name provided!"
    
    If v_comp_name = vbNullString Then
        If sCompName = vbNullString Then
            Err.Raise AppErr(2), ErrSrc(PROC), "No component name provided, section name is missing!"
        Else
            v_comp_name = sCompName
        End If
    End If
    With PPCommComps
        Value = .Value(pp_section:=v_comp_name, pp_value_name:=.Name(v_value_name))
    End With

xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value(Optional ByVal v_value_name As String, _
                           Optional ByVal v_comp_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    On Error GoTo eh
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No value-name provided!"
    
    If v_comp_name = vbNullString Then
        If sCompName = vbNullString Then
            Err.Raise AppErr(1), ErrSrc(PROC), "No component name provided, section name is missing!"
        Else
            v_comp_name = sCompName
        End If
    End If
    With PPCommComps
        .Value(pp_section:=v_comp_name, pp_value_name:=.Name(v_value_name)) = v_value
    End With
                                    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Sub ComponentRemove(ByVal s As String)
     PPCommComps.ComponentRemove s
End Sub

Public Function Components() As Dictionary
    Const PROC = "Components"
    
    mBasic.BoP ErrSrc(PROC)
    Set Components = mFso.PPsections(pp_file:=sCommCompsDatFile)
    mBasic.EoP ErrSrc(PROC)
    
End Function

Public Property Get CurrentPublicExpFileFullName() As String
    With fso
        Select Case True
            Case .FileExists(sFolderPath & "\" & sCompName & ".bas"): CurrentPublicExpFileFullName = sFolderPath & "\" & sCompName & ".bas"
            Case .FileExists(sFolderPath & "\" & sCompName & ".cls"): CurrentPublicExpFileFullName = sFolderPath & "\" & sCompName & ".cls"
            Case .FileExists(sFolderPath & "\" & sCompName & ".frm"): CurrentPublicExpFileFullName = sFolderPath & "\" & sCompName & ".frm"
        End Select
    End With
End Property

Public Function ConcurrentModification(ByVal c_comp As clsComp) As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    Dim Msg As mMsg.udtMsg
    
    sCompName = c_comp.CompName
    
    With Msg
        With .Section(1)
            .Label.Text = "Concurrent modification"
            .Label.FontBold = True
            .Label.FontColor = rgbRed
            .Text.Text = _
            "The Common Component  " & mBasic.Spaced(sCompName) & "  has already been modified in Workbook  " & _
            mBasic.Spaced(PendingReleaseModInWbkName) & " on computer  " & PendingReleaseModOnMachine & "." & _
            "When the modifications made in this Workbook are confirmed, all changes made for the Common " & _
            "Component in the other Workbook are ignored. I.e. these modifications are given priority over those made " & _
            "already elsewhere. Comparing the modifications may support the decision on how to proceed."
        End With
        With .Section(3)
            .Label.Text = "-----------------------------"
            .Text.Text = "The already made modification facts:"
        End With
        With .Section(3)
            .Label.Text = "Modified in Workbook:"
            .Label.FontBold = True
            .Label.FontColor = rgbBlue
            .Text.Text = CommComps.LastModInWbkFullName
            .Text.MonoSpaced = True
        End With
        With .Section(4)
            .Label.Text = "Modified on Machine:"
            .Label.FontBold = True
            .Label.FontColor = rgbBlue
            .Text.Text = LastModOnMachine
            .Text.MonoSpaced = True
        End With
        With .Section(5)
            .Label.Text = "Modified (exported) at:"
            .Label.FontBold = True
            .Label.FontColor = rgbBlue
            .Text.Text = LastModAtDateTime
            .Text.MonoSpaced = True
        End With
        With .Section(6)
            .Label.Text = "-----------------------------"
            .Text.Text = "The options (of which finally one will to be taken)!"
        End With
        With .Section(7)
            .Label.Text = BttnIgnoreThisChanges
            .Label.FontColor = rgbBlue
            .Text.Text = "The changes already made in the other Workbook and/or the other machine remain those pending release."
        End With
        With .Section(8)
            .Label.Text = BttnIgnoreThatChanges
            .Label.FontColor = rgbBlue
            .Text.Text = "The changes made in this Workbook become those pending release. I.e. the " & _
                         "other changes are ignored and will be replaced by thsi changes once released."
        End With
    End With
    Do
        Select Case mMsg.Dsply(dsply_title:="Severe danger of an already made code modification!" _
                             , dsply_msg:=Msg _
                             , dsply_Label_spec:="R130" _
                             , dsply_width_min:=30 _
                             , dsply_buttons:=mMsg.Buttons(BttnDsplyThisModifications, vbLf, _
                                                           BttnDsplyThatDiff, vbLf, _
                                                           BttnDsplyThisThatDiff, vbLf, _
                                                           BttnIgnoreThisChanges, vbLf, _
                                                           BttnIgnoreThatChanges))

                    Case BttnDsplyThisModifications:    DsplyDiffsOfTheCurrentModificationsVersusTheCurrentPubliicCode c_comp
                    Case BttnDsplyThatDiff:             DsplyDiffsOfTheCurrentPendingReleaseModificationsVersusTheCurrentPublicCode c_comp
                    Case BttnDsplyThisThatDiff:         DsplyDiffsOfTheCurrentModificationsVersusThePendingReleaseModifications c_comp
                    Case BttnIgnoreThatChanges:         PendingReleaseRegister = c_comp
                                                        Exit Do
                    Case BttnIgnoreThisChanges:         Exit Do ' changes made in the other Workbook remain pending release
                End Select
            Loop
    
End Function

Private Sub DsplyDiffsOfTheCurrentPendingReleaseModificationsVersusTheCurrentPublicCode(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    With d_comp.CodePublic
        .DsplyDiffs d_this_file_name:="CurrentPublicCommonComponent" _
                     , d_this_file_title:="Current public code of the Common Component  " & mBasic.Spaced(sCompName) _
                     , d_from_code:=CodePending _
                     , d_from_file_name:="CurrentPendingReleaseModifications" _
                     , d_from_file_title:="Pending release modifications for Common Component  " & mBasic.Spaced(sCompName) & "   already made in Workbook " & PendingReleaseModInWbkName
    End With
    
End Sub

Public Sub DsplyDiffsOfTheCurrentModificationsVersusTheCurrentPubliicCode(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
                    
    With d_comp
        .CodeCurrent.Source = .ExpFileFullName
        .CodeCurrent.DsplyDiffs d_this_file_name:="ModificationsInThisWorkbook" _
                                 , d_this_file_title:="Current exported last modifications of the Common Component  " & mBasic.Spaced(sCompName) _
                                 , d_from_code:=.CodePublic _
                                 , d_from_file_name:="CurrentPublicCommonComponent" _
                                 , d_from_file_title:="Current public code of Common Component  " & mBasic.Spaced(sCompName)
    End With
    
End Sub

Private Sub DsplyDiffsOfTheCurrentModificationsVersusThePendingReleaseModifications(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
                    
    With d_comp
        .CodeCurrent.Source = .ExpFileFullName
        .CodePublic.Source = PendingReleaseModExpFileFullName
        .CodeCurrent.DsplyDiffs d_this_file_name:="ModificationsInThisWorkbook" _
                                 , d_this_file_title:="Current exported last modifications of the Common Component  " & mBasic.Spaced(sCompName) _
                                 , d_from_code:=CodePending _
                                 , d_from_file_name:="CurrentPendingReleaseModifications" _
                                 , d_from_file_title:="Modifications already made for Common Component  " & mBasic.Spaced(sCompName) & "   and pending release."
    End With
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommComps" & "." & sProc
End Function

Public Function ExistsPublic(ByVal e_comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Common-Components folder has a an Export-File for
' the named component (e_comp_name), i.e. there is a public code version
' available.
' ----------------------------------------------------------------------------
    ExistsPublic = LastModInWbkFullName(e_comp_name) <> vbNullString
End Function

Public Function ExistsRegistered(ByVal x_comp_name As String) As Boolean
    ExistsRegistered = Components.Exists(x_comp_name)
End Function

Public Sub Hskpng()
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    Dim dct As Dictionary
    Dim wbk As Workbook
    Dim fld As Folder
    Dim fls As New Dictionary
    Dim fl  As File
    Dim s   As String
    
    mBasic.BoP ErrSrc(PROC)
    If Services Is Nothing Then
        Set Services = New clsServices
        Services.Initiate i_service_proc:=mCompManClient.SRVC_RELEASE_PENDING_CHANGES _
                        , i_serviced_wbk:=ThisWorkbook
    End If
    
    '~~ Collect the base-name of all files in the folder each representing
    '~~ an available public Common Component's code
    Set fld = fso.GetFolder(sFolderPath)
    With fls
        For Each fl In fld.Files
            s = fso.GetBaseName(fl.Path)
            If Not .Exists(s) _
            Then .Add s, vbNullString
        Next fl
    End With

    Set wbk = Services.ServicedWbk
    Set dct = Components
    
    Set Prgrss = New clsProgress
    With Prgrss
        .Operation = "Hskpng"
        .ItemsTotal = dct.Count _
                    + fld.Files.Count _
                    + wbk.VBProject.VBComponents.Count _
                    + dctHosted.Count
        .Figures = False
        .DoneItemsInfo = False
    End With
    HskpngSectionsObsolete dct, fls
    HskpngSectionsMissing fld, dct
    HskpngUsedOrPrivate wbk
    HskpngHosted
    Reorg
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngHosted()
' ----------------------------------------------------------------------------
' Maintains for each Commmon Component claimed hosted by the serviced Workbook
' the following values in the Common-Components folder's CommComp.dat file
' when yet not provided:
' - Component Name (section)
' - LastModifiedInWbkFullName (value)
' - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "HskpngHosted"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim sHostBaseName   As String
    Dim wbk             As Workbook
    Dim Comp            As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    sHostBaseName = fso.GetBaseName(wbk.FullName)
    
    For Each v In dctHosted
        If Not mComp.Exists(v, wbk) Then
            MsgBox "The VBComponent " & v & " claimed hosted by the serviced Workbook " & Services.ServicedWbk.Name & _
                   " will be ignored because it does not exist in the Workbook/VB-Project." & vbLf & vbLf & _
                   "The hosted argument should be updated accordingly.", _
                   vbOK, "VBComponent " & v & "does not exist!"
            Prgrss.ItemSkipped
        Else
            '~~ When the Common Component claimed hosted is not identical with the current public code
            '~~ or it has yet never been released public it will here be registered for being released.
            Set Comp = New clsComp
            With Comp
                .Wrkbk = wbk
                .CompName = v
            End With
            PendingReleaseManagement Comp
            Prgrss.ItemDone = Comp.CompName
            Set Comp = Nothing
        End If
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngSectionsMissing(ByVal h_fld As Folder, _
                                  ByVal h_comps As Dictionary)
' ------------------------------------------------------------------------------
' Adds for any Export-File in the Common-Components folder a section in the
' Common-Component folder's CommComps.dat file a section.
'
' Background:
' A missing section indicates a Common Component of wich the Export-File has
' obvously been copied manually into the Common-Components folder which now in
' the sense of CompMan has become an available Common Component ready for being
' imported into any VB-Project. A new registered Common Component remains
' un-hosted until a Workbbok claims hosting it, i.e. providing a delevelopment
' and test environment for it.
' ------------------------------------------------------------------------------
    Const PROC = "HskpngSectionsMissing"
    
    On Error GoTo eh
    Dim fle         As File
    Dim sCompName   As String
    Dim sExt        As String
    
    mBasic.BoP ErrSrc(PROC)
    With fso
        For Each fle In h_fld.Files
            sExt = .GetExtensionName(fle.Path)
            Select Case sExt
                Case "bas", "frm", "cls"
                    sCompName = .GetBaseName(fle.Path)
                    If Not h_comps.Exists(sCompName) Then
                        '~~ The Export-File is yet not registered as a known Common Component
                        '~~ It most likely has been copied manually into the Common-Components
                        '~~ folder. I.e. its "raw host" is unknown - and registered as such.
                        '~- The raw host will remain unknown until the Common Component is
                        '~~ modified in a Workbook using it and exported.
                        LastModInWbk(sCompName) = Nothing
                        LastModAtDateTime(sCompName) = CompMan.UTC
                        Prgrss.ItemDone = sCompName
                    Else
                        Prgrss.ItemSkipped
                        If LastModExpFileFullNameOrigin(sCompName) = vbNullString Then
'                            Debug.Print "The property ""LastModExpFileFullNameOrigin"" of component " & sCompName & " is not available, i.e. its origin is unknown or simply yet not registered respectively!"
                        End If
                    End If
                Case Else
                    Prgrss.ItemSkipped
            End Select
        Next fle
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngSectionsObsolete(ByVal h_comps As Dictionary, _
                                   ByVal h_files As Dictionary)
' ------------------------------------------------------------------------------
' Removes in the Common-Component folder's CommComps.dat file any section of
' which no corresponding Export-File exists in the folder
' ------------------------------------------------------------------------------
    Const PROC = "HskpngSectionsObsolete"
    
    On Error GoTo eh
    Dim sSection    As String
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
        
    '~~ Remove any component of which the corresponding Export-File
    '~~ not exists in the Common-Components folde
    For Each v In h_comps
        sSection = v
        If sSection <> PPCommComps.NamesHousekeepingSection Then
            If Not h_files.Exists(v) Then
                PPCommComps.ComponentRemove sSection
                Prgrss.ItemDone = sSection
            Else
                Prgrss.ItemSkipped
            End If
        Else
            Prgrss.ItemSkipped
        End If
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngUsedOrPrivate(ByVal h_wbk As Workbook)
' ----------------------------------------------------------------------------
' Manages the registration of used Common Components, done before change
' components are exported and used changed Common Components are updated.
' When not yet registered a confirmation dialog ensures a component not just
' accidentially has the same name. The type of confirmation is registered
' either as "used" or "private" together with the current revision number.
' When none is available the current date is registered on the fly.
' its Revision-Number is
' ----------------------------------------------------------------------------
    Const PROC = "HskpngUsedOrPrivate"
    
    On Error GoTo eh
    Dim vbc                     As VBComponent
    Dim BttnConfirmedUsed       As String
    Dim BttnConfirmedPrivate    As String
    Dim Msg                     As mMsg.udtMsg
    Dim Comp                    As clsComp
    Dim lRegState               As Long
    
    mBasic.BoP ErrSrc(PROC)
    BttnConfirmedUsed = "Yes!" & vbLf & _
                        "This is a used Common Component" & vbLf & _
                        "identical with the corresponding" & vbLf & _
                        "VBComponent's Export-File in the" & vbLf & _
                        """Common-Components folder"""
    BttnConfirmedPrivate = "No!" & vbLf & _
                           "This is a VBComponent which just" & vbLf & _
                           "by chance has the same name."
    
    For Each vbc In h_wbk.VBProject.VBComponents
        If ExistsPublic(vbc.Name) Then
            Set Comp = New clsComp
            With Comp
                .Wrkbk = h_wbk
                .VBComp = vbc
                '~~ The component is a known Common Component
                lRegState = CompManDat.RegistrationState(.CompName)
                If Not lRegState = enRegStatePrivate _
                And Not lRegState = enRegStateUsed _
                And Not lRegState = enRegStateHosted _
                Then
                    '~~ Once an equally named VBComponent is registered as private it will no longer be regarded as "used" and updated.
                    Msg.Section(1).Text.Text = "The component in the VBProject named   " & mBasic.Spaced(.CompName) & "   is known as a ""Common Component"", " & _
                                               "i.e. an equally named component exists in the ""Common-Components folder""  '" & _
                                               sFolderPath & "', but the component is yet neither registered/known as ""used"" nor as ""private"" !" & vbLf & _
                                               "Just a hint by the way: The component may as well be claimed ""hosted"" by this Workbook in case it is yet not " & _
                                               "claimed ""hosted"" by another Workbook/VBProject. *)"
                        
                    With Msg.Section(2)
                        .Label.Text = "*)"
                        With .Text
                             .Text = "See README, section ""Enabling the services (serviced or not serviced)"""
                             .OnClickAction = "https://github.com/warbe-maker/VBA-Component-Management#enabling-the-services-serviced-or-not-serviced"
                             .FontColor = rgbBlue
                        End With
                    End With
                    
                    Select Case mMsg.Dsply(dsply_title:="Not yet registered ""Common Component""" _
                                         , dsply_msg:=Msg _
                                         , dsply_Label_spec:="R25" _
                                         , dsply_buttons:=mMsg.Buttons(BttnConfirmedUsed, vbLf, BttnConfirmedPrivate))
                        Case BttnConfirmedUsed:     CompManDat.RegistrationState(.CompName) = enRegStateUsed
                                                    CompManDat.LastModAtDateTime(.CompName) = vbNullString ' yet unknown will force update when outdated
                        Case BttnConfirmedPrivate:  CompManDat.RegistrationState(.CompName) = enRegStatePrivate
                    End Select
                    Prgrss.ItemDone = .CompName
                Else
                    Prgrss.ItemSkipped
                End If
            End With
            Set Comp = Nothing
        Else
            Prgrss.ItemSkipped
        End If
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function LastModExpFileExists(ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an Export File with the name extracted from the
' LastModExpFileFullNameOrigin exists in the Common Components Folder
' ----------------------------------------------------------------------------
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(LastModExpFileFullNameOrigin(comp_name))
        LastModExpFileExists = .FileExists(sFolderPath & "\" & FileName)
    End With
End Function

Private Function MaxRawLenght() As Long
' -----------------------------------------------
' Returns the max length of a raw componen's name
' -----------------------------------------------
    Const PROC = "MaxRawLenght"
    
    On Error GoTo eh
    Dim v As Variant
    
    For Each v In Components
        MaxRawLenght = Max(MaxRawLenght, Len(v))
    Next v
    
xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub PendingReleaseRemove(Optional ByVal p_comp_name As String = vbNullString)
' ------------------------------------------------------------------------------
' All pending release information is removed from the component/section
' (r_comp_name) - when not provided from sCompName (error when neither is
' provided.)
' ------------------------------------------------------------------------------
    Const PROC = "PendingReleaseRemove"
    
    On Error GoTo eh
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    If p_comp_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No component name provided, section name is missing!"
    
    With PPCommComps
        mFso.PPremoveNames pp_file:=sCommCompsDatFile _
                          , pp_section:=p_comp_name _
                          , pp_value_names:=.Name(PENDING_RELEASE_MODIFIED_ON_MACHINE) _
                                    & "," & .Name(PENDING_RELEASE_MODIFIED_IN_WBK_NAME) _
                                    & "," & .Name(PENDING_RELEASE_MODIFIED_IN_WBK_FULL_NAME) _
                                    & "," & .Name(PENDING_RELEASE_MODIFIED_LAST_DATE_TIME) _
                                    & "," & .Name(PENDING_RELEASE_MODIFIED_EXP_FILE)
    End With
    Debug.Print "Pending Release removed for " & p_comp_name
    
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ExistsPending(ByVal e_comp_name As String, _
                     Optional ByVal e_on_machine As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' The Computer name is added for not confusing pending release modifications
' made on - and registered for - another computer.
' ------------------------------------------------------------------------------
    If PendingReleaseModInWbkName(e_comp_name) <> vbNullString Then
        If e_on_machine <> vbNullString Then
            If PendingReleaseModOnMachine(e_comp_name) = e_on_machine Then
                ExistsPending = True
            End If
        Else
            ExistsPending = True
        End If
    End If
    
End Function

Public Function PendingReleaseRegistered(Optional ByVal p_comp_name As String = vbNullString) As Boolean
    Const PROC = "PendingReleaseRegistered"
    
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    If p_comp_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No component name provided, section name is missing!"
    
    PendingReleaseRegistered = PendingReleaseModAtDateTime <> vbNullString
    
End Function

Public Function PendingReleases(Optional ByVal p_wbk As Workbook = Nothing) As Collection
' ------------------------------------------------------------------------------
' Returns a Dictionary of clsRelease class instances, each representing a
' component with a pending release whereby the class instance is the key and the
' component name is the item.
' ------------------------------------------------------------------------------
    Const PROC = "PendingReleases"
    
    On Error GoTo eh
    Dim dctComps    As Dictionary
    Dim cllPending  As New Collection
    Dim v           As Variant
    Dim sComp       As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dctComps = CommComps.Components
    For Each v In dctComps
        sComp = v
        If ExistsPending(sComp, Environ("COMPUTERNAME")) Then
            '~~ The Common Component has a pending Release registered
            If p_wbk Is Nothing Then
                '~~ When no specific Workbook is provided all pending on this machine are collected
                cllPending.Add sComp
            ElseIf PendingReleaseModInWbkFullName(sComp) = p_wbk.FullName Then
                '~~ When a Workbook is provided only pending releases concerning this Workbook on this computer are collected
                cllPending.Add sComp
            End If
        End If
    Next v
    Set PendingReleases = cllPending

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Reorg()
    Const PROC = "Reorg"
    mBasic.BoP ErrSrc(PROC)
    PPCommComps.Reorg
    mBasic.EoP ErrSrc(PROC)
End Sub

Public Sub SaveToCommonComponentsFolder(ByVal s_file As Variant)
' ------------------------------------------------------------------------------
' Save the Export-File (s_file) of a pending release to the Common-Components
' folder.
' Note: This is exclusively used by the mRelease.ReleaseComp procedure which
'       subsequently updates the corresponding properties.
' ------------------------------------------------------------------------------
    Const PROC  As String = "SaveToCommonComponentsFolder"
        
    On Error GoTo eh
    Dim fl              As File
    Dim sFileName       As String
    Dim sFileFullName   As String
    
    mBasic.BoP ErrSrc(PROC)
    With fso
        If TypeName(s_file) = "File" _
        Then Set fl = s_file _
        Else Set fl = .GetFile(s_file)
        sFileName = .GetFileName(fl)
        sFileFullName = fl.Path
        
        .CopyFile sFileFullName, wsConfig.FolderCommonComponentsPath & "\" & sFileName
        If .GetExtensionName(fl.Name) = "frm" Then
            '~~ When the Export file has a .frm extension the .frx file needs to be copied too
            .CopyFile Replace(sFileFullName, ".frm", ".frx"), wsConfig.FolderCommonComponentsPath & "\" & .GetBaseName(fl) & ".frx"
        End If
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ReleaseComp(ByVal r_comp_name As String)
' ------------------------------------------------------------------------------
' - Copy a used or hosted Common Component's Export File to the Common
'   Components folder
' - Update all relevant properties
' - Remove any pending release info (assign a vbNullString)
' ------------------------------------------------------------------------------
    Const PROC = "ReleaseComp"
    
    On Error GoTo eh
    sCompName = r_comp_name
    SaveToCommonComponentsFolder PendingReleaseModExpFileFullName
    LastModExpFileFullNameOrigin = PendingReleaseModExpFileFullName
    LastModInWbkName = PendingReleaseModInWbkName
    LastModInWbkFullName = PendingReleaseModInWbkFullName
    LastModAtDateTime = PendingReleaseModAtDateTime
    '~~ Remove any pending release info for the Common Component in the CommComps.dat file (Common-Components folder)
    PendingReleaseRemove r_comp_name

xt: Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
    
End Sub

Public Sub PendingReleaseManagement(ByVal p_comp As clsComp)
' ------------------------------------------------------------------------------
' Remove obsolete 'Pending Releases' and establish/register 'Pending Releases'
' or manage a concurrent modification.
'
' ------------------------------------------------------------------------------
    Dim bModified       As Boolean
    Dim bNotModified    As Boolean
    Dim bPending        As Boolean
    Dim bNotPending     As Boolean
    Dim bDiffWbk        As Boolean
    Dim bEqulWbk        As Boolean
    Dim bDiffMachine    As Boolean
    Dim bEqulMachine
    
    With p_comp
        CompName = .CompName
        If CurrentPublicExpFileFullName <> vbNullString Then
            .CodePublic.Source = CurrentPublicExpFileFullName
            bModified = .CodeExported.DiffersFrom(p_comp.CodePublic, True)
        Else
            bModified = True
        End If
            bNotModified = Not bModified
        bPending = CommComps.PendingReleaseRegistered
        bNotPending = Not bPending

        If bPending Then
            CodePending.Source = PendingReleaseModExpFileFullName
            bDiffWbk = LastModInWbkFullName <> vbNullString And LastModInWbkFullName <> "unknown" And .WrkbkFullName <> PendingReleaseModInWbkFullName
            bEqulWbk = Not bDiffWbk
            bDiffMachine = PendingReleaseModOnMachine <> vbNullString _
                       And PendingReleaseModOnMachine <> Environ("COMPUTERNAME")
            bEqulMachine = Not bDiffMachine
        End If
    End With
        
    ' Cases: -------------1-2-3-4-5-6--------------
    ' C1 Code modified *) y y y y y n
    ' C2 Pending          n y y y y y
    ' C3 Diff Wbk.        - n y n y n
    ' C4 Diff Machine.    - n n y y n
    ' ---------------------------------------
    ' A1 Reg. pend.       x
    ' A2 Reg. mod inf.      x
    ' A3 Manage conc.         x x x
    ' A4 Remove pending.            x
    ' ----------------------------------------
    ' *) The exported code differs from the public
    '    code in the Common-Components folder
    Select Case True
        Case bModified And bNotPending                                  ' Case 1        > A1
            PendingReleaseRegister = p_comp
        Case bModified And bPending And bEqulWbk And bEqulMachine       ' Case 2        > A2
            If PendingReleaseModAtDateTime <> p_comp.LastModAtDateTime Then
                PendingReleaseModAtDateTime = p_comp.LastModAtDateTime
            End If
            Debug.Print "Pending Release mod at updated: " & sCompName
        Case bModified And bPending And (bDiffWbk Or bDiffMachine)      ' Cases 3,4,5   > A3
            ConcurrentModification p_comp
        Case bNotModified And bPending And bEqulWbk And bEqulMachine    ' Case 6        > A4
            PendingReleaseRemove
    End Select
    
End Sub


