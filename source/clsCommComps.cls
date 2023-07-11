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
'                             b) in the CommComps.dat
' Public Methods:
' ----------------
' BttnInconsistencyExport           .
' BttnInconsistencySkip             .
' BttnInconsistencyUpdate           .
' CommCompsDatFileFullName          .
' LastModExpFile               r/w  .
' LastModExpFileFullName       r    .
' LastModExpFileFullNameOrigin r/w  .
' LastModWbk                     w  .
' LastModWbkBaseName           r    .
' LastModWbkFullName           r/w  .
' LastModWbkName               r    .
' RevisionNumber               r/w  .
' ---------------------------------------------------------------------------
Public Qoutdated                                As clsQ

' Housekeeping syntax: (allows the on-the-fly-change of value names as well as their removal)
' "<current-name>:<old-name>" = rename in all sections
' "<current>"                 = no action
' ":<remove-name>"            = remove in all sections
Private Const LAST_MOD_EXP_FILE_FULL_NAME           As String = "LAST_MOD_EXP_FILE_FULL_NAME"
Private Const LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN    As String = "LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN"
Private Const LAST_MOD_WBK_FULL_NAME                As String = "LAST_MOD_WBK_FULL_NAME"
Private Const REVISION_NUMBER                       As String = "REVISION_NUMBER"

Private BttnUpdate                              As String
Private BttnDsplyDiffs                          As String
Private BttnSkipForNow                          As String
Private BttnSkipForever                         As String
Private UpdateDialogTitle                       As String
Private UpdateDialogTop                         As Long
Private UpdateDialogLeft                        As Long
Private PP                                      As clsPP
Private fso                                     As FileSystemObject

Private Sub Class_Initialize()
    Dim dct As New Dictionary
    
    Set fso = New FileSystemObject
    Set PP = New clsPP
    
    '~~> Full name of the PP-File
    PP.FullName = wsConfig.FolderCommonComponentsPath & "\CommComps.dat"
        
    '~~> Value Names with the syntax (allows the on-the-fly-change of value names as well as their removal):
    '~~> "<current>"                 = no action
    '~~> "<current-name>:<old-name>" = rename in all sections
    '~~> ":<remove-name>"            = remove in all sections
        dct.Add LAST_MOD_EXP_FILE_FULL_NAME, "LastModExpFileFullName"
        dct.Add LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN, "LastModExpFileFullNameOrigin"
        dct.Add LAST_MOD_WBK_FULL_NAME, "LastModWbkFullName"
        dct.Add REVISION_NUMBER, "RevisionNumber"
    PP.Names = dct
    Set dct = Nothing
    
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
    Set PP = Nothing
End Sub

Friend Property Get BttnInconsistencyExport() As String
    BttnInconsistencyExport = "Export" & vbLf & _
                              "(the hosted version is the one up-to-date)"
End Property

Friend Property Get BttnInconsistencySkip() As String
    BttnInconsistencySkip = "Skip" & vbLf & "for further investigation"
End Property

Friend Property Get BttnInconsistencyUpdate() As String
    BttnInconsistencyUpdate = "Update (re-import)" & vbLf & _
                              "(the ""Common-Components Folder"" version is up-to-date)"

End Property

Friend Property Get LastModExpFile(Optional ByVal comp_name) As File
    Const PROC = "LastModExpFile Get"
    
    On Error GoTo eh
    Dim FileName    As String
    Dim sPath       As String
    
    sPath = wsConfig.FolderCommonComponentsPath & "\"
    With New FileSystemObject
        FileName = .GetFileName(LastModExpFileFullNameOrigin(comp_name))
        If FileName <> vbNullString Then
            If .FileExists(sPath & FileName) Then
                Set LastModExpFile = .GetFile(sPath & FileName)
            Else
                Set LastModExpFile = .CreateTextFile(sPath & FileName)
            End If
        End If
    End With
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let LastModExpFile(Optional ByVal comp_name, _
                                           ByVal comp_exp_file As File)
' ---------------------------------------------------------------------------
' Copies the Raw Export File from its host Workbook location to the Common
' Components Folder from where it is used as the source for the update of
' outdated Used Common Components.
' ---------------------------------------------------------------------------
    comp_name = comp_name ' not used
    comp_exp_file.Copy Destination:=wsConfig.FolderCommonComponentsPath & "\" & comp_exp_file.Name, OverWriteFiles:=True
End Property

Friend Property Get LastModExpFileFullName(Optional ByVal comp_name As String) As String
' ---------------------------------------------------------------------------
' Returns the Export File Full Name which refers to the Export File saved in
' the Common-Components folder.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        LastModExpFileFullName = wsConfig.FolderCommonComponentsPath & "\" & .GetFileName(LastModExpFileFullNameOrigin(comp_name))
    End With
End Property

Friend Property Get LastModExpFileFullNameOrigin(Optional ByVal comp_name As String) As String
    With PP
        LastModExpFileFullNameOrigin = .Value(pp_section:=comp_name, pp_value_name:=.Name(LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN))
    End With
End Property

Friend Property Let LastModExpFileFullNameOrigin(Optional ByVal comp_name As String, _
                                                          ByVal r_exp_file_full_name As String)
    With PP
        .Value(pp_section:=comp_name, pp_value_name:=.Name(LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN)) = r_exp_file_full_name
    End With
End Property

Friend Property Let LastModWbk(Optional ByVal r_comp_name As String, _
                                        ByVal r_wbk As Workbook)
    With PP
        If r_wbk Is Nothing _
        Then .Value(pp_section:=r_comp_name, pp_value_name:=.Name(LAST_MOD_WBK_FULL_NAME)) = "unknown" _
        Else .Value(pp_section:=r_comp_name, pp_value_name:=.Name(LAST_MOD_WBK_FULL_NAME)) = r_wbk.FullName
    End With
    
End Property

Friend Property Get LastModWbkBaseName(Optional ByVal comp_name As String) As String
    With PP
        LastModWbkBaseName = fso.GetBaseName(.Value(pp_section:=comp_name, pp_value_name:=.Name(LAST_MOD_WBK_FULL_NAME)))
    End With
End Property

Friend Property Get LastModWbkFullName(Optional ByVal comp_name As String) As String
    With PP
        LastModWbkFullName = .Value(pp_section:=comp_name, pp_value_name:=.Name(LAST_MOD_WBK_FULL_NAME))
    End With
End Property

Friend Property Get LastModWbkName(Optional ByVal comp_name As String) As String
    With PP
        LastModWbkName = fso.GetFileName(.Value(pp_section:=comp_name, pp_value_name:=.Name(LAST_MOD_WBK_FULL_NAME)))
    End With
End Property

Friend Property Get RevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    With PP
        RevisionNumber = .Value(pp_section:=comp_name, pp_value_name:=.Name(REVISION_NUMBER))
    End With
End Property

Friend Property Let RevisionNumber(Optional ByVal comp_name As String, _
                                            ByVal comp_rev_no As String)
' ------------------------------------------------------------------------------
' Returns a revision number in the form yy-mm-dd.00. Plus one when an existing
' revision number is provided (comp_rev-no) or the current date with .01.
' ------------------------------------------------------------------------------
    Const PROC = "RevisionNumber Let"
    
    On Error GoTo eh
    
    With PP
        If comp_rev_no = vbNullString Then
            .Value(pp_section:=comp_name, pp_value_name:=.Name(REVISION_NUMBER)) = CompManDat.RevisionNumberInitial
        Else
            .Value(pp_section:=comp_name, pp_value_name:=.Name(REVISION_NUMBER)) = comp_rev_no
        End If
    End With
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function CommCompRegStateEnum(ByVal s As String) As enCommCompRegState
    Select Case s
        Case "hosted":  CommCompRegStateEnum = enRegStateHosted
        Case "used":    CommCompRegStateEnum = mComp.enRegStateUsed
        Case "private": CommCompRegStateEnum = mComp.enRegStatePrivate
    End Select
End Function

Public Function CommCompRegStateString(ByVal en As enCommCompRegState) As String
    Select Case en
        Case enRegStateHosted:  CommCompRegStateString = "hosted"
        Case enRegStateUsed:    CommCompRegStateString = "used"
        Case enRegStatePrivate: CommCompRegStateString = "private"
    End Select
End Function

Public Sub ComponentRemove(ByVal s As String)
     PP.ComponentRemove s
End Sub

Public Function Components() As Dictionary
    Set Components = PP.Components
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommComps" & "." & sProc
End Function

Public Function Exists(ByVal x_vbc As VBComponent, _
                       ByVal x_exp_file_ext As String, _
              Optional ByRef x_modified_by_wbk_name As String, _
              Optional ByRef x_export_file_full_name As String, _
              Optional ByRef x_export_file As File, _
              Optional ByRef x_last_mod_rev_no As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the VBComponent's (x_vbc) Export-File exists in the
' global Common-Components-Folder.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Exists"
    
    Dim sFile   As String
    Dim Comp    As New clsComp
    
    sFile = wsConfig.FolderCommonComponentsPath & "\" & x_vbc.Name & x_exp_file_ext
    If fso.FileExists(sFile) Then
        Exists = True
        x_export_file_full_name = sFile
        Set x_export_file = CommComps.LastModExpFile(x_vbc.Name)
        x_modified_by_wbk_name = CommComps.LastModWbkName(x_vbc.Name)
        x_last_mod_rev_no = CommComps.RevisionNumber(x_vbc.Name)
    End If
    
xt: Set Comp = Nothing
    
End Function

Public Function ExistsRegistered(ByVal raw_vbc_name As String) As Boolean
    ExistsRegistered = Components.Exists(raw_vbc_name)
End Function

Public Function Hosted(ByVal h_hosted As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a string of component names (h_hosted) as Dictionary.
' ----------------------------------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    Dim dct     As New Dictionary
    
    For Each v In Split(h_hosted, ",")
        sComp = Trim$(v)
        If Not dct.Exists(sComp) Then
            dct.Add sComp, sComp
        End If
    Next v
    Set Hosted = dct
    Set dct = Nothing
    
End Function

Private Sub HskpngObsolete(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Remove in the PrivateProfile file CommComps.dat:
' - Sections representing VBComponents for which an Export-File does not exist
'   in the Common-Components folder
' - Sections indicating a Common Component of the serviced Workbook but the
'   component not exists.
' ------------------------------------------------------------------------------
    Const PROC = "HskpngObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wbk         As Workbook
    Dim dct         As Dictionary
    Dim sBaseName   As String
    Dim dctHosted   As Dictionary
    Dim sExpFile    As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dctHosted = CommComps.Hosted(h_hosted)
    Set wbk = Services.Serviced
    sBaseName = fso.GetBaseName(wbk.FullName)
    Set dct = CommComps.Components
    
    '~~ Obsolete because the component is no longer hosted by the indicated Workbook
    '~~ no longer exist in the indicated Workbook
    For Each v In dct
        If CommComps.LastModWbkBaseName(v) = sBaseName Then
            '~~ The component indicates being one of the serviced Workbook
            If Not mComp.Exists(v, wbk) Then
                CompManDat.ComponentRemove v
            End If
        End If
    Next v
    
    '~~ Obsolete because the corresponding Export-File
    '~~ no longer exists in the Common-Components folder
    '~~ De-register global Common Components no longer hosted
    Set dct = CommComps.Components
    For Each v In dct
        sExpFile = fso.GetFileName(CommComps.LastModExpFileFullNameOrigin(v))
        If sExpFile <> vbNullString Then
            If Not fso.FileExists(wsConfig.FolderCommonComponentsPath & "\" & sExpFile) Then
                CommComps.ComponentRemove v
            End If
        End If
    Next v
    Set dct = CommComps.Components
    
xt: Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngUsed()
' ----------------------------------------------------------------------------
' Manages the registration of used Common Components, done before change
' components are exported and used changed Common Components are updated.
' When not yet registered a confirmation dialog ensures a component not just
' accidentially has the same name. The type of confirmation is registered
' either as "used" or "private" together with the current revision number.
' When none is available the current date is registered on the fly.
' its Revision-Number is
' ----------------------------------------------------------------------------
    Const PROC          As String = "HskpngUsed"
    
    On Error GoTo eh
    Dim vbc             As VBComponent
    Dim wbk             As Workbook
    Dim BttnConfirmed   As String
    Dim BttnPrivate     As String
    Dim Msg             As mMsg.TypeMsg
    Dim Comp            As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    BttnConfirmed = "Yes!" & vbLf & _
                    "This is a used Common Component" & vbLf & _
                    "identical with the corresponding" & vbLf & _
                    "VBComponent's Export-File in the" & vbLf & _
                    """Common-Components folder"""
    BttnPrivate = "No!" & vbLf & _
                  "This is a VBComponent which" & vbLf & _
                  "accidentially has the same name."
    
    Set wbk = Services.Serviced
    For Each vbc In wbk.VBProject.VBComponents
        Set Comp = New clsComp
        With Comp
            .Wrkbk = wbk
            .VBComp = vbc
            If CommComps.Exists(vbc, .ExpFileExt) Then
                If CommComps.RevisionNumber(vbc.Name) = vbNullString Then
                    CommComps.RevisionNumber(vbc.Name) = vbNullString
                End If
                If Not CompManDat.RegistrationState(vbc.Name) = enRegStatePrivate _
                And Not CompManDat.RegistrationState(vbc.Name) = enRegStateUsed _
                And Not CompManDat.RegistrationState(vbc.Name) = enRegStateHosted _
                Then
                    '~~ Once an equally named VBComponent is registered a private it will no longer be regarded as "used" and updated.
                    Msg.Section(1).Text.Text = "The VBComponent named   " & mBasic.Spaced(vbc.Name) & "   is known as a ""Common Component"" " & _
                                               "because it exists in the ""Common-Components folder""  '" & _
                                               wsConfig.FolderCommonComponentsPath & "'  but is yet not registered either " & _
                                               "as used or private in the serviced Workbook."
                    
                    Select Case mMsg.Dsply(dsply_title:="Not yet registered ""Common Component""" _
                                         , dsply_msg:=Msg _
                                         , dsply_buttons:=mMsg.Buttons(BttnConfirmed, vbLf, BttnPrivate))
                        Case BttnConfirmed: CompManDat.RegistrationState(vbc.Name) = enRegStateUsed
                                            CompManDat.RevisionNumber(vbc.Name) = vbNullString ' yet unknown will force update when outdated
                        Case BttnPrivate:   CompManDat.RegistrationState(vbc.Name) = enRegStatePrivate
                    End Select
                End If
            Else
                '~~ The Export-File has manually been copied into the Common
                '~~ Components-Folder and thus is yet not registered
                
            End If
        End With
        Set Comp = Nothing
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Hskpng(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    
    HskpngProperties
    HskpngObsolete h_hosted
    HskpngHosted h_hosted
    HskpngNotHosted h_hosted
    HskpngUsed
    Reorg
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngHosted(ByVal m_hosted As String)
' ----------------------------------------------------------------------------
' - Registers the Workbook as 'Raw-Host' when it hosts at least one Common
'   Component
' - Maintains an up-to-date copy of the Export-File in the Common-Components
'   folder
' - Maintains for each hosted (raw) Common Component the properties:
'   - in the local CommComps.dat:
'     - Component Name
'     - Revision Number
'   - in the ComComps-RawsSaved.dat in the Common-Components folder:
'     - Component Name
'     - Export File Full Name
'     - Host Base Name
'     - Host Full Name
'     - Host Name
'     - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "HskpngHosted"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim CommCompHosted  As clsComp
    Dim sHostBaseName   As String
    Dim dctHosted       As Dictionary
    Dim wbk             As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    sHostBaseName = fso.GetBaseName(wbk.FullName)
    Set dctHosted = CommComps.Hosted(m_hosted)
                    
    For Each v In dctHosted
        If Not mComp.Exists(v, wbk) Then
            MsgBox "The VBComponent " & v & " is claimed hosted by the serviced Workbook " & Services.Serviced.Name & _
                   " will be ignored (it does not exist in the Workbook's VB-Project) !" & vbLf & vbLf & _
                   "When the component is no longer hosted or its name has changed the argument needs to be updated accordingly.", _
                   vbOK, "VBComponent " & v & "does not exist!"
        Else
            Set CommCompHosted = New clsComp
            With CommCompHosted
                .Wrkbk = wbk
                .CompName = v
                If CompManDat.RegistrationState(v) <> enRegStateHosted Then
                    '~~ The Workbook has yet not claimed the Common Component hosted but now does.
                    CompManDat.RegistrationState(v) = enRegStateHosted
                    '~~ This housekeeping is executed prior the "Export of changed components"
                    '~~ The Export-File comparison is therefore done with a temporary Export-File.
                    If Not Services.FilesDiffer(f_file_1:=.ExpFileTemp _
                                              , f_file_2:=CommComps.LastModExpFile(v) _
                                              , f_ignore_export_header:=True) _
                    Then
                        '~~ Only when the claiming Workbook's Common Component is identical
                        '~~ with the Export-File in the Common-Components folder it is also
                        '~~ registered as the raw source
                        CommComps.LastModWbk(v) = wbk
                        CommComps.LastModExpFileFullNameOrigin(v) = .ExpFileFullName
                        CommComps.RevisionNumber = .RevisionNumber
                    Else
                        '~~ Any other Workbook appears to have modified and saved the Common Component
                        '~~ to the Common-Components folde. This information remains valid.
                    End If
                End If
                
                If Not CommComps.Exists(.VBComp, .ExpFileExt) Then
                    CommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                    
                End If
                
                If Not Services.FilesDiffer(f_file_1:=.ExpFile _
                                          , f_file_2:=CommComps.LastModExpFile(v)) Then
                    If CommComps.RevisionNumber(v) <> CompManDat.RevisionNumber(v) Then
                        CommComps.RevisionNumber(v) = CompManDat.RevisionNumber(v)
                    End If

                End If
            End With
            Set CommCompHosted = Nothing
        End If
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngNotHosted(ByVal h_hosted As String)
' ----------------------------------------------------------------------------
' When a former hosting Workbook not or no longer claims a Common Component
' hosted the RegistrationState is changed to enRegStateUsed.
' ----------------------------------------------------------------------------
    Const PROC      As String = "HskpngNotHosted"
    
    On Error GoTo eh
    Dim dctHosted   As Dictionary
    Dim dctComps    As Dictionary
    Dim v           As Variant
    Dim wbk         As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    Set dctHosted = CommComps.Hosted(h_hosted)
    Set dctComps = CommComps.Components
    For Each v In dctComps
        If CommComps.LastModWbkName(v) = wbk.Name Then
            If Not dctHosted.Exists(v) Then
                If mComp.Exists(v, wbk) Then
                    CompManDat.RegistrationState(v) = enRegStateUsed
                End If
            End If
        End If
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngProperties()
' ------------------------------------------------------------------------------
' Adds for each Common Component's Export-File in the Common-Components folder
' a section to the CommComps.dat when missing or updates the ExportFileName when
' not identical with the found file.
'
' Background:
' A missing section indicates a Common Component of wich the Export-File has
' obvously been copied manually into the Common-Components folder which now in
' the sense of CompMan has become an available Common Component ready for being
' imported into any VB-Project. A new registered Common Component remains
' un-hosted until a Workbbok claims hosting it, i.e. providing a delevelopment
' and test environment for it.
' ------------------------------------------------------------------------------
    Const PROC = "HskpngProperties"
    
    On Error GoTo eh
    Dim fle         As File
    Dim dct         As Dictionary
    Dim sCompName   As String
    Dim sExt        As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dct = CommComps.Components
    With fso
        For Each fle In .GetFolder(wsConfig.FolderCommonComponentsPath).Files
            sExt = .GetExtensionName(fle.Path)
            Select Case sExt
                Case "bas", "frm", "cls"
                    sCompName = .GetBaseName(fle.Path)
                    If Not dct.Exists(sCompName) Then
                        '~~ The Export-File is yet not registered as a known Common Component
                        '~~ It most likely has been copied manually into the Common-Components
                        '~~ folder. I.e. its "raw host" is unknown - and registered as such.
                        '~- The raw host will remain unknown until the Common Component is
                        '~~ modified in a Workbook using it and exported.
                        CommComps.LastModWbk(sCompName) = Nothing
                        CommComps.RevisionNumber(sCompName) = CompManDat.RevisionNumberInitial
                    Else
                        If CommComps.LastModExpFileFullNameOrigin(sCompName) = vbNullString Then
                            Debug.Print "The property ""LastModExpFileFullNameOrigin"" of component " & sCompName & " is not available, i.e. its origin is unknown or simply yet not registered respectively!"
                        End If
                    End If
            End Select
        Next fle
    End With
    
xt: Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function InconsistencyWarning(ByVal i_file_full_name, _
                                     ByVal i_file_full_name_saved, _
                                     ByVal i_message) As Variant
' ----------------------------------------------------------------------------
' Displays an information about a modification of a Used Common Component.
' The display offers the option to display the code difference. The function
' returns the finally pressed button.
' ----------------------------------------------------------------------------
    Const PROC = "InconsistencyWarning"
    
    On Error GoTo eh
    Dim Msg                 As mMsg.TypeMsg
    Dim cllBttns            As Collection
    Dim BttnDsply           As String
    Dim BttnSkip            As String
    Dim BttnExport          As String
    Dim BttnUpdate          As String
    
    BttnDsply = "Display difference" & "(of Export-Files)"
    BttnSkip = BttnInconsistencySkip
    BttnExport = BttnInconsistencyExport
    BttnUpdate = BttnInconsistencyUpdate
    
    Set cllBttns = mMsg.Buttons(BttnDsply, vbLf, BttnExport, BttnUpdate, vbLf, BttnSkip)
    With Msg.Section(1)
        With .Label
            .Text = "Attention!"
            .FontColor = rgbRed
        End With
        With .Text
            .Text = i_message
        End With
    End With
    With Msg.Section(2)
        .Label.Text = Replace(Replace(BttnExport, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "According to the displayed differencies is the hosted version the one up-to-date. " & vbLf & _
                     "It will be exported and copied to the ""Common-Components Folder"" and the 'Revision Number' will be increased."
    End With
    With Msg.Section(3)
        .Label.Text = Replace(Replace(BttnUpdate, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "The hosted version will be updated with the ""Common-Components Folder"" version (by re-import) and " & _
                     "the 'Revision-Number' of the hosted version will be set identical with the ""Common-Components Folder"" version"
    End With
    With Msg.Section(4)
        .Label.Text = Replace(Replace(BttnSkip, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "Will be clarified later!& " & vbLf & _
                     "Note: Each Workbook Safe will redisplay this message until either Export or Update is performed."
    End With
    With Msg.Section(5)
        .Label.Text = "Background:"
        .Label.FontColor = rgbBlue
        .Text.Text = "When 'hosted Common Component' is modified within its hosting Workbook and exported, the " & _
                     "'Export File' is copied to the ""Common-Components Folder"" and the 'Revision Number' is increased " & _
                     "and set equal in the hosting Workbook's and the Common-Component Folder's ""CommComps.dat"" file. " & _
                     "When a 'used Common Component is modified within the VB-Project just using (not hosting!) it, " & _
                     "the 'Revision Number' only of this 'used Common Component' is increased. When by accident both " & _
                     "had been modified the 'Revision Numbers' may be equal but the Export Files will differ."
    End With
        
    Do
        Select Case mMsg.Dsply(dsply_title:="Inconsistency warning for " & fso.GetBaseName(i_file_full_name) & "!" _
                             , dsply_msg:=Msg _
                             , dsply_buttons:=cllBttns _
                              )
            Case BttnDsply
                Services.ExpFilesDiffDisplay e_file_left_full_name:=i_file_full_name _
                                           , e_file_left_title:="Raw Common Component's Export File: (" & i_file_full_name & ")" _
                                           , e_file_right_full_name:=i_file_full_name_saved _
                                           , e_file_right_title:="Saved Raw's Export File (" & i_file_full_name_saved & ")"
            Case BttnSkip:      InconsistencyWarning = BttnSkip:     Exit Do
            Case BttnExport:    InconsistencyWarning = BttnExport:   Exit Do
            Case BttnUpdate:    InconsistencyWarning = BttnUpdate:   Exit Do
        End Select
    Loop

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function LastModExpFileExists(ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an Export File with the name extracted from the
' LastModExpFileFullNameOrigin exists in the Common Components Folder
' ----------------------------------------------------------------------------
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(LastModExpFileFullNameOrigin(comp_name))
        LastModExpFileExists = .FileExists(wsConfig.FolderCommonComponentsPath & "\" & FileName)
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub OutdatedUpdate()
' ------------------------------------------------------------------------------
' Collects all used outdated Common Components when called for the first time
' and displays the first one queued in Qoutdated. The service is re-called until
' the queue is empty. The display of the update choices is a mode-less dialog
' which calls sub-services in accordance with the button pressed.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate"
    
    On Error GoTo eh
    If Qoutdated Is Nothing Then OutdatedUpdateCollect
    If Not Qoutdated.IsEmpty Then
        OutdatedUpdateChoice
    Else
        Services.DsplyProgress "used Common Components updated"
        Services.LogEntrySummary Application.StatusBar
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoice()
' ------------------------------------------------------------------------------
' Displays the first outdated Common Component in the queue Qoutdated in a mode-
' less dialog for one of the options: "update", "display diffs", "skip for now",
' and "skip forever".
' The service considers used and hosted Common Components to be updated by
' dedicated buttons and section texts.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoice"
    
    On Error GoTo eh
    Dim AppRunArgs      As Dictionary
    Dim cllButtons      As Collection
    Dim Comp            As clsComp
    Dim fUpdate         As fMsg
    Dim Msg             As mMsg.TypeMsg
    Dim sUpdate         As String
    Dim sSkipForever    As String
    Dim sSkipNow        As String
    Dim sModWbkName     As String
    Dim sUpdateBttnTxt  As String
    Dim sSkipNowNote    As String
    
    Set AppRunArgs = New Dictionary
    Qoutdated.First Comp ' Get the next outdated component from the queue
    
    With Comp
        sModWbkName = CommComps.LastModWbkName(.CompName)
        BttnDsplyDiffs = "Display the code modifications"
        Select Case .KindOfComp
            Case enCommCompHosted
                UpdateDialogTitle = "Hosted ""Common Component"" apparently modified within a Worbook using it!"
                sUpdate = "The ""Common Component""   " & mBasic.Spaced(.CompName) & "   hosted in this " & _
                          "Workbook has been modified within the Workbook/VB-Project   " & _
                          mBasic.Spaced(sModWbkName) & " .   It needs to be updated in this hosting Workbook " & _
                          "for consistency! The extent of the coded modification can be seen with """ & _
                          BttnDsplyDiffs & """ wich should allow an estimation wether or to which extent " & _
                          "re-testing is appropriate."
                sSkipNow = "The update will be postponed and proposed with the next Workbook open."
                sSkipNowNote = "Please note: In contrast to a ""used Common Component"" it is not possible " & _
                               "to skip this update forever. The only way to achieve this is to remove the " & _
                               "hosted indication for this Workbook. With the next Workbook open the ""Common " & _
                               "Component"" will be de-registered as hosted and registered as a used one. As a " & _
                               "consequence it will remain ""not hosted"" until another Workbooks claims " & _
                               "hosting it (or forever in case none ever does)."
                BttnUpdate = "Update the hosted" & vbLf & "Common Component" & vbLf & vbLf & .CompName
                sUpdateBttnTxt = "With this update the hosting Workbook again becomes the Workbook hosting the raw " & _
                                 "version of the Common Component."
                BttnSkipForNow = "Skip this update for now"
                Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow)
                Set fUpdate = mMsg.MsgInstance(UpdateDialogTitle)
                mMsg.BttnAppRun AppRunArgs, BttnUpdate _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceUpdate" _
                                            , Comp.CompName
                mMsg.BttnAppRun AppRunArgs, BttnDsplyDiffs _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceDsplyDiffs"
                mMsg.BttnAppRun AppRunArgs, BttnSkipForNow _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceSkipForNow" _
                                            , Comp.CompName

            Case enCommCompUsed
                UpdateDialogTitle = "A used ""Common Component"" is outdated, i.e. has been modified."
                sUpdate = "The used ""Common Component""   " & mBasic.Spaced(.CompName) & "   has been modified within " & _
                          "the Workbook/VB-Project (" & "" & "). Updating the used version is thus recommended!"
                sSkipNow = "Update will be postponed and proposed with the next Workbook open"
                sSkipForever = "The component, known as a potential 'Used Common Component' will be de-registered " & _
                               "and ignored in the future! I.e. the ""used"" status will be changed into ""private"" status. " & _
                               "Re-instantiating as a ""used Common Component"" will requires the following steps:" & vbLf & _
                               "1. Remove it" & vbLf & _
                               "2. Save the Workbook" & vbLf & _
                               "3. Re-Import it from the ""Common-Components"" folder."
                BttnUpdate = "Update the used" & vbLf & "Common Component" & vbLf & vbLf & .CompName
                sUpdateBttnTxt = "The outdated ""Common Component"" used in this Workbook becomes up-to-date again."
                BttnSkipForNow = "Skip this update" & vbLf & "for now"
                BttnSkipForever = "Skip this update" & vbLf & "f o r e v e r" & vbLf & "(I am aware of the consequence)"
                Set fUpdate = mMsg.MsgInstance(UpdateDialogTitle)
                Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow, BttnSkipForever)
                mMsg.BttnAppRun AppRunArgs, BttnUpdate _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceUpdate" _
                                            , Comp.CompName
                mMsg.BttnAppRun AppRunArgs, BttnDsplyDiffs _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceDsplyDiffs"
                mMsg.BttnAppRun AppRunArgs, BttnSkipForNow _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceSkipForNow" _
                                            , Comp.CompName
                mMsg.BttnAppRun AppRunArgs, BttnSkipForever _
                                            , ThisWorkbook _
                                            , "CommComps.OutdatedUpdateChoiceSkipForever" _
                                            , Comp.CompName
            Case Else
                Stop
        End Select
        
    End With
    
    With Msg
        .Section(1).Text.Text = sUpdate
        With .Section(2)
            With .Label
                .Text = Replace(BttnUpdate, vbLf, " ") & ":"
                .FontColor = rgbBlue
            End With
            .Text.Text = sUpdateBttnTxt
        End With
        With .Section(3)
            With .Label
                .Text = Replace(BttnDsplyDiffs, vbLf, " ") & ":"
                .FontColor = rgbBlue
            End With
            .Text.Text = "The displayed code modifications (return with Esc) may help estimating the extent to which re-testing is appropriate after an update."
        End With
        
        With .Section(4)
            With .Label
                .Text = Replace(BttnSkipForNow, vbLf, " ") & ":"
                .FontColor = rgbBlue
            End With
            .Text.Text = sSkipNow
        End With
        If Comp.KindOfComp = enCommCompHosted Then
            With .Section(5).Text
                .Text = sSkipNowNote
                .FontBold = True
            End With
        Else
            '~~ Skip forever is only an option for Workbooks using the Common Component
            With .Section(6)
                With .Label
                    .Text = Replace(BttnSkipForever, vbLf, " ")
                    .FontColor = rgbBlue
                End With
                .Text.Text = sSkipForever
            End With
        End If
        With .Section(7)
            With .Label
                .Text = "About:"
                .OpenWhenClicked = GITHUB_REPO_URL & "#common-components"
                .FontColor = rgbBlue
                .FontUnderline = True
            End With
            .Text.Text = "A ""Common Component"" is one of which an Export-File resides in the ""Common-Components folder"". " & _
                         "It may be modified within whichever Workbook's VB-Project using it, preferrably in a Workbook " & _
                         "which claims hosting it. When a ""Common Component"" is modified its ""Revision Number"" is " & _
                         "increased and the Export-File is copied/overwritten in the ""Common-Components Folder""."
        End With
    End With
    
    '~~ Display the mode-less dialog for the Names synchronization to run
    mMsg.Dsply dsply_title:=UpdateDialogTitle _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=UpdateDialogTop & ";" & UpdateDialogLeft
    DoEvents
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceDsplyDiffs()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceDsplyDiffs"
    
    On Error GoTo eh
    Dim Comp    As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    With Comp
        Services.ExpFilesDiffDisplay .ExpFileFullName, CommComps.LastModExpFile(.CompName), "Currently used (" & .ExpFileFullName & ")", "Up-to-date (" & CommComps.LastModExpFile(.CompName).Path & ")"
    End With
    Set Comp = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceSkipForever(ByVal u_comp_name)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceSkipForever"
    
    On Error GoTo eh
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    CompManDat.RegistrationState(u_comp_name) = enRegStatePrivate
    Qoutdated.DeQueue
    Set wbk = Services.Serviced
    With New clsComp
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
        .LogServicedEntry "Outdated used Commpon Component: Update skipped forever!"
    End With
    
xt: Services.MessageUnload UpdateDialogTitle
    CommComps.OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceSkipForNow(ByVal u_comp_name As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceSkipForNow"
    
    On Error GoTo eh
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.DeQueue
    Set wbk = Services.Serviced
    With New clsComp
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
        .LogServicedEntry "Outdated used Commpon Component: Update skipped for now!"
    End With
    
xt: Services.MessageUnload UpdateDialogTitle
    CommComps.OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceUpdate(ByVal u_comp_name As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceUpdate"
    
    On Error GoTo eh
    Dim wbk     As Workbook
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim sFile   As String
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    Set wbk = Services.Serviced
    Set Comp = New clsComp
    With Comp
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
        v = .KindOfComp
        sFile = CommComps.LastModExpFileFullName(.CompName)
        mUpdate.ByReImport b_wbk_target:=wbk _
                         , b_vbc_name:=u_comp_name _
                         , b_exp_file:=sFile
        
        With Services
            .NoOfItemsServiced = .NoOfItemsServiced + 1
            .NoOfItemsServicedNames = u_comp_name
            .DsplyProgress "used Common Components updated"
        End With
    
        Select Case .KindOfComp
            Case enCommCompHosted
                Services.LogServicedEntry "Outdated Common Component hosted updated by re-import of the Export-File in the Common-Components folder"
                '~~ When a hosted Common Component is updated it "again" becomes the raw host!"
                CommComps.LastModWbk(.CompName) = .Wrkbk
                CommComps.LastModExpFileFullNameOrigin(.CompName) = .ExpFileFullName
            Case enCommCompUsed
                Services.LogServicedEntry "Outdated Common Component used updated by re-import of the Export-File in the Common-Components folder"
        End Select
    
    End With
    Qoutdated.DeQueue
    Set Comp = Nothing
    
xt: Services.MessageUnload UpdateDialogTitle
    CommComps.OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateCollect()
' ------------------------------------------------------------------------------
' Collects all outdated Used Common Components and enqueues them in Qoutdated.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateCollect"
    
    On Error GoTo eh
    Dim Comp        As clsComp
    Dim wbk         As Workbook
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim sName       As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    Set Qoutdated = New clsQ
    Set dct = Comps.All ' all = all relevant for the current service
    
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            Services.ServicedItem = .VBComp
            If (.KindOfComp = enCommCompHosted Or .KindOfComp = enCommCompUsed) Then
                If .Outdated Then
                    Qoutdated.EnQueue Comp
                    sName = .CompName
                    With Services
                        .NoOfItemsServicedNames = sName
                        .NoOfItemsOutdated = Qoutdated.Size
                        .DsplyProgress "collected Common Components outdated"
                    End With
                Else
                    '~~ When not outdated due to a code difference the revision numbers ought to be equal
                    If .RevisionNumber <> CommComps.RevisionNumber(.CompName) Then
                        .RevisionNumber = CommComps.RevisionNumber(.CompName)
                    End If
                    With Services
                        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                        .LogServicedEntry "Used Common Component is up-to-date"
                    End With
                End If ' .Outdated
            End If
        End With
        Set Comp = Nothing
        Services.DsplyProgress "collected Common Components outdated"
    Next v
    Services.DsplyProgress "collected Common Components outdated"
    
xt: If wsService.CommonComponentsUsed = 0 Then wsService.CommonComponentsUsed = Services.NoOfCommonComponents
    If wsService.CommonComponentsOutdated = 0 Then wsService.CommonComponentsOutdated = Qoutdated.Size
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Reorg()
    PP.Reorg
End Sub

Public Sub SaveToCommonComponentsFolder(ByVal s_comp_name As String, _
                                        ByVal s_file As File, _
                                        ByVal s_file_full_name As String)
' ------------------------------------------------------------------------------
' Save a copy of the hosted raw`s (s_comp_name) export file to the Common
' Components folder which serves as source for the update of Common Components
' used in other VB-Projects.
' ------------------------------------------------------------------------------
    Const PROC  As String = "SaveToCommonComponentsFolder"
    
    Dim frxFile As File
    
    mBasic.BoP ErrSrc(PROC)
    LastModExpFile(s_comp_name) = s_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    If fso.GetExtensionName(s_file_full_name) = "frm" Then
        Set frxFile = fso.GetFile(Replace(s_file_full_name, "frm", "frx"))
        LastModExpFile(s_comp_name) = frxFile
    End If
    
xt: mBasic.EoP ErrSrc(PROC)

End Sub

