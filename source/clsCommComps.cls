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
' BttnInconsistencyExport          .
' BttnInconsistencySkip            .
' BttnInconsistencyUpdate          .
' CommCompsDatFileFullName         .
'
' Public Properties:
' ------------------
' Hosted                         w Provides the hosted components, provided
'                                  as a comma delimited string in a
'                                  Dictionary (dctHosted) and performs the
'                                  housekeeping for the Common-Components
'                                  folder's CommComps.dat file
' LastModExpFile               r/w .
' LastModExpFileFullName       r   .
' LastModExpFileFullNameOrigin r/w .
' LastModWbk                     w .
' LastModWbkBaseName           r   .
' LastModWbkFullName           r/w .
' LastModWbkName               r   .
' RevisionNumber               r/w .
' ---------------------------------------------------------------------------

' Housekeeping syntax: (allows the on-the-fly-change of value names as well as their removal)
' "<current-name>:<old-name>" = rename in all sections
' "<current>"                 = no action
' ":<remove-name>"            = remove in all sections
Private Const LAST_MOD_EXP_FILE_FULL_NAME           As String = "LAST_MOD_EXP_FILE_FULL_NAME"
Private Const LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN    As String = "LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN"
Private Const LAST_MOD_WBK_FULL_NAME                As String = "LAST_MOD_WBK_FULL_NAME"
Private Const REVISION_NUMBER                       As String = "REVISION_NUMBER"

Private PP                                      As clsPP
Private fso                                     As FileSystemObject
Private dctHosted                               As Dictionary
Private sFolderPath                             As String

Private Sub Class_Initialize()
    Dim dct As New Dictionary
    
    Set dctHosted = New Dictionary
    Set fso = New FileSystemObject
    Set PP = New clsPP
    
    sFolderPath = wsConfig.FolderCommonComponentsPath

    '~~> Full name of the PP-File
    PP.FullName = sFolderPath & "\" & "CommComps.dat"
        
    '~~> Value Names with the syntax (allows the on-the-fly-change of value names as well as their removal):
    '~~> "<current>"                 = no action
    '~~> "<current-name>:<old-name>" = rename in all sections
    '~~> ":<remove-name>"            = remove in all sections
        dct.Add LAST_MOD_EXP_FILE_FULL_NAME, "LastModExpFileFullName"
        dct.Add LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN, "LastModExpFileFullNameOrigin"
        dct.Add LAST_MOD_WBK_FULL_NAME, "LastModWbkFullName"
        dct.Add REVISION_NUMBER, "RevisionNumber"
        dct.Add "Obsolete_1", ":RawRevisionNumber"
    PP.Names = dct
    Set dct = Nothing

    Hskpng
    
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
    Hskpng
    
End Property

Friend Property Get LastModExpFile(Optional ByVal comp_name) As File
' ---------------------------------------------------------------------------
' Returns a component's (comp_name) Export-File in the Common-Components
' folder as file object.
' ---------------------------------------------------------------------------
    Const PROC = "LastModExpFile Get"
    
    On Error GoTo eh
    Dim FileName    As String
    Dim fl          As File
    
    With fso
        For Each fl In .GetFolder(sFolderPath).Files
            If .GetBaseName(fl.Path) = comp_name Then
                Set LastModExpFile = fl
                Exit For
            End If
        Next fl
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
    comp_name = comp_name
    comp_exp_file.Copy Destination:=sFolderPath & comp_exp_file.Name, OverWriteFiles:=True
    Me.LastModExpFileFullNameOrigin(comp_name) = comp_exp_file.Path
End Property

Friend Property Get LastModExpFileFullName(Optional ByVal comp_name As String) As String
' ---------------------------------------------------------------------------
' Returns the Common-Component folder's Export-File Full Name.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        LastModExpFileFullName = sFolderPath & "\" & .GetFileName(LastModExpFileFullNameOrigin(comp_name))
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
        LastModWbkName = fso.GetFileName(LastModWbkFullName(comp_name))
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
    
    sFile = sFolderPath & "\" & x_vbc.Name & x_exp_file_ext
    If fso.FileExists(sFile) Then
        Exists = True
        x_export_file_full_name = sFile
        Set x_export_file = LastModExpFile(x_vbc.Name)
        x_modified_by_wbk_name = LastModWbkName(x_vbc.Name)
        x_last_mod_rev_no = RevisionNumber(x_vbc.Name)
    End If
    
xt: Set Comp = Nothing
    
End Function

Public Function ExistsRegistered(ByVal x_comp_name As String) As Boolean
    ExistsRegistered = Components.Exists(x_comp_name)
End Function

Private Sub Hskpng()
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    
    HskpngSectionsObsolete
    HskpngSectionsMissing
    HskpngUsedOrPrivate
    HskpngHosted
    Reorg
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
' - LastModWbkFullName (value)
' - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "HskpngHosted"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim CommCompHosted  As clsComp
    Dim sHostBaseName   As String
    Dim wbk             As Workbook
    Dim Comp            As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    sHostBaseName = fso.GetBaseName(wbk.FullName)
                    
    For Each v In dctHosted
        If Not mComp.Exists(v, wbk) Then
            MsgBox "The VBComponent " & v & " claimed hosted by the serviced Workbook " & Services.Serviced.Name & _
                   " will be ignored because it does not exist." & vbLf & vbLf & _
                   "The hosted argument should be updated accordingly.", _
                   vbOK, "VBComponent " & v & "does not exist!"
        Else
            '~~ The component claimed hosted exists
            If Not ExistsRegistered(v) Then
                '~~ The component is claimed hosted for the very first time and is yet not
                '~~ known as a Common Component
                Set Comp = New clsComp
                With Comp
                    .Wrkbk = wbk
                    .CompName = v
                    '~~ The component's export file is saved to the Common-Components folder
                    SaveToCommonComponentsFolder v, .ExpFile
                End With
                Set Comp = Nothing
            End If
            
            If LastModWbkFullName(v) = vbNullString Then
                '~~ Yet no other Workbook is registered as the last having modified the component
                Set CommCompHosted = New clsComp
                With CommCompHosted
                    .Wrkbk = wbk
                    .CompName = v
                    If Not Services.FilesDiffer(f_file_1:=.ExpFileTemp _
                                              , f_file_2:=LastModExpFile(v) _
                                              , f_ignore_export_header:=True) _
                    Then
                        '~~ Only when the claiming Workbook's Common Component is identical
                        '~~ with the Export-File in the Common-Components folder it is also
                        '~~ registered as the origin of the Export-File
                        LastModWbk(v) = wbk
                        LastModExpFileFullNameOrigin(v) = .ExpFileFullName
                        RevisionNumber = .RevisionNumber
                    End If
                    
                End With
                Set CommCompHosted = Nothing
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

Private Sub HskpngSectionsObsolete()
' ------------------------------------------------------------------------------
' Removes in the Common-Component folder's CommComps.dat file any section of
' which no corresponding Export-File exists in the folder
' ------------------------------------------------------------------------------
    Const PROC = "HskpngSectionsObsolete"
    
    On Error GoTo eh
    Dim bExists As Boolean
    Dim v       As Variant
    Dim fl      As File
    
    mBasic.BoP ErrSrc(PROC)
    For Each v In Components
        For Each fl In fso.GetFolder(sFolderPath).Files
            If fso.GetBaseName(fl.Path) = v Then
                bExists = True
                Exit For
            End If
        Next fl
        If Not bExists Then
            PP.ComponentRemove v
        End If
    Next v
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngSectionsMissing()
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
    Dim dct         As Dictionary
    Dim sCompName   As String
    Dim sExt        As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dct = Me.Components
    With fso
        For Each fle In .GetFolder(sFolderPath).Files
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
                        LastModWbk(sCompName) = Nothing
                        RevisionNumber(sCompName) = CompManDat.RevisionNumberInitial
                    Else
                        If LastModExpFileFullNameOrigin(sCompName) = vbNullString Then
'                            Debug.Print "The property ""LastModExpFileFullNameOrigin"" of component " & sCompName & " is not available, i.e. its origin is unknown or simply yet not registered respectively!"
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

Private Sub HskpngUsedOrPrivate()
' ----------------------------------------------------------------------------
' Manages the registration of used Common Components, done before change
' components are exported and used changed Common Components are updated.
' When not yet registered a confirmation dialog ensures a component not just
' accidentially has the same name. The type of confirmation is registered
' either as "used" or "private" together with the current revision number.
' When none is available the current date is registered on the fly.
' its Revision-Number is
' ----------------------------------------------------------------------------
    Const PROC          As String = "HskpngUsedOrPrivate"
    
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
            If Exists(vbc, .ExpFileExt) Then
                '~~ The component is a known Common Component
                If Not CompManDat.RegistrationState(.CompName) = enRegStatePrivate _
                And Not CompManDat.RegistrationState(.CompName) = enRegStateUsed _
                And Not CompManDat.RegistrationState(.CompName) = enRegStateHosted _
                Then
                    '~~ Once an equally named VBComponent is registered as private it will no longer be regarded as "used" and updated.
                    Msg.Section(1).Text.Text = "The VBComponent named   " & mBasic.Spaced(.CompName) & "   is known as a ""Common Component"" " & _
                                               "because a coresponding Export-File exists in the ""Common-Components folder""  '" & _
                                               sFolderPath & "'  but is yet registered in the serviced " & _
                                               "Workbook's CompMan.dat file neither as   u s e d   nor as   p r i v a t e !"
                    
                    Select Case mMsg.Dsply(dsply_title:="Not yet registered ""Common Component""" _
                                         , dsply_msg:=Msg _
                                         , dsply_buttons:=mMsg.Buttons(BttnConfirmed, vbLf, BttnPrivate))
                        Case BttnConfirmed: CompManDat.RegistrationState(.CompName) = enRegStateUsed
                                            CompManDat.RevisionNumber(.CompName) = vbNullString ' yet unknown will force update when outdated
                        Case BttnPrivate:   CompManDat.RegistrationState(.CompName) = enRegStatePrivate
                    End Select
                End If
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
    Dim cllBttns            As Collection
    Dim BttnSkip            As String
    Dim BttnExport          As String
    Dim BttnUpdate          As String
    Dim i                   As Long
    
    BttnSkip = BttnInconsistencySkip
    BttnExport = BttnInconsistencyExport
    BttnUpdate = BttnInconsistencyUpdate
    
    Set cllBttns = mMsg.Buttons(BTTN_DSPLY_DIFF, vbLf, BttnExport, BttnUpdate, vbLf, BttnSkip)
    mCompMan.MsgInit
    i = 0
    With Msg
        i = i + 1
        With .Section(i)
            .Label.Text = "Attention!"
            .Label.FontColor = rgbRed
        End With
        i = i + 1
        With .Section(i)
            With .Text.Text = i_message
            End With
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = Replace(Replace(BttnExport, vbLf, " "), "  ", " ")
            .Text.Text = "According to the displayed differencies is the hosted version the one up-to-date. " & vbLf & _
                         "It will be exported and copied to the ""Common-Components Folder"" and the 'Revision Number' will be increased."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = Replace(Replace(BttnUpdate, vbLf, " "), "  ", " ")
            .Text.Text = "The hosted version will be updated with the ""Common-Components Folder"" version (by re-import) and " & _
                         "the 'Revision-Number' of the hosted version will be set identical with the ""Common-Components Folder"" version"
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = Replace(Replace(BttnSkip, vbLf, " "), "  ", " ")
            .Text.Text = "Will be clarified later!& " & vbLf & _
                         "Note: Each Workbook Safe will redisplay this message until either Export or Update is performed."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Background:"
            .Text.Text = "When 'hosted Common Component' is modified within its hosting Workbook and exported, the " & _
                         "'Export File' is copied to the ""Common-Components Folder"" and the 'Revision Number' is increased " & _
                         "and set equal in the hosting Workbook's and the Common-Component Folder's ""CommComps.dat"" file. " & _
                         "When a 'used Common Component is modified within the VB-Project just using (not hosting!) it, " & _
                         "the 'Revision Number' only of this 'used Common Component' is increased. When by accident both " & _
                         "had been modified the 'Revision Numbers' may be equal but the Export Files will differ."
        End With
        
    End With
        
    Do
        Select Case mMsg.Dsply(dsply_title:="Inconsistency warning for " & fso.GetBaseName(i_file_full_name) & "!" _
                             , dsply_msg:=Msg _
                             , dsply_label_spec:="R70" _
                             , dsply_buttons:=cllBttns _
                              )
            Case BTTN_DSPLY_DIFF
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Reorg()
    PP.Reorg
End Sub

Public Sub SaveToCommonComponentsFolder(ByVal s_comp_name As String, _
                                        ByVal s_file As File)
' ------------------------------------------------------------------------------
' Save a copy of a modified Common Component's (s_comp_name) export file (s_file)
' to the Common-Components folder and register the last modified Export-File.
' ------------------------------------------------------------------------------
    Const PROC  As String = "SaveToCommonComponentsFolder"
        
    mBasic.BoP ErrSrc(PROC)
    LastModExpFile(s_comp_name) = s_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    With fso
        .CopyFile s_file, sFolderPath & "\" & .GetFileName(s_file)
        If .GetExtensionName(s_file.Name) = "frm" Then
            '~~ when the save export file is a UserForm, also save the frx-file
            .CopyFile Replace(s_file, "frm", "frx"), sFolderPath & "\" & .GetBaseName(s_file) & ".frx"
        End If
    End With
    
xt: mBasic.EoP ErrSrc(PROC)

End Sub

