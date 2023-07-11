Attribute VB_Name = "mCompManDat"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCompManDat: Maintains for each serviced Workbook a file
' ============================ named CompMan.dat primarily for the
' Workbook's hosted and used Common Components with the following structure:
'
' [<component-name>]
' KindOfComponent=used|hosted
' ExpFileFullName=<export-file-full-name>
' RevisionNumber=yyyy-mm-dd.nnn
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification of a hosted Common Component.
'
' W. Rauschenberger Berlin Feb. 2023
' ---------------------------------------------------------------------------

' Housekeeping syntax (allows the on-the-fly-change of value names as well as their removal):
' -------------------------------------------------------------------------------------------
' "<current-name>:<old-name>" = rename in all sections
' "<current>"                 = no action
' ":<remove-name>"            = remove in all sections
Private Const SECTION_NAME_RECENT_EXPORT            As String = "_MostRecentExport" ' _ avoids conflict with an existing VBComponent
Private Const VALUE_NAME_LAST_MOD_REVISION_NUMBER   As String = "RevisionNumber:RawRevisionNumber"
Private Const VALUE_NAME_REG_STAT_OF_COMPONENT      As String = "KindOfComponent"
Private Const VALUE_NAME_USED_EXPORT_FOLDER         As String = "UsedExportFolder"

Public Property Get CompManDatFileFullName() As String
' ----------------------------------------------------------------------------
' Returns the current serviced Workbook's full name of the CompMan.dat file
' Precondition: Services.Serviced is set.
' ----------------------------------------------------------------------------
    Const PROC  As String = "CompManDatFileFullName-Get"
    Dim wbk     As Workbook
    Dim fso     As New FileSystemObject
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    CompManDatFileFullName = Replace(wbk.FullName, wbk.Name, "CompMan.dat")
    If Not fso.FileExists(CompManDatFileFullName) Then
        fso.CreateTextFile CompManDatFileFullName
    End If
    Set fso = Nothing

xt: mBasic.EoP ErrSrc(PROC)
End Property

Private Property Get CompRegState(Optional ByVal comp_name As String) As enCommCompRegState
    CompRegState = mCommComps.CommCompRegStateEnum(Value(pp_section:=comp_name, pp_value_name:=CurrentName("VALUE_NAME_REG_STAT_OF_COMPONENT")))
End Property

Private Property Let CompRegState(Optional ByVal comp_name As String, _
                                           ByVal comp_reg_state As enCommCompRegState)
    Value(pp_section:=comp_name, pp_value_name:=CurrentName("VALUE_NAME_REG_STAT_OF_COMPONENT")) = mCommComps.CommCompRegStateString(comp_reg_state)
End Property

Public Property Get RecentlyUsedExportFolder() As String
    RecentlyUsedExportFolder = Value(pp_section:=SECTION_NAME_RECENT_EXPORT, pp_value_name:=CurrentName("VALUE_NAME_USED_EXPORT_FOLDER"))
End Property

Public Property Let RecentlyUsedExportFolder(ByVal s As String)
    Value(pp_section:=SECTION_NAME_RECENT_EXPORT, pp_value_name:=CurrentName("VALUE_NAME_USED_EXPORT_FOLDER")) = s
End Property

Public Property Get RegistrationState(Optional ByVal comp_name As String) As enCommCompRegState
    RegistrationState = CompRegState(comp_name)
End Property

Public Property Let RegistrationState(Optional ByVal comp_name As String, _
                                               ByVal comp_reg_state As enCommCompRegState)
    CompRegState(comp_name) = comp_reg_state
End Property

Public Property Get RevisionNumber(Optional ByVal r_comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=r_comp_name, pp_value_name:=CurrentName("VALUE_NAME_LAST_MOD_REVISION_NUMBER"))
End Property

Public Property Let RevisionNumber(Optional ByVal r_comp_name As String, _
                                            ByVal r_rev_no As String)
    Value(pp_section:=r_comp_name, pp_value_name:=CurrentName("VALUE_NAME_LAST_MOD_REVISION_NUMBER")) = r_rev_no
End Property

Private Property Get Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file CompManDatFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFso.PPvalue(pp_file:=CompManDatFileFullName _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' CompManDatFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Debug.Print "Write " & pp_value_name & "="; pp_value & " into:" & CompManDatFileFullName
    mFso.PPvalue(pp_file:=CompManDatFileFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function CommCompUsedIsKnown(ByVal comp_name As String) As Boolean
    CommCompUsedIsKnown = ComponentsRegistered(enRegStateUsed).Exists(comp_name)
End Function

Private Function Components() As Dictionary
    Set Components = mFso.PPsectionNames(CompManDatFileFullName)
End Function

Private Function ComponentsRegistered(ByVal c_reg_state As enCommCompRegState) As Dictionary
    Dim dct         As New Dictionary
    Dim vSection    As Variant
    
    For Each vSection In mFso.PPsectionNames(CompManDatFileFullName)
        If mFso.PPvalue(pp_file:=CompManDatFileFullName _
                                , pp_section:=vSection _
                                , pp_value_name:=CurrentName("VALUE_NAME_REG_STAT_OF_COMPONENT")) = CommCompRegStateString(c_reg_state) _
        Then
            dct.Add vSection, vbNullString
        End If
    Next vSection
    Set ComponentsRegistered = dct
    Set dct = Nothing
    
End Function

Private Function CurrentName(ByVal sNameConst As String) As String
    Dim v   As Variant
    Dim dct As Dictionary
    
    Set dct = HskpngValueNamesCurrent
    For Each v In dct
        If v = sNameConst Then
            CurrentName = Split(dct(v), ":")(0)
        End If
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManDat" & "." & sProc
End Function

Public Sub Hskpng(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    HskpngRemoveObsoleteSections h_hosted
    HskpngHosted h_hosted
    mHskpng.ReorgDatFile CompManDatFileFullName

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngHosted(ByVal h_hosted As String)
    Const PROC      As String = "HskpngHosted"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim dctHosted   As Dictionary
    Dim wbk         As Workbook
    Dim v           As Variant
    Dim Comp        As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    Set dctHosted = mCommComps.Hosted(h_hosted)
    
    For Each v In dctHosted
        If mComp.Exists(v, wbk) Then
            If mCompManDat.RegistrationState(v) <> enRegStateHosted _
            Then mCompManDat.RegistrationState(v) = enRegStateHosted
            If mCommComps.LastModExpFileFullNameOrigin = vbNullString Then
                Set Comp = New clsComp
                With Comp
                    .Wrkbk = Services.Serviced
                    .CompName = v
                    If Services.FilesDiffer(.ExpFile, fso.GetFile(mCommComps.LastModExpFileFullName(.CompName))) Then
                        Err.Raise AppErr(1), ErrSrc(PROC), _
                                  "The origin of the Export-File saved in the Common-Components folder of the " & _
                                  "Common Component " & .CompName & " is unknown! Because it is not identical " & _
                                  "with the Export-File of the Workbook which claims hosting it, the origin " & _
                                  "is still to be clarified!"
                    Else
                        mCommComps.LastModExpFileFullNameOrigin(.CompName) = .ExpFileFullName
                    End If
                    
                End With
            End If
        Else
            mCompManDat.RemoveComponent v
        End If
    Next v
    
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngRemoveObsoleteSections(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Remove sections representing VBComponents no longer existing and those with an
' invalid name.
' ------------------------------------------------------------------------------
    Const PROC = "HskpngRemoveObsoleteSections"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    Dim dctHosted   As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Set dctHosted = mCommComps.Hosted(h_hosted)
    Set wbk = Services.Serviced
    For Each v In mCompManDat.Components
        If HskpngSectionIsInvalid(v, wbk) Then
            mCompManDat.RemoveComponent v
        End If
    Next v
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function HskpngSectionIsInvalid(ByVal h_section As String, _
                                        ByVal h_wbk As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the section (h_section) is invalid, which is the case when
' it is neither an existing VBComponent's Name nor another valid section's Name.
' ------------------------------------------------------------------------------
    Select Case True
        Case mComp.Exists(h_section, h_wbk)
        Case h_section = SECTION_NAME_RECENT_EXPORT
        Case Else
            HskpngSectionIsInvalid = True
    End Select
End Function

Public Function HskpngValueNames() As Boolean
' ----------------------------------------------------------------------------
' Renames all value names with a syntax <new>:<old> in all sections,
' Removes all value names with a syntax :<old>
' When at least one housekeeping action had been performed the function
' returns TRUE.
' ----------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    Dim dctNames    As Dictionary
    Dim dctSects    As Dictionary
    Dim i           As Long
    Dim sNew        As String
    Dim sOld        As String
    Dim v           As Variant
    Dim vName       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set dctNames = HskpngValueNamesCurrent
    
    For i = 0 To dctNames.Count - 1
        vName = Split(dctNames.Items(i), ":")
        If UBound(vName) = 1 Then
            '~~> <new>:<old> or :<remove>
            If vName(0) = vbNullString Then
                '~~ Remove the name in all sections
                Set dctSects = ValueSections
                For Each v In dctSects
                    If ValueNameRemoveInAllSections(vName(1), v) Then
                        HskpngValueNames = True
                    End If
                Next v
            Else
                '~~> Rename the old name with the new name in all sections
                sNew = vName(0)
                sOld = vName(1)
                If sOld <> vbNullString And sNew <> vbNullString Then
                    If ValueNameRenameInAllSections(sOld, sNew) = True Then
                        HskpngValueNames = True
                    End If
                End If
            End If
        End If
    Next i

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function HskpngValueNamesCurrent() As Dictionary

    Static dct As Dictionary
    If dct Is Nothing Then
        Set dct = New Dictionary
        dct.Add "SECTION_NAME_RECENT_EXPORT", _
                 SECTION_NAME_RECENT_EXPORT
        dct.Add "VALUE_NAME_LAST_MOD_REVISION_NUMBER ", _
                 VALUE_NAME_LAST_MOD_REVISION_NUMBER
        dct.Add "VALUE_NAME_REG_STAT_OF_COMPONENT", _
                 VALUE_NAME_REG_STAT_OF_COMPONENT
        dct.Add "VALUE_NAME_USED_EXPORT_FOLDER", _
                 VALUE_NAME_USED_EXPORT_FOLDER
    End If
    Set HskpngValueNamesCurrent = dct
    
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

Private Function NameExists(ByVal pp_section As String, _
                            ByVal pp_value_name As String) As Boolean
    NameExists = mFso.Exists(x_file:=CompManDatFileFullName _
                           , x_section:=pp_section _
                           , x_value_name:=pp_value_name)
End Function

Public Sub RemoveComponent(ByVal r_comp_name As String)
    mFso.PPremoveSections pp_file:=CompManDatFileFullName, pp_sections:=r_comp_name
End Sub

Public Function RevisionNumberInitial() As String
' ----------------------------------------------------------------------------
' Returns an initial revision number in the form: YYYY-MM-DD.001
' ----------------------------------------------------------------------------
    RevisionNumberInitial = Format(Now(), "YYYY-MM-DD") & ".001"
End Function

Private Function ValueNameRemoveInAllSections(ByVal pp_name As String, _
                                     Optional ByVal pp_section As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As Dictionary
    Dim v   As Variant
    
    Set dct = ValueSections
    
    For Each v In dct
        If mFso.PPremoveNames(pp_file:=CompManDatFileFullName _
                            , pp_section:=v _
                            , pp_value_names:=pp_name) Then
            ValueNameRemoveInAllSections = True
        End If
    Next v
    
End Function

Private Function ValueNameRenameInAllSections(ByVal v_old As String, _
                                              ByVal v_new As String) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim dct     As Dictionary
    Dim v       As Variant
    Dim bReorg  As Boolean
    Dim i       As Long
    
    Set dct = ValueSections
    For Each v In dct
        i = i + 1
        bReorg = i = dct.Count
        If mFso.PPvalueNameRename(v_old, v_new, CompManDatFileFullName, v, bReorg) Then
            ValueNameRenameInAllSections = True
        End If
    Next v
    
End Function

Private Function ValueSections() As Dictionary
    Static dct As Dictionary
    If dct Is Nothing Then Set dct = mFso.PPsections(CompManDatFileFullName)
    Set ValueSections = dct
End Function

