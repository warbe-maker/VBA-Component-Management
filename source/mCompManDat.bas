Attribute VB_Name = "mCompManDat"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCompManDat: Maintains for each serviced Workbook a file
' ---------------------------- named CompMan.dat primarily for the
'                              Workbook's hosted and used Common Components
'                              with the following structure:
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
Private Const SECTION_NAME_RECENT_EXPORT        As String = "_MostRecentExport" ' _ avoids conflict with an existing VBComponent
Private Const VALUE_NAME_DUE_MODIF_WARNING      As String = "DueModificationWarning"
Private Const VALUE_NAME_RAW_EXP_FILE_FULL_NAME As String = "RawExpFileFullName"
Private Const VALUE_NAME_RAW_REVISION_NUMBER    As String = "RawRevisionNumber"
Private Const VALUE_NAME_REG_STAT_OF_COMPONENT  As String = "KindOfComponent"
Private Const VALUE_NAME_USED_EXPORT_FOLDER     As String = "UsedExportFolder"

Private Property Get CompManDatFileFullName() As String
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

Public Property Let RegistrationState(Optional ByVal comp_name As String, _
                                               ByVal comp_reg_state As enCommCompRegState)
    CompRegState(comp_name) = comp_reg_state
    mHskpng.ReorgDatFile CompManDatFileFullName

End Property

Public Property Get RegistrationState(Optional ByVal comp_name As String) As enCommCompRegState
    RegistrationState = CompRegState(comp_name)
End Property

Public Function CommCompUsedIsKnown(ByVal comp_name As String) As Boolean
    CommCompUsedIsKnown = ComponentsRegistered(enRegStateUsed).Exists(comp_name)
End Function

Public Property Get RecentlyUsedExportFolder() As String
    RecentlyUsedExportFolder = Value(pp_section:=SECTION_NAME_RECENT_EXPORT, pp_value_name:=VALUE_NAME_USED_EXPORT_FOLDER)
End Property

Public Property Let RecentlyUsedExportFolder(ByVal s As String)
    Value(pp_section:=SECTION_NAME_RECENT_EXPORT, pp_value_name:=VALUE_NAME_USED_EXPORT_FOLDER) = s
End Property

Public Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_RAW_EXP_FILE_FULL_NAME)
End Property

Public Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                ByVal exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
    Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_REG_STAT_OF_COMPONENT) = "hosted"
End Property

Public Property Get RawRevisionNumber(Optional ByVal comp_name As String) As String
    RawRevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_RAW_REVISION_NUMBER)
End Property

Public Property Let RawRevisionNumber(Optional ByVal comp_name As String, _
                                               ByVal comp_rev_no As String)
    Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_RAW_REVISION_NUMBER) = comp_rev_no
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

Private Function ComponentsRegistered(ByVal c_reg_state As enCommCompRegState) As Dictionary
    Dim dct         As New Dictionary
    Dim vSection    As Variant
    
    For Each vSection In mFso.PPsectionNames(CompManDatFileFullName)
        If mFso.PPvalue(pp_file:=CompManDatFileFullName _
                                , pp_section:=vSection _
                                , pp_value_name:=VALUE_NAME_REG_STAT_OF_COMPONENT) = CommCompRegStateString(c_reg_state) _
        Then
            dct.Add vSection, vbNullString
        End If
    Next vSection
    Set ComponentsRegistered = dct
    Set dct = Nothing
    
End Function

Private Property Get CompRegState(Optional ByVal comp_name As String) As enCommCompRegState
    CompRegState = mCommComps.CommCompRegStateEnum(Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_REG_STAT_OF_COMPONENT))
End Property

Private Property Let CompRegState(Optional ByVal comp_name As String, _
                                           ByVal comp_reg_state As enCommCompRegState)
    Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_REG_STAT_OF_COMPONENT) = mCommComps.CommCompRegStateString(comp_reg_state)
End Property

Private Function Components() As Dictionary
    Set Components = mFso.PPsectionNames(CompManDatFileFullName)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManDat" & "." & sProc
End Function

'Public Function CompIsRegistered(ByVal raw_comp_name As String, ByVal comp_reg_state As enCommCompRegState) As Boolean
'    CompIsRegistered = CompRegState(raw_comp_name) = comp_reg_state
'End Function

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

Public Function RevisionNumberInitial() As String
' ----------------------------------------------------------------------------
' Returns an initial revision number in the form: YYYY-MM-DD.001
' ----------------------------------------------------------------------------
    RevisionNumberInitial = Format(Now(), "YYYY-MM-DD") & ".001"
End Function

Public Property Get RevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_RAW_REVISION_NUMBER)
End Property

Public Property Let RevisionNumber(Optional ByVal comp_name As String, _
                                            ByVal comp_rev_no As String)
    Dim RevDate As String:  RevDate = Split(comp_rev_no, ".")(0)
    Dim RevNo   As Long:    RevNo = Split(comp_rev_no, ".")(1)
    Value(pp_section:=comp_name, pp_value_name:=VALUE_NAME_RAW_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")
End Property

Private Function NameExists(ByVal pp_section As String, _
                            ByVal pp_value_name As String) As Boolean
    NameExists = mFso.Exists(ex_file:=CompManDatFileFullName _
                           , ex_section:=pp_section _
                           , ex_value_name:=pp_value_name)
End Function

Public Sub RemoveComponent(ByVal r_comp_name As String)
    mFso.PPremoveSections pp_file:=CompManDatFileFullName, pp_sections:=r_comp_name
End Sub

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
    HskpngRemoveDueModificationWarning
    mHskpng.ReorgDatFile CompManDatFileFullName

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngRemoveDueModificationWarning()
    Dim dct As Dictionary
    Dim v   As Variant
    
    Set dct = mFso.PPsections(CompManDatFileFullName)
    For Each v In dct
        mFso.PPremoveNames pp_file:=CompManDatFileFullName _
                         , pp_section:=v _
                         , pp_value_names:=VALUE_NAME_DUE_MODIF_WARNING
    Next v
    
End Sub

Private Sub HskpngHosted(ByVal h_hosted As String)
    Const PROC      As String = ""
    Dim dctHosted   As Dictionary
    Dim wbk         As Workbook
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    Set dctHosted = mCommComps.Hosted(h_hosted)
    
    For Each v In dctHosted
        If mComp.Exists(v, wbk) Then
            If mCompManDat.RegistrationState(v) <> enRegStateHosted _
            Then mCompManDat.RegistrationState(v) = enRegStateHosted
        Else
            mCompManDat.RemoveComponent v
        End If
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
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

