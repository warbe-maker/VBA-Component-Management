Attribute VB_Name = "mCompManDat"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCompManDat
' Maintains for each serviced Workbook a file named CompMan.dat primarily for
' the Workbook's hosted and used Common Components with the following
' structure:
'
' [<vb-component-name>]
' KindOfComponent=used|hosted
' ExpFileFullName=<export-file-full-name>
' RevisionNumber=yyyy-mm-dd.nnn
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification.
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_REVISION_NUMBER     As String = "RawRevisionNumber"
Private Const VNAME_DUE_MODIF_WARNING       As String = "DueModificationWarning"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME  As String = "RawExpFileFullName"
Private Const VNAME_RECENT_USED_EXP_FOLDER  As String = "RecentlyUsedExportFolder"
Private Const SNAME_EXPORT                  As String = "Export"
Private Const VNAME_REG_STAT_OF_COMPONENT   As String = "KindOfComponent"

Private Property Get CompManDatFileFullName() As String
    Dim wbk As Workbook
    Dim fso As New FileSystemObject
    
    Set wbk = mService.WbkServiced
    CompManDatFileFullName = Replace(wbk.FullName, wbk.Name, "CompMan.dat")
    If Not fso.FileExists(CompManDatFileFullName) Then
        fso.CreateTextFile CompManDatFileFullName
    End If
    Set fso = Nothing
    
End Property

Public Sub Register(ByVal comp_name As String, _
                    ByVal comp_reg_state As enCommCompRegState)
    CompRegState(comp_name) = comp_reg_state
End Sub

Public Function CommCompUsedIsKnown(ByVal comp_name As String) As Boolean
    CommCompUsedIsKnown = ComponentsRegistered(enRegStateUsed).Exists(comp_name)
End Function

Public Property Get RecentlyUsedExportFolder() As String
    RecentlyUsedExportFolder = Value(pp_section:=SNAME_EXPORT, pp_value_name:=VNAME_RECENT_USED_EXP_FOLDER)
End Property

Public Property Let RecentlyUsedExportFolder(ByVal s As String)
    Value(pp_section:=SNAME_EXPORT, pp_value_name:=VNAME_RECENT_USED_EXP_FOLDER) = s
End Property

Public Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
End Property

Public Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                ByVal exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
    Value(pp_section:=comp_name, pp_value_name:=VNAME_REG_STAT_OF_COMPONENT) = "hosted"
End Property

Public Property Get RawRevisionNumber(Optional ByVal comp_name As String) As String
    RawRevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER)
End Property

Public Property Let RawRevisionNumber(Optional ByVal comp_name As String, _
                                               ByVal comp_rev_no As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = comp_rev_no
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
                                , pp_value_name:=VNAME_REG_STAT_OF_COMPONENT) = CommCompRegStateString(c_reg_state) _
        Then
            dct.Add vSection, vbNullString
        End If
    Next vSection
    Set ComponentsRegistered = dct
    Set dct = Nothing
    
End Function

Private Property Get CompRegState(Optional ByVal comp_name As String) As enCommCompRegState
    CompRegState = CommCompRegStateEnum(Value(pp_section:=comp_name, pp_value_name:=VNAME_REG_STAT_OF_COMPONENT))
End Property

Private Property Let CompRegState(Optional ByVal comp_name As String, _
                                           ByVal comp_reg_state As enCommCompRegState)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_REG_STAT_OF_COMPONENT) = CommCompRegStateString(comp_reg_state)
End Property

Public Function Components() As Dictionary
    Set Components = mFso.PPsectionNames(CompManDatFileFullName)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManDat" & "." & sProc
End Function

Public Function CompIsRegistered(ByVal raw_comp_name As String, ByVal comp_reg_state As enCommCompRegState) As Boolean
    CompIsRegistered = CompRegState(raw_comp_name) = comp_reg_state
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

Private Sub Remove(ByVal comp_name As String)
    mFso.PPremoveSections pp_file:=CompManDatFileFullName _
                                  , pp_sections:=comp_name
End Sub

Public Sub RawRevisionNumberIncrease(ByVal comp_name As String)
' ----------------------------------------------------------------------------
' Increases the revision number by one starting with 1 for a new day.
' ----------------------------------------------------------------------------
    Dim RevNo   As Long
    Dim RevDate As String
    
    If RawRevisionNumber(comp_name) = vbNullString Then
        RevNo = 1
    Else
        RevNo = Split(RawRevisionNumber(comp_name), ".")(1)
        RevDate = Split(RawRevisionNumber(comp_name), ".")(0)
        If RevDate <> Format(Now(), "YYYY-MM-DD") _
        Then RevNo = 1 _
        Else: RevNo = RevNo + 1
    End If
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = Format(Now(), "YYYY-MM-DD") & "." & Format(RevNo, "000")

End Sub

Public Property Get DueModificationWarning(Optional ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    If NameExists(pp_section:=comp_name, pp_value_name:=VNAME_DUE_MODIF_WARNING) _
    Then DueModificationWarning = CBool(Value(pp_section:=comp_name, pp_value_name:=VNAME_DUE_MODIF_WARNING))
End Property

Public Property Let DueModificationWarning(Optional ByVal comp_name As String, _
                                                    ByVal comp_due_warning As Boolean)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_DUE_MODIF_WARNING) = Abs(CInt(comp_due_warning))
End Property

Public Property Get RevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER)
End Property

Public Property Let RevisionNumber(Optional ByVal comp_name As String, _
                                            ByVal comp_rev_no As String)
    Dim RevDate As String:  RevDate = Split(comp_rev_no, ".")(0)
    Dim RevNo   As Long:    RevNo = Split(comp_rev_no, ".")(1)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")
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

Public Sub Hskpng()
' ------------------------------------------------------------------------------
' Remove sections for VBComponents no longer existing
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    Set wbk = mService.WbkServiced
    For Each v In mCompManDat.Components
        If Not v = "Export" Then
            If Not mComp.Exists(v, wbk) Then
                mCompManDat.RemoveComponent v
            End If
        End If
    Next v
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

