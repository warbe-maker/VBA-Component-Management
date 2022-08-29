Attribute VB_Name = "mComCompsRawsHosted"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsRawsHosted
' Maintains in a file named ComCompsHosted.dat for each Workbook which hosts
' at least one Raw Common Component with the following structure:
'
' [<component-name]
' ExpFileFullName=<export-file-full-name>
' RevisionNumber=yyyy-mm-dd.nnn
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification.
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_REVISION_NUMBER     As String = "RawRevisionNumber"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME  As String = "RawExpFileFullName"

Public Property Get ComCompsHostedFileFullName() As String
    Dim wbk As Workbook: Set wbk = mService.Serviced
    ComCompsHostedFileFullName = Replace(wbk.FullName, wbk.Name, "ComCompsHosted.dat")
End Property

Public Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
End Property

Public Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                ByVal exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
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
' file ComCompsHostedFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFile.Value(pp_file:=ComCompsHostedFileFullName _
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
' ComCompsHostedFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=ComCompsHostedFileFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(ComCompsHostedFileFullName)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRaw" & "." & sProc
End Function

Public Function Exists(ByVal raw_vbc_name As String) As Boolean
    Exists = Components.Exists(raw_vbc_name)
End Function

Public Function MaxRawLenght() As Long
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

Public Sub Remove(ByVal comp_name As String)
    mFile.RemoveSections pp_file:=ComCompsHostedFileFullName _
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
        Else RevNo = RevNo + 1
    End If
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = Format(Now(), "YYYY-MM-DD") & "." & Format(RevNo, "000")

End Sub

