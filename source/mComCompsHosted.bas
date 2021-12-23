Attribute VB_Name = "mComCompsHosted"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsHosted
' Maintains in a file named ComCompsHosted.dat for each Workbook which hosts
' at least one Common Component with the following structure:
'
' [<component-name]
' ExpFileFullName=<export-file-full-name>
' RevisionNumber=yyyy-mm-dd.nnn
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification.
' ---------------------------------------------------------------------------
Private Const VNAME_REVISION_NUMBER     As String = "RevisionNumber"
Private Const VNAME_EXP_FILE_FULL_NAME  As String = "ExpFileFullName"

Private Property Get ComCompsHostedFile() As String
    Dim Wb As Workbook: Set Wb = mService.Serviced
    ComCompsHostedFile = Replace(Wb.FullName, Wb.name, "ComCompsHosted.dat")
End Property

Public Property Get ExpFileFullName( _
                     Optional ByVal comp_name As String) As String
    ExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_EXP_FILE_FULL_NAME)
End Property

Public Property Let ExpFileFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_EXP_FILE_FULL_NAME) = exp_file_full_name
End Property

Public Property Get RevisionNumber( _
                          Optional ByVal comp_name As String) As String
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER)
End Property

Public Property Let RevisionNumber( _
                          Optional ByVal comp_name As String, _
                                   ByVal comp_rev_no As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER) = comp_rev_no
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file ComCompsHostedFile.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFile.Value(pp_file:=ComCompsHostedFile _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' ComCompsHostedFile.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=ComCompsHostedFile _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(ComCompsFile)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRaw" & "." & sProc
End Function

Public Function Exists(ByVal raw_comp_name As String) As Boolean
    Exists = Components.Exists(raw_comp_name)
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

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove pp_file:=ComCompsFile _
                       , pp_sections:=comp_name
End Sub

Public Sub RevisionNumberIncrease(ByVal comp_name As String)
' ----------------------------------------------------------------------------
' Increases the revision number by one starting with 1 for a new day.
' ----------------------------------------------------------------------------
    Dim RevNo   As Long
    Dim RevDate As String
    
    If RevisionNumber(comp_name) = vbNullString Then
        RevNo = 1
    Else
        RevNo = Split(RevisionNumber(comp_name), ".")(1)
        RevDate = Split(RevisionNumber(comp_name), ".")(0)
        If RevDate <> Format(Now(), "YYYY-MM-DD") _
        Then RevNo = 1 _
        Else RevNo = RevNo + 1
    End If
    Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER) = Format(Now(), "YYYY-MM-DD") & "." & Format(RevNo, "000")

End Sub

