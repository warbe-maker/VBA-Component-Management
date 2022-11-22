Attribute VB_Name = "mComCompsUsed"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsUsed
' Maintains in a file (UsedRawClones) for all used cloned raw components,
' i.e. common components managed by CompMan services. The file has the
' following structure:
'
' [<component-name]
' RevisionNumber=yyyy-mm-dd.n
'
' The entries (sections) are maintained along with the Workbook_Open
' event via the UpdateOutdatedCommonComponents service. The revision number is the copy
' of the revision number provided by mComCompRawsGlobal.RawSavedRevisionNumber.
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_REVISION_NUMBER As String = "RawRevisionNumber"
Private Const VNAME_DUE_MODIF_WARNING   As String = "DueModificationWarning"

Public Property Get DueModificationWarning( _
                          Optional ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    If NameExists(pp_section:=comp_name, pp_value_name:=VNAME_DUE_MODIF_WARNING) _
    Then DueModificationWarning = CBool(Value(pp_section:=comp_name, pp_value_name:=VNAME_DUE_MODIF_WARNING))
End Property

Public Property Let DueModificationWarning( _
                          Optional ByVal comp_name As String, _
                                   ByVal comp_due_warning As Boolean)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_DUE_MODIF_WARNING) = Abs(CInt(comp_due_warning))
End Property

Public Property Get RevisionNumber( _
                          Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER)
End Property

Public Property Let RevisionNumber( _
                          Optional ByVal comp_name As String, _
                                   ByVal comp_rev_no As String)
    Dim RevDate As String:  RevDate = Split(comp_rev_no, ".")(0)
    Dim RevNo   As Long:    RevNo = Split(comp_rev_no, ".")(1)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")
End Property

Private Property Get ComCompsUsedFile() As String
' ----------------------------------------------------------------------------
' Returns the name of the ComCompdUsed.dat file for the serviced Workbook.
' ----------------------------------------------------------------------------
    ComCompsUsedFile = Replace(mService.WbkServiced.FullName, mService.WbkServiced.name, "ComCompsUsed.dat")
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
    Const PROC = "Value_Let"
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file ComCompsUsedFile.
' ----------------------------------------------------------------------------
    On Error GoTo eh
    Value = mFso.FilePrivProfValue(pp_file:=ComCompsUsedFile _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
' ComCompsUsedFile.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFso.FilePrivProfValue(pp_file:=ComCompsUsedFile _
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
    Set Components = mFso.FilePrivProfSectionNames(ComCompsSavedFileFullName)
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

Private Function NameExists(ByVal pp_section As String, _
                            ByVal pp_value_name As String) As Boolean
    NameExists = mFso.Exists(ex_file:=ComCompsUsedFile _
                            , ex_section:=pp_section _
                            , ex_value_name:=pp_value_name)
End Function

Public Sub Remove(ByVal comp_name As String)
    mFso.FilePrivProfRemoveSections pp_file:=ComCompsSavedFileFullName _
                                  , pp_sections:=comp_name
End Sub

