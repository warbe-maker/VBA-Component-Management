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
' event via the UpdateUsedCommonComponents service. The revision number is the copy
' of the revision number provided by mComCompsSaved.RevisionNumber.
' ---------------------------------------------------------------------------
Private Const VNAME_REVISION_NUMBER     As String = "RevisionNumber"

Private Property Get UsedRawClonesFile() As String
    Dim wb As Workbook: Set wb = mService.Serviced
    UsedRawClonesFile = Replace(wb.FullName, wb.Name, "ComCompsUsed.dat")
End Property

Public Property Get RevisionNumber( _
                          Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER)
End Property

Public Property Let RevisionNumber( _
                          Optional ByVal comp_name As String, _
                                   ByVal comp_rev_no As String)
    Dim RevDate As String:  RevDate = Split(comp_rev_no, ".")(0)
    Dim RevNo   As Long:    RevNo = Split(comp_rev_no, ".")(1)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
    Const PROC = "Value_Let"
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file UsedRawClonesFile.
' ----------------------------------------------------------------------------
    On Error GoTo eh
    Value = mFile.Value(pp_file:=UsedRawClonesFile _
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
' UsedRawClonesFile.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=UsedRawClonesFile _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(ComCompsFile)
End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove pp_file:=ComCompsFile _
                       , pp_sections:=comp_name
End Sub


