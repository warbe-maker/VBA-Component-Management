Attribute VB_Name = "mComCompsSaved"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsSaved
' Maintains in the Common Components folfer (ComCompsFolder)
' - a file named (ComCompsFile) for all registered Common Components (i.e.
'   components managed by CompMan services. The file has the following
'   structure:
'   [<component-name]
'   HostFullName=<host-full-name>
'   ExpFileFullName=<export-file-full-name>
'   RevisionNumber=yyyy-mm-dd.n
'   The entries (sections) are maintained along with the Workbook_BeforeSave
'   event via the ExportChangedComponents service. The revision number is
'   increased by one with each save whereby it starts with one for each day.
'
' - copies of the most recently changed Common Components which are the
'   source for the update of outdated Common Components used in VB-Projects.
' ---------------------------------------------------------------------------
Private Const VNAME_HOST_FULL_NAME      As String = "HostFullName"
Private Const VNAME_REVISION_DATE       As String = "RevisionDate"
Private Const VNAME_REVISION_NUMBER     As String = "RevisionNumber"
Private Const VNAME_EXP_FILE_FULL_NAME  As String = "ExpFileFullName"

Public Property Get ComCompsFile() As String
    ComCompsFile = ComCompsFolder & "\Revisions.dat"
End Property

Public Property Get ComCompsFolder() As String
    ComCompsFolder = mConfig.FolderServiced & "\Common-Components"
End Property

Public Property Get ExpFile( _
                    Optional ByVal comp_name) As File
    Dim FileName    As String
    
    With New FileSystemObject
        FileName = .GetFileName(ExpFileFullName(comp_name))
        Set ExpFile = .GetFile(ComCompsFolder & "\" & FileName)
    End With
End Property

Public Property Let ExpFile( _
                    Optional ByVal comp_name, _
                             ByVal comp_exp_file As File)
    Dim FileName As String
    
    With New FileSystemObject
        FileName = .GetFileName(comp_exp_file.name)
        .CopyFile comp_exp_file, ComCompsFolder & "\" & comp_name
    End With
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

Public Property Get HostFullName( _
                     Optional ByVal comp_name As String) As String
    HostFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_HOST_FULL_NAME)
End Property

Public Property Let HostFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal hst_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_HOST_FULL_NAME) = hst_full_name
End Property

Private Property Get RevisionDate( _
                     Optional ByVal comp_name As String) As String
    RevisionDate = Split(Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER), ".")(0)
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
    Const PROC = "RevisionNumber Let"
    On Error GoTo eh
    Dim RevDate As String:  RevDate = Split(comp_rev_no, ".")(0)
    Dim RevNo   As Long:    RevNo = Split(comp_rev_no, ".")(1)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFile.Value(pp_file:=ComCompsFile _
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
' --------------------------------------------------
' Write the value (pp_value) named (pp_value_name)
' into the file RAWS_ComCompsFile.
' --------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=ComCompsFile _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Sub Update(ByVal comp_name As String, _
                  ByVal exp_file As File)
' ----------------------------------------------------------------------------
' Updates the export file in the Common Components folder when appropriate.
' ----------------------------------------------------------------------------
    If RevisionNumber(comp_name) = vbNullString Then
        ExpFile(comp_name) = exp_file
    ElseIf RevisionNumber(comp_name) < mComCompsHosted.RevisionNumber(comp_name) Then
        ExpFile(comp_name) = exp_file
    End If
    RevisionNumber(comp_name) = mComCompsHosted.RevisionNumber(comp_name)
End Sub

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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove pp_file:=ComCompsFile _
                       , pp_sections:=comp_name
End Sub

