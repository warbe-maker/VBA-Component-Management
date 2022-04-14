Attribute VB_Name = "mComCompsRawsSaved"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsRawsSaved
' Maintains in the Common Components folder (ComCompsFolder)
' - a file named (ComCompsSavedFileFullName) for all registered Common
'   Components (i.e. components managed by CompMan services. The file has the
'   following structure:
'   [<component-name]
'   HostFullName=<host-workbook-full-name>
'   HostName=<host-workbook-name>
'   HostBaseName=<hort-workbook-base-name>
'   ExpFileFullName=<export-file-full-name>
'   RevisionNumber=yyyy-mm-dd.n
'   The entries (sections) are maintained along with the Workbook_BeforeSave
'   event via the ExportChangedComponents service. The revision number is
'   increased by one with each save whereby it starts with one for each day.
'
' - copies of the most recently changed Common Components which are the
'   source for the update of outdated Common Components used in VB-Projects.
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_HOST_NAME               As String = "RawHostName"
Private Const VNAME_RAW_HOST_BASE_NAME          As String = "RawHostBaseName"
Private Const VNAME_RAW_HOST_FULL_NAME          As String = "RawHostFullName"
Private Const VNAME_RAW_SAVED_REVISION_NUMBER   As String = "RawRevisionNumber"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME      As String = "RawExpFileFullName"

Public Property Get ComCompsFolder() As String:
    ComCompsFolder = mConfig.FolderServiced & "\Common-Components"
End Property

Public Property Get ComCompsSavedFileFullName() As String
    ComCompsSavedFileFullName = ComCompsFolder & "\ComComps-RawsSaved.dat"
End Property

Public Function SavedExpFileExists(ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an Export File with the name extracted from the
' RawExpFuleFullName exists in the Common Components Folder
' ----------------------------------------------------------------------------
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(RawExpFileFullName(comp_name))
        SavedExpFileExists = .FileExists(ComCompsFolder & "\" & FileName)
    End With
End Function

Public Property Get SavedExpFile(Optional ByVal comp_name) As File
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(RawExpFileFullName(comp_name))
        If .FileExists(ComCompsFolder & "\" & FileName) Then
            Set SavedExpFile = .GetFile(ComCompsFolder & "\" & FileName)
        End If
    End With
End Property

Public Property Let SavedExpFile(Optional ByVal comp_name, _
                                          ByVal comp_exp_file As File)
' ---------------------------------------------------------------------------
' Copies the Raw Export File from its host Workbook location to the Common
' Components Folder from where it is used as the source for the update of
' outdated Used Common Components.
' ---------------------------------------------------------------------------
    comp_name = comp_name ' not used
    comp_exp_file.Copy Destination:=ComCompsFolder & "\" & comp_exp_file.Name, OverWriteFiles:=True
End Property

Public Property Get SavedExpFileFullName(Optional ByVal comp_name As String) As String
' ---------------------------------------------------------------------------
' Returns the Export File Full Name which refers to the Export File saved in
' the Common Components folder.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        SavedExpFileFullName = ComCompsFolder & "\" & .GetFileName(RawExpFileFullName(comp_name))
    End With
End Property

Public Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
End Property

Public Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                ByVal exp_file_full_name As String)
' ---------------------------------------------------------------------------
' Sets the Export-File-Full-Name based on the provided (exp_file_full_name)
' from which only the File Name is used.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
    End With
End Property

Public Property Get RawHostWbFullName(Optional ByVal comp_name As String) As String
    RawHostWbFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_FULL_NAME)
End Property

Public Property Let RawHostWbFullName(Optional ByVal comp_name As String, _
                                            ByVal hst_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_FULL_NAME) = hst_full_name
End Property


Public Property Get RawHostWbName(Optional ByVal comp_name As String) As String
    RawHostWbName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_NAME)
End Property

Public Property Let RawHostWbName(Optional ByVal comp_name As String, _
                                        ByVal host_wb_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_NAME) = host_wb_name
End Property

Public Property Get RawHostWbBaseName(Optional ByVal comp_name As String) As String
    RawHostWbBaseName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_BASE_NAME)
End Property

Public Property Let RawHostWbBaseName(Optional ByVal comp_name As String, _
                                            ByVal host_wb_base_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_BASE_NAME) = host_wb_base_name
End Property

Public Property Get RawSavedRevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RawSavedRevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER)
End Property

Public Property Let RawSavedRevisionNumber( _
                     Optional ByVal comp_name As String, _
                              ByVal comp_rev_no As String)
    Const PROC = "RawSavedRevisionNumber Let"
    On Error GoTo eh
    Dim RevDate As String:  RevDate = Split(comp_rev_no, ".")(0)
    Dim RevNo   As Long:    RevNo = Split(comp_rev_no, ".")(1)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")

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
    Value = mFile.Value(pp_file:=ComCompsSavedFileFullName _
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
' into the file ComCompsSavedFileFullName.
' --------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=ComCompsSavedFileFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

'Public Sub Update(ByVal comp_name As String, _
'                  ByVal exp_file As File)
'' ----------------------------------------------------------------------------
'' Updates the export file in the Common Components folder when appropriate.
'' ----------------------------------------------------------------------------
'    ExpFile(comp_name) = exp_file
'End Sub

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(ComCompsSavedFileFullName)
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
    mFile.RemoveSections pp_file:=ComCompsSavedFileFullName _
                       , pp_sections:=comp_name
End Sub

