Attribute VB_Name = "mCompManCfg"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCompManCfg
' ---------------------------
' Maintains the CompMan.cfg file, when not existing initially filled with the
' current configuration from the wsConfig Worksheet, when existing the
' wsConfig Worksheet is provided with the data from the ini file. The ini
' file is maintained along with the change of the configuration in the
' wsConfig Worksheet.
' ---------------------------------------------------------------------------
Private Const VNAME_FOLDER_ADDIN                As String = "FolderAddin"
Private Const VNAME_FOLDER_COMMON_COMPONENTS    As String = "CommonComponentsFolder"
Private Const VNAME_FOLDER_COMPMAN_ROOT         As String = "FolderCompManRoot"
Private Const VNAME_FOLDER_EXPORT               As String = "FolderExport"
Private Const VNAME_FOLDER_SRVCD_COMPMAN_ROOT   As String = "FolderServicedCompManRoot"
Private Const VNAME_FOLDER_SRVCD_SYNC_ARCHIVE   As String = "FolderServicedSyncArchive"
Private Const VNAME_FOLDER_SRVCD_SYNC_TARGET    As String = "FolderServicedSyncTarget"

Public Property Get CompManCfgFileFullName() As String
    CompManCfgFileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "CompMan.cfg")
End Property

Public Property Get FolderAddin() As String:                        FolderAddin = Value(VNAME_FOLDER_ADDIN):                                End Property

Public Property Let FolderAddin(ByVal s As String):                 Value(VNAME_FOLDER_ADDIN) = s:                                          End Property

Public Property Get FolderCommonComponents() As String:             FolderCommonComponents = Value(VNAME_FOLDER_COMMON_COMPONENTS):         End Property

Public Property Let FolderCommonComponents(ByVal s As String):      Value(VNAME_FOLDER_COMMON_COMPONENTS) = s:                              End Property

Public Property Get FolderCompManRoot() As String:                  FolderCompManRoot = Value(VNAME_FOLDER_COMPMAN_ROOT):                   End Property

Public Property Let FolderCompManRoot(ByVal s As String):           Value(VNAME_FOLDER_COMPMAN_ROOT) = s:                                   End Property

Public Property Get FolderExport() As String:                       FolderExport = Value(VNAME_FOLDER_EXPORT):                              End Property

Public Property Let FolderExport(ByVal s As String):                Value(VNAME_FOLDER_EXPORT) = s:                                         End Property

Public Property Get FolderServicedCompManRoot() As String:          FolderServicedCompManRoot = Value(VNAME_FOLDER_SRVCD_COMPMAN_ROOT):     End Property

Public Property Let FolderServicedCompManRoot(ByVal s As String):    Value(VNAME_FOLDER_SRVCD_COMPMAN_ROOT) = s:                            End Property

Public Property Get FolderServicedSyncArchive() As String:          FolderServicedSyncArchive = Value(VNAME_FOLDER_SRVCD_SYNC_ARCHIVE):     End Property

Public Property Let FolderServicedSyncArchive(ByVal s As String):   Value(VNAME_FOLDER_SRVCD_SYNC_ARCHIVE) = s:                             End Property

Public Property Get FolderServicedSyncTarget() As String:           FolderServicedSyncTarget = Value(VNAME_FOLDER_SRVCD_SYNC_TARGET):       End Property

Public Property Let FolderServicedSyncTarget(ByVal s As String):    Value(VNAME_FOLDER_SRVCD_SYNC_TARGET) = s:                              End Property

Private Property Get Value(Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file CompManCfgFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFso.PPvalue(pp_file:=CompManCfgFileFullName _
                      , pp_section:="Config" _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value(Optional ByVal pp_value_name As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' CompManCfgFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFso.PPvalue(pp_file:=CompManCfgFileFullName _
              , pp_section:="Config" _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManCfg" & "." & sProc
End Function

Public Function Exists() As Boolean
    With New FileSystemObject
        Exists = .FileExists(CompManCfgFileFullName)
    End With
End Function

