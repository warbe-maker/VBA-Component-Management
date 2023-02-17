Attribute VB_Name = "mCompManIni"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCompManIni
' ---------------------------
' Maintains the CompMan.ini file, when not existing initially filled with the
' current configuration from the wsConfig Worksheet, when existing the
' wsConfig Worksheet is provided with the data from the ini file. The ini
' file is maintained along with the change of the configuration in the
' wsConfig Worksheet.
' ---------------------------------------------------------------------------
Private Const VNAME_FOLDER_ADDIN                As String = "FolderAddin"
Private Const VNAME_FOLDER_EXPORT               As String = "FolderExport"
Private Const VNAME_FOLDER_SRVCD_DEV_AND_TEST   As String = "FolderServicedDevAndTest"
Private Const VNAME_FOLDER_SRVCD_SYNC_ARCHIVE   As String = "FolderServicedSyncArchive"
Private Const VNAME_FOLDER_SRVCD_SYNC_TARGET    As String = "FolderServicedSyncTarget"
Private Const VNAME_FOLDER_COMPMAN_ROOT         As String = "FolderCompManRoot"

Public Property Get CompManIniFileFullName() As String
    Dim wbk As Workbook
    Set wbk = ThisWorkbook
    CompManIniFileFullName = Replace(wbk.FullName, wbk.Name, "CompMan.ini")
End Property

Public Property Get FolderCompManRoot() As String
    FolderCompManRoot = Value(VNAME_FOLDER_COMPMAN_ROOT)
End Property

Public Property Let FolderCompManRootn(ByVal s As String)
    Value(VNAME_FOLDER_COMPMAN_ROOT) = s
End Property

Public Property Get FolderAddin() As String
    FolderAddin = Value(VNAME_FOLDER_ADDIN)
End Property

Public Property Let FolderAddin(ByVal s As String)
    Value(VNAME_FOLDER_ADDIN) = s
End Property

Public Property Get FolderExport() As String
    FolderExport = Value(VNAME_FOLDER_EXPORT)
End Property

Public Property Let FolderExport(ByVal s As String)
    Value(VNAME_FOLDER_EXPORT) = s
End Property

Public Property Get FolderServicedDevAndTest() As String
    FolderServicedDevAndTest = Value(VNAME_FOLDER_SRVCD_DEV_AND_TEST)
End Property

Public Property Let FolderServicedDevAndTest(ByVal s As String)
    Value(VNAME_FOLDER_SRVCD_DEV_AND_TEST) = s
End Property

Public Property Get FolderServicedSyncArchive() As String
    FolderServicedSyncArchive = Value(VNAME_FOLDER_SRVCD_SYNC_ARCHIVE)
End Property

Public Property Let FolderServicedSyncArchive(ByVal s As String)
    Value(VNAME_FOLDER_SRVCD_SYNC_ARCHIVE) = s
End Property

Public Property Get FolderServicedSyncTarget() As String
    FolderServicedSyncTarget = Value(VNAME_FOLDER_SRVCD_SYNC_TARGET)
End Property

Public Property Let FolderServicedSyncTarget(ByVal s As String)
    Value(VNAME_FOLDER_SRVCD_SYNC_TARGET) = s
End Property

Private Property Get Value(Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file CompManIniFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFso.PPvalue(pp_file:=CompManIniFileFullName _
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
' CompManIniFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFso.PPvalue(pp_file:=CompManIniFileFullName _
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
    ErrSrc = "mCompManDat" & "." & sProc
End Function

