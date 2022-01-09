Attribute VB_Name = "mComCompsHosts"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsHosts
' Maintains the raw component's hosts (Workbook-Base-Name) with their full
' name.
' ---------------------------------------------------------------------------
Private Const VNAME_HOST_FULL_NAME  As String = "HostFullName"
Private Const REG_BASE_KEY          As String = "HKLM\CompManVBP\ComCompsHosts\"

Public Property Get ComCompsFolder() As String:    ComCompsFolder = mConfig.FolderServiced & "\Common-Components":   End Property

Public Property Get ComCompsHostsFile() As String: ComCompsHostsFile = ComCompsFolder & "\Hosts.dat":                           End Property

Private Property Get ComCompsHostsReg(ByVal raw_host As String)
    ComCompsHostsReg = REG_BASE_KEY & raw_host
End Property

Public Property Get FullName( _
                     Optional ByVal host_base_name As String) As String
    FullName = Value(pp_section:=host_base_name, pp_value_name:=VNAME_HOST_FULL_NAME)
End Property

Public Property Let FullName( _
                     Optional ByVal host_base_name As String, _
                              ByVal host_full_name As String)
    Value(pp_section:=host_base_name, pp_value_name:=VNAME_HOST_FULL_NAME) = host_full_name
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
    Value = ValueFile(pp_section, pp_value_name)
End Property

Private Property Get ValueFile( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file ComCompsHostsFile.
' ----------------------------------------------------------------------------
    ValueFile = mFile.Value(pp_file:=ComCompsHostsFile _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
End Property

Private Property Let Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
    ValueFile(pp_section, pp_value_name) = pp_value
End Property

Private Property Let ValueFile( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Write value (pp_value) named (pp_value_name) into file 'ComCompsHostsFile'.
' ----------------------------------------------------------------------------
    mFile.Value(pp_file:=ComCompsHostsFile _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value
End Property

Private Property Get ValueReg( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Read value (pp_value) named (pp_value_name) from the Registry.
' ----------------------------------------------------------------------------
    ValueReg = mReg.Value(reg_key:=REG_BASE_KEY & pp_section _
                        , reg_value_name:=pp_value_name _
                         )
End Property

Private Property Let ValueReg( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Write value (pp_value) named (pp_value_name) to the Registry.
' ----------------------------------------------------------------------------
    mReg.Value(reg_key:=REG_BASE_KEY & pp_section _
             , reg_value_name:=pp_value_name _
              ) = pp_value
End Property

Public Function Exists(ByVal raw_host_base_name As String) As Boolean
    Exists = ExistsFile(raw_host_base_name)
End Function

Private Function ExistsFile(ByVal raw_host_base_name As String) As Boolean
    ExistsFile = HostsFile.Exists(Key:=raw_host_base_name)
End Function

Public Function ExistsReg(ByVal raw_host_base_name As String) As Boolean
    ExistsReg = mReg.Exists(REG_BASE_KEY & raw_host_base_name)
End Function

Private Function HostsFile() As Dictionary
    Set HostsFile = mFile.Sections(ComCompsHostsFile)
End Function

Private Function HostsReg() As Dictionary
    Dim v As Variant
    Set HostsReg = New Dictionary
    For Each v In mReg.Keys(REG_BASE_KEY)
        HostsReg.Add v, mReg.Value(reg_key:=REG_BASE_KEY & v, reg_value_name:=VNAME_HOST_FULL_NAME)
    Next v
End Function

Public Sub Remove(ByVal raw_host_base_name As String)
    RemoveFile raw_host_base_name
End Sub

Private Sub RemoveFile(ByVal raw_host_base_name As String)
    mFile.SectionsRemove pp_file:=ComCompsHostsFile _
                       , pp_sections:=raw_host_base_name
End Sub

Private Sub RemoveReg(ByVal raw_host_base_name As String)
    mReg.Delete reg_key:=REG_BASE_KEY & raw_host_base_name
End Sub

