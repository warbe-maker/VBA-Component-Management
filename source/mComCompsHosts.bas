Attribute VB_Name = "mComCompsHosts"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsHosts
' Maintains for raw components identified by their
' component name the values HostFullName and ExpFileFullName.
' ---------------------------------------------------------------------------
Private Const VNAME_HOST_FULL_NAME = "HostFullName"

Public Property Get ComCompsFolder() As String:    ComCompsFolder = mMe.ServicedRootFolder & "\Common-Components":    End Property
Public Property Get ComCompsHostsFile() As String: ComCompsHostsFile = ComCompsFolder & "\ComCompsHosts.dat":       End Property

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
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file ComCompsHostsFile.
' ----------------------------------------------------------------------------
    
    Value = mFile.Value(pp_file:=ComCompsHostsFile _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
End Property

Private Property Let Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' --------------------------------------------------
' Write the value (pp_value) named (pp_value_name)
' into the file 'ComCompsHostsFile'.
' --------------------------------------------------
    
    mFile.Value(pp_file:=ComCompsHostsFile _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

End Property

Public Function Exists(ByVal raw_host_base_name As String) As Boolean
    Exists = Hosts.Exists(Key:=raw_host_base_name)
End Function

Public Function Hosts() As Dictionary
    Set Hosts = mFile.Sections(ComCompsHostsFile)
End Function

Public Sub Remove(ByVal raw_host_base_name As String)
    mFile.SectionsRemove pp_file:=ComCompsHostsFile _
                       , pp_sections:=raw_host_base_name
End Sub
