Attribute VB_Name = "mHost"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mHost
'                 Maintains for raw components identified by their
'                 component name the values HostFullName and ExpFileFullName.
' ---------------------------------------------------------------------------
Private Const VALUE_HOST_FULL_NAME = "HostFullName"

Public Property Get DAT_FILE() As String: DAT_FILE = mMe.CompManAddinPath & "\Hosts.dat":   End Property

Public Property Get FullName( _
                     Optional ByVal host_base_name As String) As String
    FullName = Value(vl_section:=host_base_name, vl_value_name:=VALUE_HOST_FULL_NAME)
End Property

Public Property Let FullName( _
                     Optional ByVal host_base_name As String, _
                              ByVal host_full_name As String)
    Value(vl_section:=host_base_name, vl_value_name:=VALUE_HOST_FULL_NAME) = host_full_name
End Property

Private Property Get Value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
    
    Value = mFile.Value(vl_file:=DAT_FILE _
                      , vl_section:=vl_section _
                      , vl_value_name:=vl_value_name _
                       )
End Property

Private Property Let Value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
' --------------------------------------------------
' Write the value (vl_value) named (vl_value_name)
' into the file RAWS_DAT_FILE.
' --------------------------------------------------
    
    mFile.Value(vl_file:=DAT_FILE _
              , vl_section:=vl_section _
              , vl_value_name:=vl_value_name _
               ) = vl_value

End Property

Public Function Exists(ByVal raw_host_base_name As String) As Boolean
    Exists = Hosts.Exists(Key:=raw_host_base_name)
End Function

Public Function Hosts() As Dictionary
    Set Hosts = mFile.SectionNames(sn_file:=DAT_FILE)
End Function

Public Sub Remove(ByVal raw_host_base_name As String)
    mFile.SectionsRemove sr_file:=DAT_FILE _
                       , sr_section_names:=raw_host_base_name
End Sub
