Attribute VB_Name = "mConfig"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mConfig
' Read/Write CompMan configuration properties from /to the Registry
' ----------------------------------------------------------------------------
Public Const CONFIG_BASE_KEY                As String = "HKCU\SOFTWARE\CompManVBP\BasicConfig\"
Private Const VNAME_ADDIN_IS_PAUSED         As String = "AddinIsPaused"
Private Const VNAME_FOLDER_ADDIN            As String = "FolderAddin"
Private Const VNAME_FOLDER_EXPORT           As String = "FolderExport"
Private Const VNAME_FOLDER_SERVICED         As String = "FolderServiced"

Public Property Get FolderAddin() As String
    If Exists(VNAME_FOLDER_ADDIN) _
    Then FolderAddin = Value(VNAME_FOLDER_ADDIN)
End Property

Public Property Let FolderAddin(ByVal s As String)
    Value(VNAME_FOLDER_ADDIN) = s
End Property

Public Property Get FolderServiced() As String
    If Exists(reg_value_name:=VNAME_FOLDER_SERVICED) _
    Then FolderServiced = Value(VNAME_FOLDER_SERVICED)
End Property

Public Property Let FolderServiced(ByVal s As String)
    Value(VNAME_FOLDER_SERVICED) = s
End Property

Public Property Get FolderExport() As String
    If Exists(reg_value_name:=VNAME_FOLDER_EXPORT) _
    Then FolderExport = Value(VNAME_FOLDER_EXPORT) _
    Else FolderExport = "source" ' The default
End Property

Public Property Let FolderExport(ByVal s As String)
    Value(VNAME_FOLDER_EXPORT) = s
End Property

Public Property Get CompManAddinIsPaused() As Boolean
    If Exists(reg_value_name:=VNAME_ADDIN_IS_PAUSED) _
    Then CompManAddinIsPaused = CBool(Value(VNAME_ADDIN_IS_PAUSED)) _
    Else CompManAddinIsPaused = False ' The default
End Property

Public Property Let CompManAddinIsPaused(ByVal b As Boolean)
    Value(VNAME_ADDIN_IS_PAUSED) = b
End Property

' ---------------------------------------------------------------------------
' Interfaces to mReg.Value Get/Let and NameExists
' ---------------------------------------------------------------------------
Private Property Get Value(Optional ByVal reg_value_name As String) As Variant
    Value = mReg.Value(reg_key:=CONFIG_BASE_KEY, reg_value_name:=reg_value_name)
End Property

Private Property Let Value(Optional ByVal reg_value_name As String, _
                                    ByVal reg_value As Variant)
    mReg.Value(reg_key:=CONFIG_BASE_KEY, reg_value_name:=reg_value_name) = reg_value
End Property

Private Function Exists(ByVal reg_value_name As String) As Boolean
    Exists = mReg.Exists(reg_key:=CONFIG_BASE_KEY, reg_value_name:=reg_value_name)
End Function

