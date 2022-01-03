Attribute VB_Name = "mConfig"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mConfig
' Read/Write CompMan configuration properties from /to the Registry
' ----------------------------------------------------------------------------
Private Const VNAME_ADDIN_FOLDER            As String = "HKCU\CompMan\BasicConfig\AddinFolder"
Private Const VNAME_SERVICED_ROOT_FOLDER    As String = "HKCU\CompMan\BasicConfig\ServicedRootFolder"
Private Const VNAME_EXPORT_FOLDER           As String = "HKCU\CompMan\BasicConfig\ExportFolder"
Private Const VNAME_ADDIN_IS_PAUSED         As String = "HKCU\CompMan\BasicConfig\AddinIsPaused"

Public Property Get CompManAddinFolder() As String
    If mReg.Exists(VNAME_ADDIN_FOLDER) _
    Then CompManAddinFolder = mReg.Value(VNAME_ADDIN_FOLDER)
End Property

Public Property Let CompManAddinFolder(ByVal s As String)
    mReg.Value(VNAME_ADDIN_FOLDER) = s
End Property

Public Property Get CompManServicedRootFolder() As String
    If mReg.Exists(VNAME_SERVICED_ROOT_FOLDER) _
    Then CompManServicedRootFolder = mReg.Value(VNAME_SERVICED_ROOT_FOLDER)
End Property

Public Property Let CompManServicedRootFolder(ByVal s As String)
    mReg.Value(VNAME_SERVICED_ROOT_FOLDER) = s
End Property

Public Property Get CompManExportFolder() As String
    If mReg.Exists(VNAME_EXPORT_FOLDER) _
    Then CompManExportFolder = mReg.Value(VNAME_EXPORT_FOLDER) _
    Else CompManExportFolder = "source" ' The default
End Property

Public Property Let CompManExportFolder(ByVal s As String)
    mReg.Value(VNAME_EXPORT_FOLDER) = s
End Property

Public Property Get CompManAddinIsPaused() As Boolean
    If mReg.Exists(VNAME_ADDIN_IS_PAUSED) _
    Then CompManAddinIsPaused = CBool(mReg.Value(VNAME_ADDIN_IS_PAUSED)) _
    Else CompManAddinIsPaused = False ' The default
End Property

Public Property Let CompManAddinIsPaused(ByVal b As Boolean)
    mReg.Value(VNAME_ADDIN_IS_PAUSED) = CInt(b)
End Property


