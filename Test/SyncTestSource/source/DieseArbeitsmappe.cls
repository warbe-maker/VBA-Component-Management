VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ----------------------------------------------------------------------------
' Data Module wbSyncTest (Workbook)
'
' W. Rauschenberger, Berlin Dec 2022
' ----------------------------------------------------------------------------

Private Sub Workbook_Open()
    '~~ preparations at Workbook_Open in the Sync-Source- and the Sync-Target-Workbook
    wsTestNamesSync.TestNamesSetup
    
    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Serviced-Folder' path and either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook.
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED
    
    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Synchronize-Folder' path, either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook and the corresponding 'Synchronization-Source-
    '~~ Workbook' is available (must not be open!) in the 'Serviced-Folder' path.
    mCompManClient.CompManService mCompManClient.SRVC_SYNCHRONIZE
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Serviced-Folder' path and either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook.
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED

End Sub
