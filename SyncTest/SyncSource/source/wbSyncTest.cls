VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbSyncTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ----------------------------------------------------------------------------
' CompManSyncTest Workbook (module wbSyncTest)
'
' While a synchronysation has not been performed and this Workbook resides in
' CompMan's 'Serviced-Folder' path it is the 'Synchronization-Source-Workbook'.
' When this Workbook is opened from within CompMan's 'Synchronization-Folder'
' path it will automatically invoke the synchronization service.
'
' W. Rauschenberger, Berlin May 2022
' ----------------------------------------------------------------------------

Private Sub Workbook_Open()
    
    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Serviced-Folder' path and either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook.
    mCompManClient.CompManService "UpdateOutdatedCommonComponents"
    
    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Synchronize-Folder' path, either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook and the corresponding 'Synchronization-Source-
    '~~ Workbook' is available (must not be open!) in the 'Serviced-Folder' path.
    mCompManClient.CompManService "Synchronize"

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Serviced-Folder' path and either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook.
    mCompManClient.CompManService "ExportChangedComponents"

End Sub

