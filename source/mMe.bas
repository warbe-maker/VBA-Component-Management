Attribute VB_Name = "mMe"
Option Explicit
Option Private Module
' ---------------------------------------------------------------------------
' Standard Module mMe   Services for the self management like the relation
'                       between the Component Management AddIn instance
'                       and the development instance.
'
' Public services:
' - CfgAsserted                    Returns True when the required properties
'                                  (the paths) are configured and exist
' - ControlItemRenewAdd            Adds a 'Renew___AddIn' control item to the
'                                  'Add-Ins' poupup menu.
' - ControlItemRenewRemove         Removes the 'Renew___AddIn' control item
'                                  from the 'Add-Ins' poupup menu when the
'                                  Workbook is closed.
' - Renew___AddIn                     Called via the 'Renew___AddIn' control item
'                                  in the 'Add-Ins' popup menu or executed
'                                  via the corresponding Command Button at
'                                  the 'Manage CompMan Add-in' Worksheet.
' - UpdateOutdatedCommonComponents Updates any outdated Used Common
'                                  Component by means of the Raw's Export
'                                  File which had been saved to the Common
'                                  Components folder with the last export.
'                                  Because a Workbook cannot update its own
'                                  components the Development instance
'                                  Workbook requires an active 'Add-in
'                                  Instance' to get its outdated Used
'                                  Common Components updated. For any other
'                                  Workbook the service can be provided by
'                                  the open 'Development Instance'.
' - FolderServicedIsValid          CompMan services are only applied for
'                                  Workbooks which are located in the
'                                  configured 'Serviced Folder' - which
'                                  prevents productive Workbooks are bothered
' Uses Common Components:
' - mFso                   Get/Let PrivateProperty value service
' - mWrkbk                  GetOpen and Opened service
' - mMsg                    Dsply, Box, and Buttons service used by the
'                           RenewAddin,  Renew_1_ConfirmConfig service
' - mErH                    Common VBA Error Handling
'
' W. Rauschenberger, Berlin Nov 2020
' ---------------------------------------------------------------------------
Public Const COMPMAN_ADMIN_FOLDER_NAME      As String = "\CompManAdmin\"
Public Const FOLDER_SERVICED                As String = "Serviced-By-CompMan folder"
Private Const DEVLP_WORKBOOK_EXTENSION      As String = "xlsb"  ' Extension may depend on Excel version

Private wbDevlp                 As Workbook
Private wbkSource               As Workbook                     ' This development instance as the renew source
Private wbkTarget               As Workbook                     ' The Add-in instance as renew target
Private bSucceeded              As Boolean
Private bAllRemoved             As Boolean
Private dctAddInRefs            As Dictionary
Private lRenewStep              As Long
Private sRenewAction            As String
Private bRenewTerminatedByUser  As Boolean

Private Property Get ADDIN_FORMAT() As XlFileFormat ' = ... needs adjustment when the above is changed
    ADDIN_FORMAT = xlOpenXMLAddIn
End Property

Public Property Get AutoOpenShortCutCompManWbk()
    AutoOpenShortCutCompManWbk = Environ$("APPDATA") & "\Microsoft\Excel\XLSTART\CompManWbk.lnk"
End Property

Public Property Get DevInstncFullName() As String
    Dim fso As New FileSystemObject
    DevInstncFullName = wsConfig.FolderDevAndTest & DBSLASH _
                          & fso.GetBaseName(DevInstncName) & DBSLASH _
                          & DevInstncName
End Property

Public Property Get DevInstncName() As String
    With New FileSystemObject
        DevInstncName = .GetBaseName(ThisWorkbook.FullName) & "." & DEVLP_WORKBOOK_EXTENSION
    End With
End Property

Private Property Get DEVLP_FORMAT() As XlFileFormat  ' = .xlsb ! may require adjustment when the above is changed
    DEVLP_FORMAT = xlExcel12
End Property

Public Property Get IsAddinInstnc() As Boolean
    With New FileSystemObject
        IsAddinInstnc = .GetExtensionName(ThisWorkbook.FullName) = "xlam"
    End With
End Property

Public Property Get IsDevInstnc() As Boolean
    With New FileSystemObject
        IsDevInstnc = .GetExtensionName(ThisWorkbook.FullName) = "xlsb"
    End With
End Property

Public Property Get RenewAction(Optional ByVal la_last As Boolean = False) As String
    RenewAction = sRenewAction
End Property

Public Property Let RenewAction(Optional ByVal la_last As Boolean = False, _
                                         ByVal la_action As String)
    lRenewStep = lRenewStep + 1
    sRenewAction = la_action
End Property

Public Property Let RenewMonitorResult(Optional ByVal la_result_text As String = vbNullString, _
                                       Optional ByVal la_last As Boolean = False, _
                                                ByVal la_result As String)
    wsConfig.MonitorRenewStep rn_result:=la_result _
                            , rn_action:=la_result_text _
                            , rn_last_step:=la_last
End Property

Public Property Get RenewStep() As Long:        RenewStep = lRenewStep: End Property

Public Property Get RenewTerminatedByUser() As Boolean
    RenewTerminatedByUser = bRenewTerminatedByUser
End Property

Public Property Let RenewTerminatedByUser(ByVal b As Boolean)
    bRenewTerminatedByUser = b
End Property

Public Property Get xlAddInFormat() As Long:            xlAddInFormat = ADDIN_FORMAT:                                       End Property

Public Property Get xlDevlpFormat() As Long:            xlDevlpFormat = DEVLP_FORMAT:                                       End Property

Public Sub CompManConfig()
' ----------------------------------------------------------------------------
' Invoked by the corresponding button in the wsAddin (Manage CompMan addin)
' Worksheet.
' ----------------------------------------------------------------------------
    mMe.Config cfg_silent:=False, cfg_addin:=False, cfg_sync:=False
End Sub

Public Function Config(Optional ByVal cfg_silent As Boolean = False, _
                       Optional ByVal cfg_addin As Boolean = False, _
                       Optional ByVal cfg_sync As Boolean = False) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the basic configuration and any additional requested
' configurations (cfg_addin or cfg_sync) are valid.
' - Basic configuration: Serviced-Root-Folder and Export-Folder are valid.
' - Add-in configuration: The Add-in-Folder is valid
' - Sync configuration:  The Synchronization-Folder is valid
' When a silent configuration check is requested (cfg_silent = True) the
' configuration dialog is only displayed when something required is still
' invalid or yet not configured.
' ----------------------------------------------------------------------------
    Const PROC = "Config"

    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim cfgAddin    As Boolean
    Dim cfgSync     As Boolean
    Dim cfgExpUpdt  As Boolean
    
    If cfg_addin Then Config = wsConfig.FolderDevAndTestIsValid _
                           And wsConfig.FolderExportIsValid _
                           And wsConfig.FolderAddInIsValid
    
    If cfg_sync Then Config = wsConfig.FolderDevAndTestIsValid _
                          And wsConfig.FolderExportIsValid _
                          And wsConfig.FolderSyncTargetIsValid
    
    If cfg_silent Then
        If Not wsConfig.FolderDevAndTestIsValid _
        Or Not wsConfig.FolderExportIsValid Then
            wsConfig.ConfigInfo = "At least one essential configuration is still missing!"
            Config = False
        Else
            Config = True
            wsConfig.ConfigInfo = vbNullString
        End If
        GoTo xt
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub DevInstncWorkbookClose()
    Const PROC = "DevInstncWorkbookClose"
    
    On Error GoTo eh

    mMe.RenewAction = "Close 'Development-Instance-Workbook' (" & DevInstncName & ")"
    
    On Error Resume Next
    wbDevlp.Activate
    wbDevlp.Close False
    If Err.Number <> 0 Then
        mMe.RenewMonitorResult("Closing the 'Development-Instance-Workbook' (" & DevInstncName & ") failed with:" & vbLf & _
                               "(" & Err.Description & ")" _
                              ) = "Failed"
    Else
        mMe.RenewMonitorResult() = "Passed"
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub DevInstncWorkbookDelete()
    Const PROC = "DevInstncWorkbookDelete"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    mMe.RenewAction = "Delete the 'Development-Instance-Workbook' (" & DevInstncName & ")"
    
    With fso
        If .FileExists(DevInstncFullName) Then
            On Error Resume Next
            .DeleteFile DevInstncFullName
            If Err.Number = 0 Then
                mMe.RenewMonitorResult() = "Passed"
            Else
                mMe.RenewMonitorResult("Deleting the 'Development-Instance-Workbook' (" & DevInstncName & ") failed with:" & vbLf & _
                                       "(" & Err.Description & ")" _
                                      ) = "Failed"
            End If
        Else
            mMe.RenewMonitorResult() = "Passed"
        End If
    End With

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function DevInstncWorkbookExists() As Boolean
Dim fso As New FileSystemObject
    DevInstncWorkbookExists = fso.FileExists(DevInstncFullName)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Private Sub RenewFinalResult(ByVal r_fs As Boolean)
    If r_fs Then
        mMe.RenewAction = "Successful!"
        RenewMonitorResult(la_last:=True _
                         , la_result_text:=mBasic.Spaced("Successful!") & "   The Add-in '" & mAddin.WbkName & "' has been renewed by the development instance '" & DevInstncName & "'" _
                          ) = "Passed"
    Else
        mMe.RenewAction = "Not Successful!"
        mMe.RenewMonitorResult(la_last:=True _
                             , la_result_text:="Renewing the Add-in '" & mAddin.WbkName & "' by the development instance '" & DevInstncName & "'  " & mBasic.Spaced("failed!") _
                              ) = "Failed"
    End If
End Sub

Private Function Renew_01_ConfirmConfig() As Boolean
    mMe.RenewAction = "Assert 'CompMan's Basic Configuration'"
    Renew_01_ConfirmConfig = mMe.Config(cfg_addin:=True)
    If Renew_01_ConfirmConfig _
    Then mMe.RenewMonitorResult = "Passed" _
    Else mMe.RenewMonitorResult = "Failed"
End Function

Private Function Renew_02_DevInstnc() As Boolean
    mMe.RenewAction = "Assert this 'Setup/Renew' service is executed from the 'Development-Instance-Workbook'"
    Renew_02_DevInstnc = IsDevInstnc()
    If Not Renew_02_DevInstnc Then
        mMe.RenewMonitorResult("The 'Renew___AddIn' service had not been executed from within the development instance Workbook (" & DevInstncName & ")!" _
                              ) = "Failed"
    Else
        mMe.RenewMonitorResult = "Passed"
    End If
End Function

Private Sub Renew_03_SaveAndRemoveAddInReferences()
' ----------------------------------------------------------------------------
' - Allows the user to close any open Workbook which refers to the Add-in,
'   which definitly hinders the Add-in from being re-newed.
' - Returns TRUE when the user closed all open Workbboks.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_03_SaveAndRemoveAddInReferences"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim wb          As Workbook
    Dim ref         As Reference
    Dim sWbs        As String
    Dim bOneRemoved As Boolean
    Dim fso         As New FileSystemObject

    mMe.RenewAction = "Save and remove references to the Add-in from open Workbooks"
    mAddin.ReferencesRemove dctAddInRefs, sWbs, bOneRemoved, bAllRemoved
    If bOneRemoved Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewMonitorResult() = "Passed"
    Else
        mMe.RenewMonitorResult(sRenewAction & vbLf & "None of the open Workbook's VBProject had a 'Reference' to the 'CompMan Add-in'" _
                          ) = "Passed"
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_04_DevInstncWorkbookSave()
    Const PROC = "Renew_04_DevInstncWorkbookSave"
    
    On Error GoTo eh
    mMe.RenewAction = "Save the 'Development-Instance-Workbook' (" & DevInstncName & ")"
    
    Set wbkSource = Application.Workbooks(DevInstncName)
    wbkSource.Save
    wbkSource.Activate
    mMe.RenewMonitorResult() = "Passed"

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_05_Set_IsAddin_ToFalse()
    Const PROC = "Renew_05_Set_IsAddin_ToFalse"
    
    On Error GoTo eh
    Dim wbk As Workbook
    mMe.RenewAction = "Set the 'IsAddin' property of the 'CompMan Add-in' to FALSE"
    If mAddin.SetIsAddinToFalse Then
        mMe.RenewMonitorResult() = "Passed"
    Else
        mMe.RenewMonitorResult("CompMan's 'Add-in Instance was not open or the 'IsAddin' property was already set to FALSE" _
                                  ) = "Passed"
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Renew_06_CloseCompManAddinWorkbook() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the Add-in has successfully been closed.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_06_CloseCompManAddinWorkbook"
    
    Dim fso         As New FileSystemObject
    Dim sErrDesc    As String
    
    mMe.RenewAction = "Close the 'CompMan Add-in'"
    wbkSource.Activate
    Renew_06_CloseCompManAddinWorkbook = mAddin.WbkClose(sErrDesc)
    If Not Renew_06_CloseCompManAddinWorkbook Then
        mMe.RenewMonitorResult("Closing the 'CompMan Add-in' (" & mAddin.WbkName & ") failed with:" & vbLf & _
                               "(" & sErrDesc & ")" _
                              ) = "Failed"
    Else
        mMe.RenewMonitorResult = "Passed"
    End If
    
xt: Set fso = Nothing
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Renew_07_DeleteAddInInstanceWorkbook() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the Add-in instance Workbbook has been deleted
' ----------------------------------------------------------------------------
    Const PROC = "Renew_07_DeleteAddInInstanceWorkbook"
    
    On Error GoTo eh
    mMe.RenewAction = "Delete the 'CompMan Add-in' Workbook' (" & mAddin.WbkName & ")"
    
    With New FileSystemObject
        If .FileExists(WbkFullName) Then
            On Error Resume Next
            mAddin.WbkRemove WbkFullName
            Renew_07_DeleteAddInInstanceWorkbook = Err.Number = 0
            If Renew_07_DeleteAddInInstanceWorkbook Then
                mMe.RenewMonitorResult = "Passed"
            Else
                mMe.RenewMonitorResult("Deleting the 'CompMan Add-in' (" & mAddin.WbkName & ") failed with:" & vbLf & _
                                       "(" & Err.Description & ")" _
                                      ) = "Failed"
            End If
        Else
            Renew_07_DeleteAddInInstanceWorkbook = True
            mMe.RenewMonitorResult = "Passed"
        End If
    End With
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Renew_08_SaveDevInstncWorkbookAsAddin() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the development instance Workbook has successfully saved
' as Add-in.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_08_SaveDevInstncWorkbookAsAddin"
    
    On Error GoTo eh
    mMe.RenewAction = "Save the 'Development-Instance-Workbook' (" & DevInstncName & ") as 'CompMan Add-in' (" & mAddin.WbkName & ")"
    
    With Application
        If Not mAddin.Exists Then
            '~~ At this point the Add-in must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbkSource.SaveAs WbkFullName, FileFormat:=xlAddInFormat
            Renew_08_SaveDevInstncWorkbookAsAddin = Err.Number = 0
            If Not Renew_08_SaveDevInstncWorkbookAsAddin Then
                mMe.RenewMonitorResult("Save Development instance as Add-in instance  " & mBasic.Spaced("failed!") _
                                      ) = "Failed"
            Else
                mMe.RenewMonitorResult() = "Passed"
            End If
            .EnableEvents = True
        Else ' file still exists
            mMe.RenewMonitorResult("Setup/Renew of the 'CompMan Add-in' as copy of the 'Development-Instance-Workbook'  " & mBasic.Spaced("failed!") _
                              ) = "Failed"
        End If
    End With
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Renew_09_OpenAddinInstncWorkbook() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the Add-in instance Workbook has successfully been opened.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_09_OpenAddinInstncWorkbook"
    
    On Error GoTo eh
    Dim wb              As Workbook
    Dim sBaseAddinName  As String
    Dim sBaseDevName    As String
    
    If Not mAddin.IsOpen Then
        If mAddin.Exists Then
            mMe.RenewAction = "Re-open the 'CompMan Add-in' (" & mAddin.WbkName & ")"
            On Error Resume Next
            Set wb = Application.Workbooks.Open(WbkFullName)
            If Err.Number = 0 Then
                With New FileSystemObject
                    sBaseAddinName = .GetBaseName(wb.name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.name)
                    wb.VBProject.name = sBaseAddinName
                End With
                mMe.RenewMonitorResult() = "Passed"
                Renew_09_OpenAddinInstncWorkbook = True
            Else
                mMe.RenewMonitorResult("(Re)opening the 'CompMan Add-in' (" & mAddin.WbkName & ") failed with:" & vbLf & _
                                       "(" & Err.Description & ")" _
                                      ) = "Failed"
            End If
        End If
    End If

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub Renew_11_SetupAutoOpen()
    Const PROC = "Renew_11_SetupAutoOpen"
    
    On Error GoTo eh
    mMe.RenewAction = "Setup/maintain auto-open for the 'CompMan Add-in'"
    wsConfig.AutoOpenCompManAddinSetup
    mMe.RenewMonitorResult() = "Passed"
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_10_RestoreReferencesToAddIn()
    Const PROC = "Renew_10_RestoreReferencesToAddIn"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim wb              As Workbook
    Dim sWbs            As String
    Dim bOneRestored    As Boolean
    
    mMe.RenewAction = "Restore all saved 'References' to the 'CompMan Add-in' in open Workbooks"
    
    For Each v In dctAddInRefs
        Set wb = v
        wb.VBProject.References.AddFromFile WbkFullName
        sWbs = wb.name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewMonitorResult() = "Passed"
    Else
        mMe.RenewMonitorResult(sRenewAction & vbLf & "Restoring 'References' did not find any saved to restore" _
                              ) = "Passed"
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Renew___AddIn()
' ------------------------------------------------------------------------------
' Called via the Command Button in the "Manage CompMan Add-in" sheet.
' Renews the code of the Add-in instance of this Workbook with this Workbook's
' code by displaying a detailed result of the whole Renew___AddIn process.
' Note: It cannot be avoided that this procedure is available also in the Add-in
'       instance. However, its execution is limited to this Workbook's
'       development instance.
' ------------------------------------------------------------------------------
    Const PROC = "Renew___AddIn"
    
    On Error GoTo eh
    lRenewStep = 0

    Application.EnableEvents = False
    bSucceeded = False
                            
    '~~ Get the CompMan base configuration confirmed or changed
    If Not Renew_01_ConfirmConfig Then GoTo xt
                         
    '~~ Assert the Renew service is performed from within the development instance Workbbok
    '~~ Note that the distinction of the instances requires the above basic configuration confirmed
    If Not Renew_02_DevInstnc Then GoTo xt
    
    '~~ Assert no Workbooks are open referring to the Add-in
    Renew_03_SaveAndRemoveAddInReferences
    If Not bAllRemoved Then GoTo xt

    '~~ Assure the current version of the Add-in's development instance has been saved
    '~~ Note: Unconditionally saving the Workbook does an incredible trick:
    '~~       The un-unstalled and IsAddin=False Workbook is released from the Application
    '~~       and no longer considered "used"
    Renew_04_DevInstncWorkbookSave
    wbkSource.Activate
          
    '~~ Attempt to turn Add-in to "IsAddin=False" in order to uninstall and subsequently close it
    Renew_05_Set_IsAddin_ToFalse
    If Not Renew_06_CloseCompManAddinWorkbook Then GoTo xt

    
    '~~ Attempt to delete the Add-in Workbook file
    If Not Renew_07_DeleteAddInInstanceWorkbook Then GoTo xt
        
    '~~ Attempt to save the development instance as Add-in
    If Not Renew_08_SaveDevInstncWorkbookAsAddin Then GoTo xt
    
    '~~ Saving the development instance as Add-in may also open the Add-in.
    '~~ So if not already open it is re-opened and thus re-activated
    If Not Renew_09_OpenAddinInstncWorkbook Then GoTo xt
        
    '~~ Re-instate references to the Add-in which had been removed
    Renew_10_RestoreReferencesToAddIn
    Renew_11_SetupAutoOpen
    
    bSucceeded = True
    wsConfig.Activate
    
xt: RenewFinalResult bSucceeded
    wsConfig.CurrentStatus
    Application.EnableEvents = True
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SaveAddinInstncWorkbookAsDevlp()
    Const PROC = "SaveAddinInstncWorkbookAsDevlp"

    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    With Application
        If Not DevInstncWorkbookExists Then
            '~~ At this point the Development instance Workbook must no longer exist at its location
            .EnableEvents = False
            mMe.RenewAction = "Save the 'CompMan Add-in' (" & mAddin.WbkName & ") as 'Development-Instance-Workbook' (" & DevInstncName & ")"
            
            On Error Resume Next
            wbAddIn.SaveAs DevInstncFullName, FileFormat:=xlDevlpFormat, ReadOnlyRecommended:=False
            
            If Not mCompMan.WbkIsOpen(io_name:=DevInstncName) _
            Then Stop _
            Else wbDevlp.VBProject.name = fso.GetBaseName(DevInstncName)
            
            If Err.Number <> 0 Then
                mMe.RenewMonitorResult("Saving the 'CompMan Add-in' (" & mAddin.WbkName & ") as 'Development-Instance-Workbook' (" & DevInstncName & ")   " & mBasic.Spaced("failed!") _
                                      ) = "Failed"
            Else
                mMe.RenewMonitorResult() = "Passed"
            End If
            .EnableEvents = True
        Else ' file still exists
            mMe.RenewMonitorResult("Saving the 'CompMan Add-in' (" & mAddin.WbkName & ") as 'Development-Instance-Workbook' (" & DevInstncName & ")  " & mBasic.Spaced("failed!") _
                                  ) = "Failed"
        End If
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


