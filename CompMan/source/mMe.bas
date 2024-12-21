Attribute VB_Name = "mMe"
Option Explicit
Option Private Module
' ---------------------------------------------------------------------------
' Standard Module mMe: Services for the self management like the relation
' ==================== between the Component Management AddIn instance
' and the development instance.
'
' Public services:
' ----------------
' CfgAsserted                    Returns True when the required properties
'                                (the paths) are configured and exist
' ControlItemRenewAdd            Adds a 'Renew___AddIn' control item to the
'                                'Add-Ins' poupup menu.
' ControlItemRenewRemove         Removes the 'Renew___AddIn' control item
'                                from the 'Add-Ins' poupup menu when the
'                                Workbook is closed.
' ErrMsg                         Used throughout the VB-Project, keeps the
'                                usage of the components mErH, mMsg/fMsg
'                                optional through Conditional Compile
'                                Arguments.
' Renew___AddIn                  Called via the 'Renew___AddIn' control
'                                item in the 'Add-Ins' popup menu or
'                                executed via the corresponding Command
'                                Button at the 'Manage CompMan Add-in'
'                                Worksheet.
' UpdateOutdatedCommonComponents Updates any outdated Used Common
'                                Component by means of the Raw's Export
'                                File which had been saved to the Common
'                                Components folder with the last export.
'                                Because a Workbook cannot update its own
'                                components the Development instance
'                                Workbook requires an active 'Add-in
'                                Instance' to get its outdated Used
'                                Common Components updated. For any other
'                                Workbook the service can be provided by
'                                the open 'Development Instance'.
' FolderServicedIsValid          CompMan services are only applied for
'                                Workbooks which are located in the
'                                configured 'Serviced Folder' - which
'                                prevents productive Workbooks are bothered
' Uses Common Components:
' -----------------------
' mFso                           Get/Let PrivateProperty value service
' mMsg                           Dsply, Box, and Buttons service used by the
'                                RenewAddin,  Renew_1_ConfirmConfig service
' mErH                           Common VBA Error Handling
'
' W. Rauschenberger, Berlin Jan 2024
' ---------------------------------------------------------------------------
Public BaseName                         As String

Private Const DEVLP_WORKBOOK_EXTENSION  As String = "xlsb"              ' Extension may depend on Excel version
Private bAllRemoved                     As Boolean
Private bRenewTerminatedByUser          As Boolean
Private bSucceeded                      As Boolean
Private dctAddInRefs                    As Dictionary
Private Extension                       As String
Private lRenewStep                      As Long
Private sRenewAction                    As String
Private wbDevlp                         As Workbook
Private wbkSource                       As Workbook                     ' This development instance as the renew source

Public Function AssertedServicingEnabled(ByVal a_hosted As String) As Boolean
' ---------------------------------------------------------------------------
' When TRUE is returned either the opened Workbook is the Addin instance or
' this Workbook is ready for servicing, which means:
' - The required Office version is installed
' - The required files and folder structure is set up
' - WinMerge is installed
' Note: In case the Workbook is opened at a location the required files and
'       folder structure is not setup, e.g. after download, it will be setup
'       by the way.
' ---------------------------------------------------------------------------
    Const PROC = "AssertedServicingEnabled"
    
    On Error GoTo eh
    
    If Trc Is Nothing Then Set Trc = New clsTrc
    BaseName = FSo.GetBaseName(ThisWorkbook.Name)
    Extension = FSo.GetExtensionName(ThisWorkbook.Name)
    mEnvironment.Provide False

    Select Case True
        Case mMe.IsAddinInstnc
            AssertedServicingEnabled = True
        Case Not AssertedOfficeVersion
            Debug.Print ErrSrc(PROC) & ": Office version not asserted!"
        Case Not AssertedFilesAndFldrsStructure(a_hosted)
            Debug.Print ErrSrc(PROC) & ": Files and folders structure not asserted!"
        Case Not AssertedWinMerge
            Debug.Print ErrSrc(PROC) & ": WinMerge not installed!"
        Case Else
            AssertedServicingEnabled = True
    End Select
        
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AssertedFilesAndFldrsStructure(ByVal a_hosted As String) As Boolean
' ---------------------------------------------------------------------------
' Performed with each open. Ensures that the Workbook is running from within
' the required files and folders structure.
' ---------------------------------------------------------------------------
    Const PROC = "AssertedFilesAndFldrsStructure"
    
    On Error GoTo eh
    Dim BttnAbort   As String
    Dim BttnGoAhead As String
    Dim sWrkbkOpnd  As String
    Dim sCompManFldr As String
    
    Application.EnableEvents = False
    Set Servicing = New clsServicing
        
    If Not FSo.FolderExists(mEnvironment.CommCompsPath) Then
        '~~ The CompMan Workbook has been opened the very first time at this location.
        '~~ All default folders and files environment is now setup - provided the user confirms it.
        BttnGoAhead = "Ok!" & vbLf & vbLf & _
                      "Go ahead and set it up"
        BttnAbort = "Abort!" & vbLf & vbLf & _
                    "Serviced root folder needs to be changed"
                     
        If mConfig.DefaultEnvDisplay(BttnGoAhead, BttnAbort) = BttnGoAhead Then
            mConfig.SelfSetupDefaultEnvironment sCompManFldr
            '~~ Save the Workbook to its dedicated folder withing the servicing environment
            sWrkbkOpnd = ThisWorkbook.FullName
            ThisWorkbook.SaveAs sCompManFldr & "\" & ThisWorkbook.Name
            mWinMergeIni.Setup
            AssertedFilesAndFldrsStructure = True
            DoEvents
            On Error Resume Next
            mEnvironment.Provide True
            Set ConfigLocal = New clsConfigLocal
            mConfig.SelfSetupPublishHostedCommonComponents (a_hosted)
            mConfig.SetupConfirmed
            FSo.DeleteFile sWrkbkOpnd
        End If
    Else
        '~~  The existing folders indicate that CompMan's default environment is already set up
        If mConfig.ServicedRootFolderNameCurrent <> wsConfig.FolderCompManServicedRoot _
        Then mConfig.Adjust
    End If ' Config exists
    
    If wsConfig.Verified Then
        '~~ Nothing had been changed while the Workbook was closed
        AssertedFilesAndFldrsStructure = True
    Else
        '~~ The configuration is no loger valid. This may be the case when the CompMan root folder
        '~~ has been renamed or moved to another location
        AssertedFilesAndFldrsStructure = False
        wsConfig.Activate
    End If
                        
xt: Set Servicing = Nothing
    Application.EnableEvents = True
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AssertedOfficeVersion() As Boolean
' ---------------------------------------------------------------------------
'
' ---------------------------------------------------------------------------
    Dim lVersion As Long
    
    lVersion = CLng(Split(Application.Version, ".")(0))
    AssertedOfficeVersion = True
    If lVersion < 14 Then
        AssertedOfficeVersion = MsgBox(Title:="Expected minimum Excel version not installed!" _
              , Prompt:="'CompMan' had been developed and tested with an Excel version 14.0 and higher. " & _
                        "Because the available Excel version is '" & Application.Version & "' CompMan " & _
                        "may not work properly." & vbLf & vbLf & _
                        "Yes: Go ahead trying if it will work - and if it does " & vbLf & _
                        "       an eMail to warb@tutanota.com " & vbLf & _
                        "       would be very much appreaciated. :-))" & vbLf & vbLf & _
                        "No: Give up. 'CompMan' will nor be able to provide any of its services." _
              , Buttons:=vbYesNo) = vbYes
    End If

End Function

Private Function AssertedWinMerge() As Boolean
    
    Dim Msg     As mMsg.udtMsg
    Dim Title   As String
    
    Title = "WinMerge is not installed!"
    AssertedWinMerge = mCompMan.WinMergeIsInstalled
    
    If Not AssertedWinMerge Then
        With Msg.Section(1)
            .Text.Text = "WinMerge is used by CompMan to display code changes in case an " & _
                         "update for a used Common Component is due. When WinMerge is not " & _
                         "installed the service """ & mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY & """ will be denied! " & _
                         "However, the services """ & mCompManClient.SRVC_EXPORT_CHANGED_DSPLY & """ and " & _
                         """" & mCompManClient.SRVC_SYNCHRONIZE_DSPLY & """ will be provides when requested."
        End With
        With Msg.Section(2).Label
            .FontColor = rgbBlue
            .Text = "Download and install the desired language version of WinMerge"
            .OnClickAction = "https://winmerge.org/downloads/"
        End With
        With Msg.Section(3)
            With .Label
                .FontColor = rgbBlack
                .Text = "Please note:"
            End With
            .Text.Text = "When continued without having downloaded and installed WinMerge CompMan will not be " & _
                         "able to provide any service but re-displays this message when re-opened!"
        End With
        mMsg.Dsply Title, Msg, vbOKOnly
    
        AssertedWinMerge = mCompMan.WinMergeIsInstalled     ' May have been downloaded and installed along with the displayed message
        If AssertedWinMerge Then
            If Not FSo.FileExists(mWinMergeIni.WinMergeIniFullName) Then
                mWinMergeIni.Setup ' ensures that the required options are established
            End If
        End If
    End If
    
End Function

Private Property Get ADDIN_FORMAT() As XlFileFormat ' = ... needs adjustment when the above is changed
    ADDIN_FORMAT = xlOpenXMLAddIn
End Property

Public Property Get AutoOpenShortCutCompManWbk()
    AutoOpenShortCutCompManWbk = Environ$("APPDATA") & "\Microsoft\Excel\XLSTART\CompManWbk.lnk"
End Property

Public Property Get DevInstncFullName() As String
    DevInstncFullName = wsConfig.FolderCompManServicedRoot _
                      & DBSLASH _
                      & FSo.GetBaseName(DevInstncName) & DBSLASH _
                      & DevInstncName
End Property

Private Property Get DevInstncName() As String
    With New FileSystemObject
        DevInstncName = .GetBaseName(ThisWorkbook.FullName) & "." & DEVLP_WORKBOOK_EXTENSION
    End With
End Property

Public Property Get IsAddinInstnc() As Boolean
    IsAddinInstnc = Extension = "xlam"
End Property

Public Property Get IsDevInstnc() As Boolean
    With New FileSystemObject
        IsDevInstnc = .GetExtensionName(ThisWorkbook.FullName) = "xlsb"
    End With
End Property

Public Property Get RenewAction() As String
    RenewAction = sRenewAction
End Property

Public Property Let RenewAction(ByVal la_action As String)
    lRenewStep = lRenewStep + 1
    sRenewAction = la_action
End Property

Private Property Let RenewMonitorResult(Optional ByVal la_last As Boolean = False, _
                                                ByVal la_result As String)
    wsConfig.MonitorRenewStep rn_result:=la_result _
                            , rn_last_step:=la_last
End Property

Public Property Get RenewStep() As Long:        RenewStep = lRenewStep: End Property

Public Property Get RenewTerminatedByUser() As Boolean
    RenewTerminatedByUser = bRenewTerminatedByUser
End Property

Public Property Let RenewTerminatedByUser(ByVal b As Boolean)
    bRenewTerminatedByUser = b
End Property

Private Function Config(Optional ByVal cfg_silent As Boolean = False, _
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
    
    If cfg_addin Then Config = wsConfig.FolderCompManRootIsValid _
                           And wsConfig.FolderExportIsValid
    
    If cfg_sync Then Config = wsConfig.FolderCompManRootIsValid _
                          And wsConfig.FolderExportIsValid
    
    If cfg_silent Then
        If Not wsConfig.FolderCompManRootIsValid _
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
    
    mMe.RenewAction = "Delete the 'Development-Instance-Workbook' (" & DevInstncName & ")"
    
    With FSo
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

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Private Sub RenewFinalResult(ByVal r_fs As Boolean)
    If r_fs Then
        mMe.RenewAction = "Successful!"
        RenewMonitorResult(True) = "Passed"
    Else
        mMe.RenewAction = "Not Successful!"
        mMe.RenewMonitorResult(True) = "Failed"
    End If
End Sub

Private Function Renew_010_ConfirmConfig() As Boolean
    Const PROC = "Renew_010_ConfirmConfig"
    
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Assert 'CompMan's Basic Configuration'"
    Renew_010_ConfirmConfig = mMe.Config(cfg_addin:=True)
    If Renew_010_ConfirmConfig _
    Then mMe.RenewMonitorResult = "Passed" _
    Else mMe.RenewMonitorResult = "Failed"

xt: mBasic.EoP ErrSrc(PROC)
End Function

Private Function Renew_020_DevInstnc() As Boolean
    Const PROC = "Renew_020_DevInstnc"
    
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Assert this 'Setup/Renew' service is executed from the 'Development-Instance-Workbook'"
    Renew_020_DevInstnc = IsDevInstnc()
    If Not Renew_020_DevInstnc Then
        mMe.RenewMonitorResult("The 'Renew___AddIn' service had not been executed from within the development instance Workbook (" & DevInstncName & ")!" _
                              ) = "Failed"
    Else
        mMe.RenewMonitorResult = "Passed"
    End If

xt: mBasic.EoP ErrSrc(PROC)
End Function

Private Sub Renew_030_SaveAndRemoveAddInReferences()
' ----------------------------------------------------------------------------
' - Allows the user to close any open Workbook which refers to the Add-in,
'   which definitly hinders the Add-in from being re-newed.
' - Returns TRUE when the user closed all open Workbboks.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_030_SaveAndRemoveAddInReferences"
    
    On Error GoTo eh
    Dim sWbs        As String
    Dim bOneRemoved As Boolean

    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Save and remove references to the Add-in from open Workbooks"
    mAddin.ReferencesRemove dctAddInRefs, sWbs, bOneRemoved, bAllRemoved
    If bOneRemoved Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewMonitorResult() = "Passed"
    Else
'        mMe.RenewMonitorResult(sRenewAction & vbLf & "None of the open Workbook's VBProject had a 'Reference' to the 'CompMan Add-in'" _
                          ) = "Passed"
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_040_DevInstncWorkbookSave()
    Const PROC = "Renew_040_DevInstncWorkbookSave"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Save the 'Development-Instance-Workbook' (" & DevInstncName & ")"
    
    Set wbkSource = Application.Workbooks(DevInstncName)
    With wbkSource
        If Not .Saved Then .Save
        .Activate
    End With
    mMe.RenewMonitorResult() = "Passed"

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_050_Set_IsAddin_ToFalse()
    Const PROC = "Renew_050_Set_IsAddin_ToFalse"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Set the 'IsAddin' property of the 'CompMan Add-in' to FALSE"
    If mAddin.Set_IsAddin_ToFalse _
    Then mMe.RenewMonitorResult() = "Passed" _
    Else mMe.RenewMonitorResult("CompMan's 'Add-in Instance was not open or the 'IsAddin' property was already set to FALSE" _
                               ) = "Passed"
     
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Renew_060_CloseCompManAddinWorkbook() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the Add-in has successfully been closed.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_060_CloseCompManAddinWorkbook"
    
    Dim sErrDesc    As String
    
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Close the 'CompMan Add-in'"
    wbkSource.Activate
    Renew_060_CloseCompManAddinWorkbook = mAddin.WbkClose(sErrDesc)
    If Not Renew_060_CloseCompManAddinWorkbook Then
        mMe.RenewMonitorResult("Closing the 'CompMan Add-in' (" & mAddin.WbkName & ") failed with:" & vbLf & _
                               "(" & sErrDesc & ")" _
                              ) = "Failed"
    Else
        mMe.RenewMonitorResult = "Passed"
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Renew_070_DeleteAddInInstanceWorkbook() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the Add-in instance Workbbook has been deleted
' ----------------------------------------------------------------------------
    Const PROC = "Renew_070_DeleteAddInInstanceWorkbook"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Delete the 'CompMan Add-in' Workbook' (" & mAddin.WbkName & ")"
    
    With New FileSystemObject
        If .FileExists(WbkFullName) Then
            On Error Resume Next
            mAddin.WbkRemove WbkFullName
            Renew_070_DeleteAddInInstanceWorkbook = Err.Number = 0
            If Renew_070_DeleteAddInInstanceWorkbook Then
                mMe.RenewMonitorResult = "Passed"
            Else
                mMe.RenewMonitorResult("Deleting the 'CompMan Add-in' (" & mAddin.WbkName & ") failed with:" & vbLf & _
                                       "(" & Err.Description & ")" _
                                      ) = "Failed"
            End If
        Else
            Renew_070_DeleteAddInInstanceWorkbook = True
            mMe.RenewMonitorResult = "Passed"
        End If
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Renew_080_SaveDevInstncWorkbookAsAddin() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the development instance Workbook has successfully saved
' as Add-in.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_080_SaveDevInstncWorkbookAsAddin"
    
    On Error GoTo eh
    mCompManClient.Events ErrSrc(PROC), False
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Save the 'Development-Instance-Workbook' (" & DevInstncName & ") as 'CompMan Add-in' (" & mAddin.WbkName & ")"
    
    With Application
        If Not FSo.FileExists(mAddin.WbkFullName) Then
            '~~ At this point the Add-in must no longer exist at its location
            On Error Resume Next
            wbkSource.SaveAs WbkFullName, FileFormat:=ADDIN_FORMAT
            Renew_080_SaveDevInstncWorkbookAsAddin = Err.Number = 0
            If Not Renew_080_SaveDevInstncWorkbookAsAddin Then
                mMe.RenewMonitorResult("Save Development instance as Add-in instance  " & mBasic.Spaced("failed!") _
                                      ) = "Failed"
            Else
                mMe.RenewMonitorResult() = "Passed"
            End If
        Else ' file still exists
            mMe.RenewMonitorResult("Setup/Renew of the 'CompMan Add-in' as copy of the 'Development-Instance-Workbook'  " & mBasic.Spaced("failed!") _
                              ) = "Failed"
        End If
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    mCompManClient.Events ErrSrc(PROC), True
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Renew_090_OpenAddinInstncWorkbook() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the Add-in instance Workbook has successfully been opened.
' ----------------------------------------------------------------------------
    Const PROC = "Renew_090_OpenAddinInstncWorkbook"
    
    On Error GoTo eh
    Dim wb              As Workbook
    Dim sBaseAddinName  As String
    Dim sBaseDevName    As String
    
    mBasic.BoP ErrSrc(PROC)
    If Not mAddin.IsOpen Then
        If FSo.FileExists(mAddin.WbkFullName) Then
            mMe.RenewAction = "Re-open the 'CompMan Add-in' (" & mAddin.WbkName & ")"
            On Error Resume Next
            Set wb = Application.Workbooks.Open(mAddin.WbkFullName, False, True)
            If Err.Number = 0 Then
                With New FileSystemObject
                    sBaseAddinName = .GetBaseName(wb.Name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.Name)
                    wb.VBProject.Name = sBaseAddinName
                End With
                mMe.RenewMonitorResult() = "Passed"
                Renew_090_OpenAddinInstncWorkbook = True
            Else
                mMe.RenewMonitorResult("(Re)opening the 'CompMan Add-in' (" & mAddin.WbkName & ") failed with:" & vbLf & _
                                       "(" & Err.Description & ")" _
                                      ) = "Failed"
            End If
        End If
    Else
        Renew_090_OpenAddinInstncWorkbook = True
    End If

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub Renew_110_SetupAutoOpen()
    Const PROC = "Renew_110_SetupAutoOpen"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Setup/maintain auto-open for the 'CompMan Add-in'"
    wsConfig.AutoOpenAddinSetup
    mMe.RenewMonitorResult() = "Passed"
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_120_SetupWinMergeIni()
    Const PROC = "Renew_120_SetupWinMergeIni"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    mWinMergeIni.Setup
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Renew_100_RestoreReferencesToAddIn()
    Const PROC = "Renew_100_RestoreReferencesToAddIn"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim wb              As Workbook
    Dim sWbs            As String
    Dim bOneRestored    As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    mMe.RenewAction = "Restore all saved 'References' to the 'CompMan Add-in' in open Workbooks"
    
    For Each v In dctAddInRefs
        Set wb = v
        wb.VBProject.References.AddFromFile WbkFullName
        sWbs = wb.Name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewMonitorResult() = "Passed"
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
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
    
    mCompManClient.Events ErrSrc(PROC), False
    mCompMan.ServiceInitiate ThisWorkbook, PROC
    
    mBasic.BoP ErrSrc(PROC)
    lRenewStep = 0
    bSucceeded = False
                            
    '~~ Get the CompMan base configuration confirmed or changed
    If Not Renew_010_ConfirmConfig Then GoTo xt
                         
    mCompManMenuVBE.MenuRemove
    
    '~~ Assert the Renew service is performed from within the development instance Workbbok
    '~~ Note that the distinction of the instances requires the above basic configuration confirmed
    If Not Renew_020_DevInstnc Then GoTo xt
    
    '~~ Assert no Workbooks are open referring to the Add-in
    Renew_030_SaveAndRemoveAddInReferences
    If Not bAllRemoved Then GoTo xt

    '~~ Assure the current version of the Add-in's development instance has been saved
    '~~ Note: Unconditionally saving the Workbook does an incredible trick:
    '~~       The un-unstalled and IsAddin=False Workbook is released from the Application
    '~~       and no longer considered "used"
    Renew_040_DevInstncWorkbookSave
    wbkSource.Activate
          
    '~~ Attempt to turn Add-in to "IsAddin=False" in order to uninstall and subsequently close it
    Renew_050_Set_IsAddin_ToFalse
    If Not Renew_060_CloseCompManAddinWorkbook Then GoTo xt
  
    '~~ Attempt to delete the Add-in Workbook file
    If Not Renew_070_DeleteAddInInstanceWorkbook Then GoTo xt
        
    '~~ Attempt to save the development instance as Add-in
    If Not Renew_080_SaveDevInstncWorkbookAsAddin Then GoTo xt
    
    '~~ Re-open, re-activate the Add-in
    If Not Renew_090_OpenAddinInstncWorkbook Then GoTo xt
        
    '~~ Re-instate references to the Add-in which had been removed
    Renew_100_RestoreReferencesToAddIn
    Renew_110_SetupAutoOpen
    Renew_120_SetupWinMergeIni
    
    bSucceeded = True
    
xt: RenewFinalResult bSucceeded
    Application.ScreenUpdating = False
    Servicing.ServicedWbk.Activate
    wsService.Activate
    wsConfig.CurrentStatus
    wsConfig.Activate
    DoEvents
    Application.SendKeys "%{Tab}" ' should bring the monitor message to front
    ThisWorkbook.Saved = True
    mBasic.EoP ErrSrc(PROC)
    mCompManClient.Events ErrSrc(PROC), True
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


