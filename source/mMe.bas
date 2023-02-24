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
Private Const DEVLP_WORKBOOK_EXTENSION          As String = "xlsb"              ' Extension may depend on Excel version
Private Const DEFAULT_FOLDER_COMPMAN_PARENT     As String = "CompMan"
Private Const DEFAULT_FOLDER_COMMON_COMPONENTS  As String = "Common-Components"
Private Const DEFAULT_FOLDER_COMPMAN_ROOT       As String = "CompManServiced"

Private wbDevlp                 As Workbook
Private wbkSource               As Workbook                     ' This development instance as the renew source
Private bSucceeded              As Boolean
Private bAllRemoved             As Boolean
Private dctAddInRefs            As Dictionary
Private lRenewStep              As Long
Private sRenewAction            As String
Private bRenewTerminatedByUser  As Boolean

Public CompManRoot              As String
Public BaseName                 As String
Public Extension                As String
Public ServicingEnabled         As Boolean

Public Function AssertedServicingEnabled() As Boolean
' ---------------------------------------------------------------------------
' When TRUE is returned either the opened Workbook is the Addin instance or
' - the Workbook has either been opened the very first time after the very
'   first download and the default environment has been established - ready
'   for the Workbook for being re-opened from within it
' - the Workbook has been opened and the environment conforms with the
'   current configuration or
' - the Workbook has been opened from within an environment which has changed
'   but still meet all the requirements following is asserted:
' - When the Workbook is the Addin instanceEnsures that the CompMan development instance Workbook is able to function
' as the servicing instance. Because the Addin is (in case) saved from a
' servicing enabled Workbook it is enabled by default.
' ---------------------------------------------------------------------------
    Const PROC = "AssertedServicingEnabled"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    BaseName = fso.GetBaseName(ThisWorkbook.Name)
    Extension = fso.GetExtensionName(ThisWorkbook.Name)
    If mMe.IsAddinInstnc Then
        '~~ Because the Addin is (in case) saved from a servicing enabled Workbook it is enabled by default
        ServicingEnabled = True
    Else
        If Not AssertedOfficeVersion Then GoTo xt
        If Not AssertedFilesAndFldrsStructure Then GoTo xt
        If Not fso.FileExists(mCompManCfg.CompManCfgFileFullName) Then
            wsConfig.CompManCfgSaveConfig
        Else
            wsConfig.CompManCfgRestoreConfig
        End If
    End If
    AssertedServicingEnabled = True
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function AssertedFilesAndFldrsStructure() As Boolean
' ---------------------------------------------------------------------------
' Performed with each open. Ensures that the Workbook is running from within
' the required files and folders structure.
' ---------------------------------------------------------------------------
    Const PROC = "AssertedFilesAndFldrsStructure"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim Msg                 As mMsg.TypeMsg
    Dim BttnGoAhead         As String
    Dim FldrAddin           As String
    Dim FldrCompManRoot     As String
    Dim FldrCompManParent   As String
    Dim FldrCommonComps     As String
    Dim FldrExport          As String
    Dim lMax                As Long
    
    If mCompManCfg.Exists Then
        FldrCompManRoot = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path
        FldrAddin = ThisWorkbook.Path & "\" & "Addin"
        If FldrCompManRoot <> wsConfig.FolderCompManRoot Then
            With wsConfig
                '~~ The CompMan-Root-Folder must have been renamed or moved to a different location ...
                .FolderCompManRoot = FldrCompManRoot
                mCompManCfg.FolderCompManRoot = FldrCompManRoot
                '~~ ... and this concerns also the Addin folder
                .FolderAddin = FldrAddin
                mCompManCfg.FolderAddin = FldrAddin
                '~~ Re-setup any setup auto-open
                If .AutoOpenAddinIsSetup Then .AutoOpenAddinSetup
                If .AutoOpenCompManIsSetup Then .AutoOpenCompManSetup
            End With
        End If
        
        '~~ Note: When a configured Sync-Target-Folder and or the Sync-Archice-Folder had been renamed or moved
        '         the configuration restored from the CompMan.cfg file will become invalid.
        '~~ Restore the last saved configuration. This ensures that for a subsequently downloaded
        '~~ CompMan.xlsb Workbook the local configuration is made available again in the wsConfig Worksheet.
        wsConfig.CompManCfgRestoreConfig
        If wsConfig.Verified Then
            '~~ Nothing had been changed while the Workbook was closed
            AssertedFilesAndFldrsStructure = True
        Else
            '~~ The configuration is no loger valid. This may be the case when the CompMan root folder
            '~~ has been renamed or moved to another location
            AssertedFilesAndFldrsStructure = False
            wsConfig.Activate
        End If
        CompManRoot = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path
        GoTo xt
    End If
        
    '~~ Because no CompMan.cfg exists the CompMan Workbook has either been downloaded and opened the
    '~~ very first time or at least has been opened for the vers first time from this location.
    '~~ It is thus concluded that there is currently no default environment (folders and files)
    '~~ and it is now to be setup - provided the user confirms it.
    FldrCompManRoot = ThisWorkbook.Path & "\" & DEFAULT_FOLDER_COMPMAN_ROOT     ' 1. to be setup
    FldrCommonComps = FldrCompManRoot & "\" & DEFAULT_FOLDER_COMMON_COMPONENTS  ' 2. to be setup
    FldrCompManParent = FldrCompManRoot & "\" & DEFAULT_FOLDER_COMPMAN_PARENT   ' 3. to be setup
    FldrExport = FldrCompManParent & "\" & wsConfig.FolderExport                ' 4. to be setup
    FldrAddin = FldrCompManParent & "\Addin"                                    ' 5. to be setup
    
    BttnGoAhead = "Ok!" & vbLf & vbLf & _
                  "Go ahead and set it up"
                 
    With Msg
        With .Section(1).Text
            .Text = "CompMan will now setup the below default folder/files environment at the current " & _
                    "download location, which is essential in order to establish it as a servicing instance. " & _
                    "Once set up the top level folder may be moved to any other location."
        End With
        With .Section(2).Text
            .MonoSpaced = True
            .Text = DEFAULT_FOLDER_COMPMAN_ROOT & vbLf & _
                    "|                                       " & vbLf & _
                    "+--" & DEFAULT_FOLDER_COMPMAN_PARENT & vbLf & _
                    "|  |                                    " & vbLf & _
                    "|  +---CompMan.xlsb                     " & vbLf & _
                    "|  +---" & wsConfig.FolderExport & vbLf & _
                    "|  +---CompMan.cfg                      " & vbLf & _
                    "|                                       " & vbLf & _
                    "+--" & DEFAULT_FOLDER_COMMON_COMPONENTS & vbLf & _
                    "   |                                    " & vbLf & _
                    "   +---CompManClient.bas                "
        End With
        With .Section(3)
            .Label.FontBold = True
            .Label.Text = Replace(BttnGoAhead, vbLf, " ")
            .Text.Text = "The above files and folders structure will be created and the opened Workbook will be saved into the" & vbLf & _
                         FldrCompManParent & vbLf & _
                         "folder and closed. Re-opening it from within its new folder structure will finalize " & _
                         "CompMan's self setup process."
        End With
        lMax = mBasic.Max(Len(DEFAULT_FOLDER_COMPMAN_ROOT) _
                        , Len(DEFAULT_FOLDER_COMPMAN_PARENT) _
                        , Len(ThisWorkbook.Name) _
                        , Len(wsConfig.FolderExport) _
                        , Len("CompMan.cfg") _
                        , Len(DEFAULT_FOLDER_COMMON_COMPONENTS) _
                        , Len("CompManClient.bas"))
        
        With .Section(4).Text
            .MonoSpaced = True
            .Text = mBasic.Align(DEFAULT_FOLDER_COMPMAN_ROOT, lMax, AlignLeft, " ", ".") & _
                                                                                                ": CompMan's ""serviced"" folder (only Workbooks when opened from within this folder will be serviced)" & vbLf & _
                    mBasic.Align(DEFAULT_FOLDER_COMPMAN_PARENT, lMax, AlignLeft, " ", ".") & _
                                                                                                ": CompMan's dedicated default parent folder" & vbLf & _
                    mBasic.Align("CompMan.xlsb", lMax, AlignLeft, " ", ".") & _
                                                                                                ": CompMan's (this) ""servicing"" Workbook" & vbLf & _
                    mBasic.Align(wsConfig.FolderExport, lMax, AlignLeft, " ", ".") & _
                                                                                                ": Default folder for exported (changed) components, maintained by CompMan for each serviced" & vbLf & _
                    String(lMax, " ") & _
                                                                                                "  Workbook's. The name may be re-configured." & vbLf & _
                    mBasic.Align("CompMan.cfg", lMax, AlignLeft, " ", ".") & _
                                                                                                ": Initialized with the self-setup defaults, subsequently maintained through CompMan's configuration" & vbLf & _
                    String(lMax, " ") & _
                                                                                                "  Worksheet """ & wsConfig.Name & """." & vbLf & _
                    mBasic.Align(DEFAULT_FOLDER_COMMON_COMPONENTS, lMax, AlignLeft, " ", ".") & ": The default folder for ""Common Components""" & vbLf & _
                    mBasic.Align("CompManClient.bas", lMax, AlignLeft, " ", ".") & _
                                                                                                ": The ""Common Component"" hosted by CompMan for being imported" & vbLf & _
                    String(lMax, " ") & _
                                                                                                "  into any Workbook 's VB-Project for being serviced by CompMan." & vbLf & _
                    String(lMax, " ") & _
                                                                                                "  Will be provided the first time the Workbook is saved/closed. "
        End With
    End With
    
    If mMsg.Dsply(dsply_title:="CompMan's self setup (when opened for the very first time after download)" _
                , dsply_msg:=Msg _
                , dsply_buttons:=mMsg.Buttons(BttnGoAhead)) = BttnGoAhead _
    Then
        If Not fso.FolderExists(FldrCompManRoot) Then fso.CreateFolder FldrCompManRoot
        If Not fso.FolderExists(FldrCommonComps) Then fso.CreateFolder FldrCommonComps
        If Not fso.FolderExists(FldrCompManParent) Then fso.CreateFolder FldrCompManParent
        If Not fso.FolderExists(FldrExport) Then fso.CreateFolder FldrExport
        If Not fso.FolderExists(FldrAddin) Then fso.CreateFolder FldrAddin
        
        '~~ Save the initially opened CompMan Workbook to its new setup parent folder
        Application.EnableEvents = False
        With wsConfig
            .FolderAddin = FldrAddin
            .FolderCompManRoot = FldrCompManRoot
            .FolderCommonComponentsPath = FldrCommonComps
            .FolderSyncArchive = vbNullString
            .FolderSyncTarget = vbNullString
            .AutoOpenCompManRemove
            .AutoOpenAddinRemove
            If Not .Verified Then .Activate
        End With
        
        ThisWorkbook.SaveAs FldrCompManParent & "\" & ThisWorkbook.Name
        '~~ CompMan's .cfg-file
        wsConfig.CompManCfgSaveConfig
        AssertedFilesAndFldrsStructure = True
        Application.EnableEvents = True
        ThisWorkbook.Close False
    End If
                        
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

    Private Function FolderCompManRootHasMovedToAnotherLocation() As Boolean
    With New FileSystemObject
        FolderCompManRootHasMovedToAnotherLocation = wsConfig.FolderCompManRoot <> .GetFolder(ThisWorkbook.Path).ParentFolder.Path
    End With
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

Private Property Get ADDIN_FORMAT() As XlFileFormat ' = ... needs adjustment when the above is changed
    ADDIN_FORMAT = xlOpenXMLAddIn
End Property

Public Property Get AutoOpenShortCutCompManWbk()
    AutoOpenShortCutCompManWbk = Environ$("APPDATA") & "\Microsoft\Excel\XLSTART\CompManWbk.lnk"
End Property

Public Property Get DevInstncFullName() As String
    Dim fso As New FileSystemObject
    DevInstncFullName = wsConfig.FolderCompManRoot & DBSLASH _
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

Public Property Let RenewMonitorResult(Optional ByVal la_last As Boolean = False, _
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
    
    If cfg_addin Then Config = wsConfig.FolderCompManRootIsValid _
                           And wsConfig.FolderExportIsValid _
                           And wsConfig.FolderAddInIsValid
    
    If cfg_sync Then Config = wsConfig.FolderCompManRootIsValid _
                          And wsConfig.FolderExportIsValid _
                          And wsConfig.FolderSyncTargetIsValid
    
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
        RenewMonitorResult(True) = "Passed"
    Else
        mMe.RenewAction = "Not Successful!"
        mMe.RenewMonitorResult(True) = "Failed"
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
    Dim sWbs        As String
    Dim bOneRemoved As Boolean

    mMe.RenewAction = "Save and remove references to the Add-in from open Workbooks"
    mAddin.ReferencesRemove dctAddInRefs, sWbs, bOneRemoved, bAllRemoved
    If bOneRemoved Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewMonitorResult() = "Passed"
    Else
'        mMe.RenewMonitorResult(sRenewAction & vbLf & "None of the open Workbook's VBProject had a 'Reference' to the 'CompMan Add-in'" _
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
    
    mMe.RenewAction = "Set the 'IsAddin' property of the 'CompMan Add-in' to FALSE"
    If mAddin.Set_IsAddin_ToFalse Then
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
                    sBaseAddinName = .GetBaseName(wb.Name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.Name)
                    wb.VBProject.Name = sBaseAddinName
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
    wsConfig.AutoOpenAddinSetup
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
        sWbs = wb.Name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewMonitorResult() = "Passed"
    Else
'        mMe.RenewMonitorResult(sRenewAction & vbLf & "Restoring 'References' did not find any saved to restore" _
'                              ) = "Passed"
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


