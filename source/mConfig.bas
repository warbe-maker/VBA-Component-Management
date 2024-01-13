Attribute VB_Name = "mConfig"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mConfig: Services for the initial setup of the default
' ======================== environment and the maintenance of the current
'                          setup.
' Public Properties:
' ------------------
' - CompManParentFolderNameDefault
' - ServicedRootFolderNameCurrent
'
' Public services:
' ----------------
' - Adjust              Adjusts the configuration in the "Config" Worksheet
'                       when the serviced root folder has been moved to
'                       another location and/or has been renamed.
' - DefaultEnvDisplay   Displays the to-be-setup default environment.
' - EnvIsMissing        Returns TRUE when the current ThisWorkbook parent
'                       folder has no Addin and ThisWorkbooks parent.parent
'                       folder has no Common-Components folder.
' - DefaultEnvConfirmed Confirms the setup of the default environment.
' - SelfSetupDefaultEnvironment     Sets up CompMan's default environment of files and
'                       folders
' ----------------------------------------------------------------------------
Public Const DEFAULT_FOLDER_COMPMAN_PARENT     As String = "CompMan"
Public Const DEFAULT_FOLDER_COMMON_COMPONENTS  As String = "Common-Components"
Public Const DEFAULT_FOLDER_COMPMAN_ROOT       As String = "CompManServiced"
Public Const DEFAULT_FOLDER_EXPORT             As String = "source"

Private Property Get CommonCompsFolderNameCurrent() As String:      CommonCompsFolderNameCurrent = ServicedRootFolderNameCurrent & "\" & DEFAULT_FOLDER_COMMON_COMPONENTS:                  End Property

Private Property Get CommonCompsFolderNameDefault() As String:      CommonCompsFolderNameDefault = ServicedRootFolderNameDefault & "\" & DEFAULT_FOLDER_COMMON_COMPONENTS:                  End Property

Private Property Get CompManParentFolderNameCurrent() As String:    CompManParentFolderNameCurrent = fso.GetFile(ThisWorkbook.FullName).ParentFolder:                                       End Property

Public Property Get CompManParentFolderNameDefault() As String:     CompManParentFolderNameDefault = ServicedRootFolderNameDefault & "\" & DEFAULT_FOLDER_COMPMAN_PARENT:                   End Property

Private Property Get ExportFolderNameCurrent() As String:           ExportFolderNameCurrent = CompManParentFolderNameCurrent & "\" & wsConfig.FolderExport:                                 End Property

Private Property Get ExportFolderNameDefault() As String:           ExportFolderNameDefault = CompManParentFolderNameDefault & "\" & DEFAULT_FOLDER_EXPORT:                                 End Property

Public Property Get ServicedRootFolderNameCurrent() As String:      ServicedRootFolderNameCurrent = fso.GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder:                           End Property

Private Property Get ServicedRootFolderNameDefault() As String:     ServicedRootFolderNameDefault = fso.GetFile(ThisWorkbook.FullName).ParentFolder & "\" & DEFAULT_FOLDER_COMPMAN_ROOT:    End Property

Public Property Get VBCompManAddinFolderNameCurrent() As String:   VBCompManAddinFolderNameCurrent = CompManParentFolderNameCurrent & "\" & "Addin":                                        End Property

Private Property Get VBCompManAddinFolderNameDefault() As String:   VBCompManAddinFolderNameDefault = CompManParentFolderNameDefault & "\" & "Addin":                                       End Property

Public Sub Adjust()
' ----------------------------------------------------------------------------
' Adjusts the configuration in the "Config" Worksheet when the serviced root
' folder has been moved to another location and/or has been renamed.
' ----------------------------------------------------------------------------
    With wsConfig
        .FolderCompManRoot = ServicedRootFolderNameCurrent          ' adjust the root path and
        .FolderCommonComponentsPath = CommonCompsFolderNameCurrent  ' adjust the common components folder
        If .AutoOpenAddinIsSetup Then .AutoOpenAddinSetup           ' re-setup a setup Addin auto-open
        If .AutoOpenCompManIsSetup Then .AutoOpenCompManSetup       ' re-setup a setup CompMan.xlsb auto-open
    End With
    
End Sub

Public Function DefaultEnvDisplay(ByVal BttnGoAhead As String) As Variant
' ----------------------------------------------------------------------------
' Displays the to-be-setup default environment.
' ----------------------------------------------------------------------------
    Dim lMax    As Long
    Dim Msg     As mMsg.udtMsg
    
    With Msg
        With .Section(1).Text
            .Text = "CompMan will now setup the below default files and folder environment at the current " & _
                    "location. Once set up the Workbook is closed. The setup top level folder may then be " & _
                    "moved to any other location and also renamed to any desired name."
        End With
        With .Section(2).Text
            .MonoSpaced = True
            .Text = DEFAULT_FOLDER_COMPMAN_ROOT & vbLf & _
                    " |                                       " & vbLf & _
                    " +--" & DEFAULT_FOLDER_COMPMAN_PARENT & vbLf & _
                    " |  +--CompMan.xlsb                     " & vbLf & _
                    " |  +--" & wsConfig.FolderExport & vbLf & _
                    " |  +--WinMerge.ini                     " & vbLf & _
                    " |                                       " & vbLf & _
                    " +--" & DEFAULT_FOLDER_COMMON_COMPONENTS & vbLf & _
                    "    +--CompManClient.bas                "
        End With
        With .Section(3)
            .Label.FontBold = True
            .Label.Text = Replace(BttnGoAhead, vbLf, " ")
            .Text.Text = "The above files and folders structure will be created and the opened Workbook will be saved into the" & vbLf & _
                         CompManParentFolderNameDefault & vbLf & _
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
        With .Section(5).Label
            .FontColor = rgbBlue
            .Text = "See README chapter 'Files and Folders' for more information"
            .OnClickAction = mCompMan.GITHUB_REPO_URL & mCompMan.README_DEFAULT_FILES_AND_FOLDERS
        End With
    End With
    
    DefaultEnvDisplay = mMsg.Dsply(dsply_title:="CompMan's self setup (when opened for the very first time after download)" _
                                 , dsply_msg:=Msg _
                                 , dsply_buttons:=mMsg.Buttons(BttnGoAhead))

End Function

Public Function EnvIsMissing() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current ThisWorkbook parent folder has no Addin and
' ThisWorkbooks parent.parent folder has no Common-Components folder
' ----------------------------------------------------------------------------
    EnvIsMissing = Not fso.FolderExists(VBCompManAddinFolderNameCurrent) _
               And Not fso.FolderExists(CommonCompsFolderNameCurrent)
End Function

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mConfig" & "." & es_proc
End Function

Public Sub SetupConfirmed()
' ----------------------------------------------------------------------------
' Confirms the setup of the default environment.
' ----------------------------------------------------------------------------
    Dim Msg             As mMsg.udtMsg
    Dim sSetupLocation  As String
    
    sSetupLocation = fso.GetFolder(ServicedRootFolderNameCurrent).ParentFolder.ParentFolder
    
    With Msg
        .Section(1).Text.Text = "CompMan's default environment has been setup at the location the Workbook " & _
                                "had been opened (" & sSetupLocation & ") as follows:"
        With .Section(2).Text
            .MonoSpaced = True
            .Text = fso.GetFolder(ServicedRootFolderNameCurrent).Name & vbLf & _
                    " |                                       " & vbLf & _
                    " +--" & fso.GetFolder(CompManParentFolderNameCurrent).Name & vbLf & _
                    " |  +--" & ThisWorkbook.Name & vbLf & _
                    " |  +--" & wsConfig.FolderExport & vbLf & _
                    " |  +--WinMerge.ini                      " & vbLf & _
                    " |                                       " & vbLf & _
                    " +--" & fso.GetFolder(CommonCompsFolderNameCurrent).Name & vbLf & _
                    "    +--mCompManClient.bas                 "
        End With
        .Section(3).Text.Text = "CompMan is now ready for servicing any Workbook with an enabled service. When " & _
                                "the Workbook has been closed the ""Serviced Root Folder"" (" & _
                                ServicedRootFolderNameCurrent & ") may be moved to its final destination and/or renamed."
        With .Section(4).Label
            .FontColor = rgbBlue
            .Text = "See the corresponding README for how to enable a Workbook for being serviced"
            .OnClickAction = "https://github.com/warbe-maker/VBCompMan/blob/master/README.md?#enabling-the-services-serviced-or-not-serviced"
        End With
    End With
    mMsg.Dsply dsply_title:="Setup of CompMan's default environment completed!" _
             , dsply_msg:=Msg _
             , dsply_buttons:=vbOKOnly
    
End Sub

Public Sub SelfSetupExportCompManClient()

    Dim Comp        As clsComp
    Dim vbc         As VBComponent
    
    '~~ Export the mCompManClient to the Common-Components folder
    If Services Is Nothing Then Set Services = New clsServices
    Services.ServicedWbk = ThisWorkbook
    For Each vbc In Services.ServicedWbk.VBProject.VBComponents
        If vbc.Name = "mCompManClient" Then
            Services.ServicedItem = vbc
                    
            Set Comp = New clsComp
            With Comp
                .Wrkbk = ThisWorkbook
                .VBComp = vbc
                .KindOfComp = enCommCompHosted

                With LogServiced
                    .KeepDays = 2 ' a new log-file is created after 48 hours
                    .WithTimeStamp
                    .FileFullName = Services.ServicedWbk.Path & "\" & fso.GetBaseName(Services.ServicedWbk.Name) & ".Services.log"
                    .MaxItemLengths Len(Comp.TypeString), Len(vbc.Name)
                    .AlignmentItems "|L|L.:|L|"
                    mCompMan.LogFileService = .FileName
                    wsService.CurrentServiceLogFileFullName = .FileFullName
                    .NewFile
                End With
        
                '~~ Export, initialize the Revision Number of both the hosted and
                '~~ the raw in the Common-Components folder and register it as hosted.
                .Export
                With Services
                    .ServicedItem = vbc
                    .NoOfItemsServiced = .NoOfItemsServiced + 1
                    .ServicedItemLogEntry "Common Component hosted: initially exported by the ""self-setup"" process!"
                    .ServicedItemLogEntry "Common Component hosted: Revision Number initialized with " & Comp.LastModAtDateTime
                    .ServicedItemLogEntry "Common Component hosted: Export-File copied to " & wsConfig.FolderCommonComponentsPath
                End With
            End With
            Exit For
        End If
    Next vbc

    Set Comp = Nothing
    Set LogServiced = Nothing
    
End Sub

Public Sub SelfSetupDefaultEnvironment()
' ----------------------------------------------------------------------------
' Sets up CompMan's default environment of files and folders
'
' See: https://github.com/warbe-maker/VBCompMan/blob/master/README.md?#compmans-default-files-and-folders-environment
' ----------------------------------------------------------------------------
    Const PROC = "SelfSetupDefaultEnvironment"
        
    On Error GoTo eh
    Dim FldrExport          As String
        
    With fso
        If Not .FolderExists(ServicedRootFolderNameDefault) Then .CreateFolder ServicedRootFolderNameDefault
        If Not .FolderExists(mConfig.CommonCompsFolderNameDefault) Then .CreateFolder mConfig.CommonCompsFolderNameDefault
        If Not .FolderExists(CompManParentFolderNameDefault) Then .CreateFolder CompManParentFolderNameDefault
        If Not .FolderExists(ExportFolderNameDefault) Then .CreateFolder ExportFolderNameDefault
        If Not .FolderExists(VBCompManAddinFolderNameDefault) Then .CreateFolder VBCompManAddinFolderNameDefault
    End With
    
    With wsConfig
        .FolderCompManRoot = ServicedRootFolderNameDefault
        .FolderCommonComponentsPath = mConfig.CommonCompsFolderNameDefault
        .FolderSyncArchive = vbNullString
        .FolderSyncTarget = vbNullString
        .AutoOpenCompManRemove
        .AutoOpenAddinRemove
        If Not .Verified Then .Activate
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

