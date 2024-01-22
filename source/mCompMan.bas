Attribute VB_Name = "mCompMan"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard Module mCompMan: Services for the management of VB-Components in
' ========================= Workbooks. Services for Workbooks having them
' enabled are onyl provided when:
' - either this Workbook or its Addin instance is open
' - the serviced Workbook resides in its dedicated folder
' - the Workbook is opened from within the 'CompMan-Managed' folder. I.e.
'   that productive Workbooks at their productive location are ignored.
' The services are provided through an interface component (mCompManClient)
' which is copied into the to-be-serviced Workbook's VB-Project.
'
' Public services:
' ----------------
' DisplayCodeChange       Displays the current difference between a
'                         component's code and its current Export-File
' ExportAll               Exports all components into the Workbook's
'                         dedicated folder (created when not existing)
' ExportChangedComponents Exports all components of which the code in the
'                         Export-File differs from the current code.
' Service
'
' Uses Common Components: mBasic
'                         mErrhndlr
'                         mFso
'                         mWrkbk
' Requires:
' - Reference to: - "Microsoft Visual Basic for Applications Extensibility ..."
'                 - "Microsoft Scripting Runtime"
'                 - "Windows Script Host Object Model"
'                 - Trust in the VBA project object modell (Security
'                   setting for makros)

' Usage:   This Workbbok's services may additionally be available Add-in when
'          setup and open.
'
' Common coding rules: (where case matters!)
' - nme indicates a Name object
' - wbk indicates a Workbook object
' - wsh indicate a Worksheet object
' - shp indicates a Shape object
' - dct indicates a Dictionary object
' - procedure's arguments (in order to distinguish them from
'   anything else):
'   -- have at least one underscore letter (_), usually after a prefix
'   -- are strictly in lower case letters
'   -- only lower case letters and underscrores (_) are not used for anything
'      else!
' - Constants are in upper-case letters with underscores (_) for better
'   readability and their lower-case counterpart is never used as argument
'
' See also:
' https://warbe-maker.github.io/warbe-maker.github.io/vba/excel/code/component/management/2021/03/02/Programatically-updating-Excel-VBA-code.html
' ----------------------------------------------------------------------------
'
' W. Rauschenberger Berlin Jan 2024
' -------------------------------------------------------------------------------
Public Const GITHUB_REPO_URL                    As String = "https://github.com/warbe-maker/VBA-Component-Management"
Public Const README_SYNC_CHAPTER                As String = "#using-the-synchronization-service"
Public Const README_SYNC_CHAPTER_NAMES          As String = "#names-synchronization"
Public Const README_DEFAULT_FILES_AND_FOLDERS   As String = "#compmans-default-files-and-folders-environment"
Public Const README_CONFIG_CHANGES              As String = "#configuration-changes-compmans-config-worksheet"
Public Const FORMAT_REV_DATE                    As String = "YYYY-MM-DD"
Public Const FORMAT_REV_NO                      As String = "000"
Public Const BTTN_DSPLY_DIFF                    As String = "Display difference" & vbLf & "(of Export-Files)"
Public Const SRVC_PROGRESS_SCHEME               As String = "<srvc> <by> <serviced>: <n> of <m> <op> <comps> <dots>"

Public LogServiced                              As clsLog ' log writen for the serviced Workbook
Public LogService                               As clsLog ' the servicing Workbooks own log
Public Comps                                    As clsComps
Public Services                                 As clsServices
Public CompManDat                               As clsCompManDat
Public CommComps                                As clsCommComps
Public Msg                                      As udtMsg
Public fso                                      As New FileSystemObject
Public Prgrss                                   As clsProgress

#If clsTrc = 1 Then
    Public Trc                                  As clsTrc
#End If

Public Enum enKindOfComp            ' The kind of VBComponent in the sense of CompMan
    enUnknown = 0                   ' When the kind of component yet has not been analysed
    enCommCompHosted = 1            ' When the component is claimed hosted by the open serviced Workbook
    enCommCompUsed = 2              ' When the component is a used Common Component, i.e. one known in the Common-Componwents folder
    enInternal = 3                  ' When the component is not a Common Component (may still be one with the same name)
End Enum

Public Enum siCounter
    sic_comps
    sic_comps_changed
    sic_used_comm_vbc_outdated
End Enum

Private sLogFileService         As String
Private sLogFileServicedSummary  As String

Public Property Get LogFileService() As String:         LogFileService = sLogFileService:   End Property

Public Property Let LogFileService(ByVal s As String):  sLogFileService = s:                End Property

Public Property Get LogFileServicedSummary() As String:         LogFileServicedSummary = sLogFileServicedSummary:   End Property

Public Property Let LogFileServicedSummary(ByVal s As String):  sLogFileServicedSummary = s:                End Property

Public Sub CheckForUnusedPublicItems()
' ----------------------------------------------------------------
' Attention! The service requires the VBPUnusedPublic.xlsb
'            Workbook open. When not open the service terminates
'            without notice.
' ----------------------------------------------------------------
    Const COMPS_EXCLUDED = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"
    Const LINES_EXCLUDED = "Select Case mMe.ErrMsg(ErrSrc(PROC))" & vbLf & _
                           "Case vbResume:*Stop:*Resume" & vbLf & _
                           "Case Else:*GoTo xt"
    On Error Resume Next
    '~~ Providing all (optional) arguments saves the Workbook selection dialog and the VBComponents selection dialog
    Application.Run "VBPUnusedPublic.xlsb!mUnused.Unused", ThisWorkbook, COMPS_EXCLUDED, LINES_EXCLUDED

End Sub

Public Function CommCompRegStateEnum(ByVal s As String) As enCommCompRegState
    Select Case s
        Case "hosted":  CommCompRegStateEnum = enRegStateHosted
        Case "used":    CommCompRegStateEnum = mComp.enRegStateUsed
        Case "private": CommCompRegStateEnum = mComp.enRegStatePrivate
    End Select
End Function

Public Function CommCompRegStateString(ByVal en As enCommCompRegState) As String
    Select Case en
        Case enRegStateHosted:  CommCompRegStateString = "hosted"
        Case enRegStateUsed:    CommCompRegStateString = "used"
        Case enRegStatePrivate: CommCompRegStateString = "private"
    End Select
End Function

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCompMan" & "." & es_proc
End Function

Public Function ExportChangedComponents(ByRef e_wbk_serviced As Workbook, _
                               Optional ByVal e_hosted As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Exports any component the code had been modified (UserForm also when the
' form has changed) to the configured export folder (defaults to 'source').
'
' The function is terminated (returns FALSE) without further notice when:
' a) the serviced root folder is invalid (not configured or not existing)
' b) the serviced Workbook is located outside the serviced folder
'
' The function is terminated (returns FALSE but loggs the reason) when:
' a) the Workbook is one restored by Excel
' b) the serviced Workbook does not reside in a folder exclusivelyx (i.e. the
'    Workbook does not live in its own dedicated folder
' c) WinMerge is not installed
'
' Precondition: The service has been checked by the client to be able to run.
' ----------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
'    mCompManClient.Events ErrSrc(PROC), False
    
    Set Services = New clsServices
    With Services
        .ServicedWbk = e_wbk_serviced
        .CurrentService = mCompManClient.SRVC_EXPORT_CHANGED
        .EstablishExecTraceFile
    End With
        
    mBasic.BoP ErrSrc(PROC)
    Services.Initiate mCompManClient.SRVC_EXPORT_CHANGED, e_wbk_serviced, e_hosted
    If Services.Denied(mCompManClient.SRVC_EXPORT_CHANGED) Then GoTo xt
        
    Services.ExportChangedComponents e_hosted
    ExportChangedComponents = True
    ExportChangedComponents = Application.StatusBar
    
xt: mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Services.LogEntrySummary Application.StatusBar
    Set Services = Nothing
    Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub MsgInit()
    Dim i As Long
    
    For i = 1 To mMsg.NoOfMsgSects
        With Msg.Section(i)
            .Label.Text = vbNullString
            .Label.FontColor = rgbBlue
            .Text.Text = vbNullString
            .Text.MonoSpaced = False
        End With
    Next i

End Sub

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        mBasic.ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        mBasic.ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Public Function RunTest(ByVal r_service_proc As String, _
                        ByRef r_serviced_wbk As Workbook) As Variant
' ----------------------------------------------------------------------------
' Ensures the requested service (r_service_proc) is able to run or returns the
' reason why not. The function returns:
' AppErr(1): CompMan's current configuration does not support the requested
'            service, i.e. either the 'Synchronization-Folder' for the
'            'Synchronization' service is invalid or the 'Servicing-Folder'
'            for all other services is invalid.
' AppErr(2): The servicing Workbook is the available Add-in but the Add-in is
'            currently paused.
'            It requires the CompMan Workbook (mCompManClient.COMPMAN_DEVLP)
'            to turn the Add-in to a status continued. When the CompMan
'            Workbook (mCompManClient.COMPMAN_DEVLP) is open, it will provide
'            the service provided it is not also the serviced Workbook.
' AppErr(3): For the Update Outdated service the Addin is not available (not
'            open or opene but paused).
' AppErr(4): For the VBProject-Synchronization servicethe CompMan devinstance
'            Workbook is not open.
' ----------------------------------------------------------------------------
    Const PROC = "RunTest"
    
    On Error GoTo eh
    
    RunTest = 0
    Select Case r_service_proc
        Case mCompManClient.SRVC_UPDATE_OUTDATED, mCompManClient.SRVC_EXPORT_CHANGED
            Select Case True
                Case Not wsConfig.FolderCompManRootIsValid:                             RunTest = mBasic.AppErr(1) ' Configuration for the service is invalid
                Case Not r_serviced_wbk.FullName Like wsConfig.FolderCompManRoot & "*": RunTest = mBasic.AppErr(2) ' Denied because outside configured 'Dev and Test' folder
                Case r_service_proc = mCompManClient.SRVC_UPDATE_OUTDATED
                    If mMe.IsDevInstnc And ((mAddin.IsOpen And mAddin.Paused) Or Not mAddin.IsOpen) Then
                        RunTest = mBasic.AppErr(3) ' Denied because serviced is the DevInstance but the Addin is paused or not open
                    End If
            End Select
        
        Case mCompManClient.SRVC_SYNCHRONIZE
            If wsConfig.FolderSyncTarget = vbNullString Or wsConfig.FolderSyncArchive = vbNullString Then
                RunTest = mBasic.AppErr(1) ' Not configured
            ElseIf Not r_serviced_wbk.FullName Like wsConfig.FolderSyncTarget & "*" Then
                RunTest = mBasic.AppErr(2) ' Denied because not opened from within the configured 'Sync-Target' folder
            ElseIf Not mMe.IsDevInstnc Then
                RunTest = mBasic.AppErr(3)
            End If
    End Select

xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SynchronizeVBProjects(ByVal sync_wbk_opened As Workbook)
' ----------------------------------------------------------------------------
' Initiates the synchronization of the opened Workbook's (sync_wbk_opened)
' VB-Project - which is considered the Sync-Target-Workbook - with the
' VB-Project of the corresponding Sync-Source-Workbook - which is the Workbook
' with the same name located in CompMan's 'ServicedDecAndTest' folder. The
' service is performed provided:
' - the Workbook (sync_wbk_opened) is open/ed from within the configured
'   'SynchronizeTarget' folder. The Workbook is initially opened by its origin
'   name but immediately saved as the Sync-Target-Workbook's working copy
' - a corresponding Workbook (not open!) is located in CompMan's configured
'   'ServicedCompManRoot' folder
' - CompMan's synchronization service is available. i.e. the opened Workbook
'   is able to be served either by the CompMan development instance or by the
'   Add-in instance
' ----------------------------------------------------------------------------
    Const PROC = "SynchronizeVBProjects"
            
    On Error GoTo eh
    Set Services = New clsServices
    With Services
        .ServicedWbk = sync_wbk_opened
        .CurrentService = mCompManClient.SRVC_SYNCHRONIZE
        .EstablishExecTraceFile
    End With
    
    mBasic.BoP ErrSrc(PROC)
    Services.Initiate mCompManClient.SRVC_SYNCHRONIZE, sync_wbk_opened, vbNullString
    
    If mSync.SourceExists(sync_wbk_opened) Then
        '~~ - Keep records of the full name of all three Workbooks involved in this synchronization
        '~~   derived from the Workbook opened which may be the Sync-Target-Workbook or the
        '~~   Sync-Target-Worlkbook's working copy and
        '~~ - Display an open decision dialog
        wsService.SyncTargetFullNameCopy = mSync.TargetWorkingCopyFullName(sync_wbk_opened)
        mSync.OpenDecision ' Display mode-less open decision dialog
    Else
        sync_wbk_opened.Close False
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Set Services = Nothing
    Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub UpdateOutdatedCommonComponents(ByRef u_wbk_serviced As Workbook, _
                                 Optional ByVal u_hosted As String = vbNullString)
' ------------------------------------------------------------------------------
' Presents the serviced Workbook's outdated components in a modeless dialog with
' two buttons for each component. One button executes Application.Run mRenew.Run
' for a component to update it, the other executes Application.Run
' Services.ExpFilesDiffDisplay to display the code changes.
' Note: u_unused is for backwards compatibility only.
'
' Precondition: The service has been checked by the client to be able to run.
' ------------------------------------------------------------------------------
    Const PROC = "UpdateOutdatedCommonComponents"
    
    On Error GoTo eh
'    mCompManClient.Events ErrSrc(PROC), False
'    '~~ When the Open event is raised through the VBE Immediate Window or when
'    '~~ the open houskeeping has registered new used Common Components ...
'    Set Services = Nothing
'    mCompMan.ExportChangedComponents u_wbk_serviced, u_hosted
    
    Set Services = Nothing
    Set Services = New clsServices
    With Services
        .ServicedWbk = u_wbk_serviced
        .CurrentService = mCompManClient.SRVC_UPDATE_OUTDATED
        .EstablishExecTraceFile
    End With
    
    mBasic.BoP ErrSrc(PROC)
    Services.Initiate mCompManClient.SRVC_UPDATE_OUTDATED, u_wbk_serviced, u_hosted
    CommComps.Hosted = u_hosted
    CompManDat.Hosted = u_hosted
    Set Qoutdated = Nothing
    Set Prgrss = New clsProgress
    With Prgrss
        .Figures = True
        .DoneItemsInfo = True
    End With
    
    mCommComps.OutdatedUpdate ' Dialog to update/renew one by one
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function UTC() As String
    Dim dt As Object
    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate Now
    UTC = Format(dt.GetVarDate(False), "YYYY-MM-DD-hh-mm-ss") & " (UTC)"
End Function

Public Function WbkGetOpen(ByVal go_wbk_full_name As String) As Workbook
' ----------------------------------------------------------------------------
' Returns an opened Workbook object named (go_wbk_full_name) or Nothing when a
' file named (go_wbk_full_name) not exists.
' ----------------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    
    If fso.FileExists(go_wbk_full_name) Then
        If mCompMan.WbkIsOpen(io_name:=go_wbk_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wbk_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wbk_full_name)
    End If
    
xt: Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsOpen(Optional ByVal io_name As String = vbNullString, _
                           Optional ByVal io_full_name As String) As Boolean
' ----------------------------------------------------------------------------
' When the full name is provided the check spans all Excel instances else only
' the current one.
' ----------------------------------------------------------------------------
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim xlApp   As Excel.Application
    
    If io_name = vbNullString And io_full_name = vbNullString Then GoTo xt
    
    If io_full_name <> vbNullString Then
        '~~ With the full name the open test spans all application instances
        If Not fso.FileExists(io_full_name) Then GoTo xt
        If io_name = vbNullString Then io_name = fso.GetFileName(io_full_name)
        On Error Resume Next
        Set xlApp = GetObject(io_full_name).Application
        WbkIsOpen = Err.Number = 0
    Else
        On Error Resume Next
        io_name = Application.Workbooks(io_name).Name
        WbkIsOpen = Err.Number = 0
    End If

xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function WinMergeIsInstalled() As Boolean
    WinMergeIsInstalled = AppIsInstalled("WinMerge")
End Function

