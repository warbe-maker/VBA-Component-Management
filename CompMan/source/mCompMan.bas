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
' Values in CompMan.dat, Pending.dat, and ComComp.dat
Public Const VALUE_NAME_LAST_MOD_EXP_FILE_ORIG  As String = "LastModExpFileOrigin"
Public Const VALUE_NAME_LAST_MOD_AT             As String = "LastModAt"
Public Const VALUE_NAME_LAST_MOD_BY             As String = "LastModBy"
Public Const VALUE_NAME_LAST_MOD_IN             As String = "LastModIn"
Public Const VALUE_NAME_LAST_MOD_ON             As String = "LastModOn"

Public CommCompsPendingRelease                  As Dictionary
Public CommonPending                            As clsCommonPending
Public CommonPublic                             As clsCommonPublic
Public CompManDat                               As clsCommonServiced    ' Serviced Workbook's Common Components Private Profile file
Public Comps                                    As clsComps
Public Environment                              As clsEnvironment       ' Initializes, forwards, sets up CompMan's environment
Public LogServiced                              As clsLog               ' log writen for the serviced Workbook
Public LogServicesSummary                       As clsLog               ' the servicing Workbooks own log
Public Msg                                      As udtMsg
Public Prgrss                                   As clsProgress
Public Serviced                                 As New clsServiced
Public Services                                 As clsServices
#If clsTrc = 1 Then
    Public Trc                                  As clsTrc
#End If

Public Enum enCompManService
    enExportChanged
    enSynchVBProjects
    enUpdateOutdatedCommComps
    enReleasePending
End Enum

Public Enum enKindOfComp    ' Kind of VBComponent in the serviced Workbook
    enCompInternal        ' Neither of the below, i.e. not known as a Common Component
    enCompCommonHosted    ' Common Component claimed hosted by the serviced Workbook
    enCompCommonUsed      ' Public Common Component used by the serviced Workbook's VBProject
    enCompCommonPrivate   ' Common Component explicitely registered private (accidentially same name)
End Enum

Public Enum siCounter
    sic_comps
    sic_comps_changed
    sic_used_comm_vbc_outdated
End Enum

Private Type SYSTEM_TIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEM_TIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEM_TIME
    DaylightBias As Long
End Type

Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpLocalTime As SYSTEM_TIME, lpUniversalTime As SYSTEM_TIME) As Integer

Private sLogFileService         As String
Private sLogFileServicedSummary As String
Private sExecTraceFile          As String
Private sCurrentServiceName     As String
Private wbkServiced             As Workbook
Private sCommCompsServicedPrivProfFileFullName  As String

Public Property Get CommCompsServicedPrivProfFileFullName() As String:          CommCompsServicedPrivProfFileFullName = sCommCompsServicedPrivProfFileFullName: End Property

Public Property Let CommCompsServicedPrivProfFileFullName(ByVal s As String):   sCommCompsServicedPrivProfFileFullName = s:                                     End Property

Public Property Get ServicedWrkbk() As Workbook:                                Set ServicedWrkbk = wbkServiced:                                                End Property

Public Property Let ServicedWrkbk(ByVal wbk As Workbook):                       Set wbkServiced = wbk:                                                          End Property

Public Property Get LogFileService() As String:                                 LogFileService = sLogFileService:                                               End Property

Public Property Let LogFileService(ByVal s As String):                          sLogFileService = s:                                                            End Property

Public Property Get LogFileServicesSummary() As String:                         LogFileServicesSummary = sLogFileServicedSummary:                               End Property

Public Property Let LogFileServicesSummary(ByVal s As String):                  sLogFileServicedSummary = s:                                                    End Property

Public Function AboutCommComps() As String
    AboutCommComps = "A ""Common Component"" is one of which the Export-File exists in the ""Common-Components"" folder. It may be modified " & _
                     "in any Workbook using it provided it is up-to-date and there's yet no modification made in another Workbook which " & _
                     "is still ""pending release"". When a ""Common Component"" is modified it becomes ""pending release"" and the modification " & _
                     "becomes ""public"" when it is ""released"" by means of the ""CompMan"" menu in the VBE."
End Function

Public Sub CheckForUnusedPublicItems()
' ----------------------------------------------------------------
' Attention! The service requires the VBPUnusedPublic.xlsb
'            Workbook open. When not open the service terminates
'            without notice.
' ----------------------------------------------------------------
    Const COMPS_EXCLUDED = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"
    Const LINES_EXCLUDED = "Select Case mBasic.ErrMsg(ErrSrc(PROC))" & vbLf & _
                           "Case vbResume:*Stop:*Resume" & vbLf & _
                           "Case Else:*GoTo xt"
    On Error Resume Next
    '~~ Providing all (optional) arguments saves the Workbook selection dialog and the VBComponents selection dialog
    Application.Run "VBPUnusedPublic.xlsb!mUnused.Unused", ThisWorkbook, COMPS_EXCLUDED, LINES_EXCLUDED

End Sub

Public Function CurrentServiceStatusBar() As String
' ------------------------------------------------------------------------------
' Returns the current services base status bar message.
' ------------------------------------------------------------------------------
    CurrentServiceStatusBar = mCompMan.CurrentServiceName & " (by "
    If ThisWorkbook.Name = mAddin.WbkName _
    Then CurrentServiceStatusBar = CurrentServiceStatusBar & "Add-in" _
    Else CurrentServiceStatusBar = CurrentServiceStatusBar & ThisWorkbook.Name
    CurrentServiceStatusBar = CurrentServiceStatusBar & ") for " & ActiveWorkbook.Name

End Function

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCompMan" & "." & es_proc
End Function

Public Function ExportChangedComponents(ByRef e_wbk_serviced As Workbook, _
                               Optional ByVal e_hosted As String = vbNullString, _
                               Optional ByVal e_public_proc_copies As String = vbNullString) As Variant
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
    
    mCompMan.CurrentServiceName = mCompManClient.SRVC_EXPORT_CHANGED_DSPLY
    mCompMan.ServicedWrkbk = e_wbk_serviced
    Set Environment = New clsEnvironment
    
    mBasic.BoP ErrSrc(PROC)
    mCompMan.ServiceInitiate s_serviced_wbk:=e_wbk_serviced _
                           , s_service:=mCompManClient.SRVC_EXPORT_CHANGED _
                           , s_hosted:=e_hosted _
                           , s_public_proc_copies:=e_public_proc_copies
        
    Services.Initiate mCompManClient.SRVC_EXPORT_CHANGED, e_wbk_serviced, e_hosted
    If Services.Denied(mCompManClient.SRVC_EXPORT_CHANGED) Then GoTo xt
        
    Services.ExportChangedComponents
    ExportChangedComponents = True
    ExportChangedComponents = Application.StatusBar
    
xt: mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Services.LogEntrySummary Application.StatusBar
    CompMan.ServiceTerminate
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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

Public Sub README(Optional ByVal r_url_subject As String = "/blob/master/README.md", _
                  Optional ByVal r_url_bookmark As String = vbNullString)
    
    If r_url_bookmark <> vbNullString _
    Then r_url_bookmark = Replace("#" & r_url_bookmark, "##", "#") ' add # if missing

    mBasic.ShellRun GITHUB_REPO_URL & r_url_subject & r_url_bookmark
    
End Sub

Public Function RegState(ByVal r_state As Variant) As Variant
' ---------------------------------------------------------------------------
' Returns a transformed registration state (r_state):
' A numeric registration state (r_state) into a string and vice versa.
' ---------------------------------------------------------------------------

    If TypeName(r_state) = "String" Then
        Select Case r_state
            Case "hosted":                  RegState = enCompCommonHosted
            Case "private":                 RegState = enCompCommonPrivate
            Case "used":                    RegState = enCompCommonUsed
            Case vbNullString:              RegState = enCompInternal
        End Select
    Else
        Select Case r_state
            Case enCompCommonHosted:          RegState = "hosted"
            Case enCompCommonPrivate:         RegState = "private"
            Case enCompCommonUsed:            RegState = "used"
            Case enCompInternal:              RegState = "internal"
        End Select
    End If
    
End Function

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
                Case Not wsConfig.FolderCompManRootIsValid
                    RunTest = mBasic.AppErr(1) ' Configuration for the service is invalid
                    Debug.Print ErrSrc(PROC) & ": " & ErrSrc(PROC) & ": Configuration invalid"
                Case Not r_serviced_wbk.FullName Like wsConfig.FolderCompManRoot & "*"
                    RunTest = mBasic.AppErr(2) ' Denied because outside configured 'Dev and Test' folder
                    Debug.Print ErrSrc(PROC) & ": " & ErrSrc(PROC) & ": Service denied because out of serviced folder"
                Case r_service_proc = mCompManClient.SRVC_UPDATE_OUTDATED
                    Select Case True
                        Case r_serviced_wbk.Name <> ThisWorkbook.Name
                        Case mMe.IsDevInstnc And ((mAddin.IsOpen And mAddin.Paused) Or Not mAddin.IsOpen)
                            RunTest = mBasic.AppErr(3) ' Denied because serviced is the DevInstance but the Addin is paused or not open
                            Debug.Print ErrSrc(PROC) & ": " & ErrSrc(PROC) & ": Service denied because the Addin is paused"
                    End Select
            End Select
        
        Case mCompManClient.SRVC_SYNCHRONIZE
            If wsConfig.FolderSyncTarget = vbNullString Or wsConfig.FolderSyncArchive = vbNullString Then
                RunTest = mBasic.AppErr(1) ' Not configured
                Debug.Print ErrSrc(PROC) & ": " & ErrSrc(PROC) & ": Service denied because not configured"
            ElseIf Not r_serviced_wbk.FullName Like wsConfig.FolderSyncTarget & "*" Then
                RunTest = mBasic.AppErr(2) ' Denied because not opened from within the configured 'Sync-Target' folder
                Debug.Print ErrSrc(PROC) & ": " & ErrSrc(PROC) & ": Service denied because not opened from within the serviced folder"
            ElseIf Not mMe.IsDevInstnc Then
                RunTest = mBasic.AppErr(3)
                Debug.Print ErrSrc(PROC) & ": " & ErrSrc(PROC) & ": Service denied because not applied by Dev instance"
            End If
    End Select

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Service(ByVal s_service As Variant) As Variant
' ---------------------------------------------------------------------------
' Converts an enumerated service into a string and a string into an
' enumerated service.
' ---------------------------------------------------------------------------
    If TypeName(s_service) = "String" Then
        Select Case s_service
            Case mCompManClient.SRVC_EXPORT_CHANGED:            Service = enCompManService.enExportChanged
            Case mCompManClient.SRVC_RELEASE_PENDING:           Service = enCompManService.enReleasePending
            Case mCompManClient.SRVC_SYNCHRONIZE:               Service = enCompManService.enSynchVBProjects
            Case mCompManClient.SRVC_UPDATE_OUTDATED:           Service = enCompManService.enUpdateOutdatedCommComps
            Case Else:                                          Service = s_service
        End Select
    Else
        Select Case s_service
            Case enCompManService.enExportChanged:              Service = mCompManClient.SRVC_EXPORT_CHANGED
            Case enCompManService.enReleasePending:             Service = mCompManClient.SRVC_RELEASE_PENDING
            Case enCompManService.enSynchVBProjects:            Service = mCompManClient.SRVC_SYNCHRONIZE
            Case enCompManService.enUpdateOutdatedCommComps:    Service = mCompManClient.SRVC_UPDATE_OUTDATED
        End Select
    End If
    
End Function

Public Property Get CurrentServiceName() As String:             CurrentServiceName = sCurrentServiceName:           End Property

Public Property Let CurrentServiceName(ByVal s As String):      sCurrentServiceName = s:                            End Property

Public Sub ServiceInitiate(ByVal s_serviced_wbk As Workbook, _
                           ByVal s_service As String, _
                  Optional ByVal s_hosted As String, _
                  Optional ByVal s_public_proc_copies As String, _
                  Optional ByVal s_do_housekeeping As Boolean = True)
' ----------------------------------------------------------------------------
' Establishes all resources for any CompMan service.
' Note: There is a strong dependency of procedures and thus the sequence in
'       which resuources are established matters!
' ----------------------------------------------------------------------------
    Const PROC = "ServiceInitiate"
    
    On Error GoTo eh
    Dim lMaxLenType As Long
    Dim lMaxLenItem As Long
    
    
    mBasic.BoP ErrSrc(PROC)
    Set Services = New clsServices
    With Services
        .CurrentService = mCompMan.Service(s_service)
        .ServicedWbk = s_serviced_wbk
    End With
    
    Set Serviced = New clsServiced
    Serviced.Wrkbk = s_serviced_wbk
             
    With Serviced
        .HostedCommComps = s_hosted
        .ServiceName = s_service
        .PublProcCpys = s_public_proc_copies
    End With
    
    Set CommonPublic = New clsCommonPublic
    Set Prgrss = New clsProgress
    
    Set CommonPending = New clsCommonPending
    Set CommCompsPendingRelease = CommonPending.Components
    
    Serviced.MaxLengths
    lMaxLenType = Serviced.MaxLenType
    lMaxLenItem = Serviced.MaxLenItem
    Environment.EstablishServicesLog lMaxLenType, lMaxLenItem
    Environment.EstablishServicesSummaryLog lMaxLenType, lMaxLenItem
    
    If s_do_housekeeping Then mHskpng.CommComps
    

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ServicesFolder() As String
    Const PROC = "ServicedFolder"
    
    Dim s As String
    
    If Serviced Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The ServicesFolder requires the known serviced Workbook, which is not the case yet!"
    s = ActiveWorkbook.Path & "\CompMan"
    If Not FSo.FolderExists(s) Then FSo.CreateFolder s
    ServicesFolder = s
    
End Function

Public Sub ServiceTerminate()

    Dim i   As Long
    Dim s   As String
    
    Set CommCompsPendingRelease = Nothing
    Set CommonPending = Nothing
    Set CommonPublic = Nothing
    Set CompManDat = Nothing
    Set Comps = Nothing
    Set LogServiced = Nothing
    Set LogServicesSummary = Nothing
    Set Serviced = Nothing
    Set Services = Nothing
    Set Prgrss = Nothing
    
    s = Application.StatusBar
    For i = 5 To 0 Step -1
        mBasic.DelayedAction 0.6
        Application.StatusBar = s & " " & String(i, " ") & i
    Next i
    Application.StatusBar = " "
    
#If mTrc = 1 Then
    mTrc.Terminate
#ElseIf clsTrc = 1 Then
    Set Trc = Nothing
#End If

End Sub

Private Function SystemTimeToVBTime(systemTime As SYSTEM_TIME) As Date
    With systemTime
        SystemTimeToVBTime = DateSerial(.wYear, .wMonth, .wDay) + _
                TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Sub UpdateOutdatedCommonComponents(ByRef u_wbk_serviced As Workbook, _
                                 Optional ByVal u_hosted As String = vbNullString, _
                                 Optional ByVal u_public_proc_copies As String = vbNullString)
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
    
    CurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY
    mCompMan.ServicedWrkbk = u_wbk_serviced
    Set Environment = New clsEnvironment

    mBasic.BoP ErrSrc(PROC)
    mCompMan.ServiceInitiate s_serviced_wbk:=u_wbk_serviced _
                           , s_service:=mCompManClient.SRVC_UPDATE_OUTDATED _
                           , s_hosted:=u_hosted _
                           , s_public_proc_copies:=u_public_proc_copies
    
    mBasic.BoP ErrSrc(PROC)
    Services.Initiate mCompManClient.SRVC_UPDATE_OUTDATED, u_wbk_serviced, u_hosted
    With Prgrss
        .Operation = "Update outdated Common Components"
        .Figures = True
        .DoneItemsInfo = True
    End With
    mCommComps.Update ' Dialog to update/renew one by one
        
xt: mBasic.EoP ErrSrc(PROC)
    mCompMan.ServiceTerminate
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function UTC(Optional ByVal u_time As Variant = 0) As String
' ------------------------------------------------------------------------------
' Returns the UTC of a given date-time or Now().
' ------------------------------------------------------------------------------
    Dim timeZoneInfo    As TIME_ZONE_INFORMATION
    Dim utcSystemTime   As SYSTEM_TIME
    Dim localSystemTime As SYSTEM_TIME
    Dim utcResult       As Date
    
    GetTimeZoneInformation timeZoneInfo

    If u_time = 0 Then u_time = Now()
    With localSystemTime
        .wYear = Year(u_time)
        .wMonth = Month(u_time)
        .wDay = Day(u_time)
        .wHour = Hour(u_time)
        .wMinute = Minute(u_time)
        .wSecond = Second(u_time)
    End With

    If TzSpecificLocalTimeToSystemTime(timeZoneInfo, localSystemTime, utcSystemTime) <> 0 Then
        utcResult = SystemTimeToVBTime(utcSystemTime)
        UTC = Format(utcResult, "YYYY-MM-DD hh:mm:ss") & " (UTC)"
    Else
        Err.Raise 1, "WINAPI", "Windows API call failed"
    End If

End Function

Public Function WbkGetOpen(ByVal go_wbk_full_name As String) As Workbook
' ----------------------------------------------------------------------------
' Returns an opened Workbook object named (go_wbk_full_name) or Nothing when a
' file named (go_wbk_full_name) not exists.
' ----------------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    
    If FSo.FileExists(go_wbk_full_name) Then
        If mCompMan.WbkIsOpen(io_name:=go_wbk_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wbk_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wbk_full_name)
    End If
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
        If Not FSo.FileExists(io_full_name) Then GoTo xt
        If io_name = vbNullString Then io_name = FSo.GetFileName(io_full_name)
        On Error Resume Next
        Set xlApp = GetObject(io_full_name).Application
        WbkIsOpen = Err.Number = 0
    Else
        On Error Resume Next
        io_name = Application.Workbooks(io_name).Name
        WbkIsOpen = Err.Number = 0
    End If

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function WinMergeIsInstalled() As Boolean
    WinMergeIsInstalled = AppIsInstalled("WinMerge")
End Function

