Attribute VB_Name = "mCompMan"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard Module mCompMan
'          Services for the management of VBComponents in Workbooks provided:
'          - stored within the 'FolderServiced'
'          - have the Conditional Compile Argument 'CompMan = 1'
'          - have 'CompMan' referenced
'          - the Workbook resides in its own dedicated folder
'          - the Workbook calls the '' service with the Open event
'          - the Workbook calls the '' service with the Save event
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
' Services:
' - DisplayCodeChange       Displays the current difference between a
'                           component's code and its current Export-File
' - ExportAll               Exports all components into the Workbook's
'                           dedicated folder (created when not existing)
' - ExportChangedComponents Exports all components of which the code in the
'                           Export-File differs from the current code.
' - Service
'
' Uses Common Components: - mBasic
'                         - mErrhndlr
'                         - mFso
'                         - mWrkbk
' Requires:
' - Reference to: - "Microsoft Visual Basic for Applications Extensibility ..."
'                 - "Microsoft Scripting Runtime"
'                 - "Windows Script Host Object Model"
'                 - Trust in the VBA project object modell (Security
'                   setting for makros)
'
' W. Rauschenberger Berlin August 2019
' -------------------------------------------------------------------------------
Public Const MAX_LEN_TYPE                   As Long = 17
Public Const README_URL                     As String = "https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md"

Public Enum enKindOfComp            ' The kind of VBComponent in the sense of CompMan
    enUnknown = 0
    enCommCompHosted = 1
    enCommCompUsed = 2              ' The Component is a used raw, i.e. the raw is hosted by another Workbook
    enInternal = 3                  ' Neither a hosted nor a used Raw Common Component
End Enum

Public Enum siCounter
    sic_clone_changed
    sic_used_comm_comps
    sic_cols_new
    sic_cols_obsolete
    sic_comps
    sic_comps_changed
    sic_comps_total
    sic_named_ranges
    sic_named_ranges_total
    sic_names_new
    sic_names_obsolete
    sic_names_total
    sic_non_doc_mods_code
    sic_non_doc_mod_new
    sic_non_doc_mod_obsolete
    sic_non_doc_mod_total
    sic_oobs_new
    sic_oobs_obsolete
    sic_oobs_total
    sic_raw_comm_vbc_changed
    sic_used_comm_vbc_Outdated
    sic_used_comm_vbc_updated
    sic_refs_new
    sic_refs_obsolete
    sic_refs_total
    sic_rows_new
    sic_rows_obsolete
    sic_sheet_controls_new
    sic_sheet_controls_obsolete
    sic_sheet_controls_total
    sic_shape_properties
    sic_sheets_code
    sic_sheets_codename
    sic_sheets_name
    sic_sheets_new
    sic_sheets_obsolete
    sic_sheets_total
End Enum

Public lMaxCompLength   As Long
Public dctHostedRaws    As Dictionary
Public Stats            As clsStats
Public Log              As clsLog

    
Public Property Get HostedRaws() As Variant:           Set HostedRaws = dctHostedRaws:                 End Property

Public Property Let HostedRaws(ByVal hr As Variant)
' ----------------------------------------------------------------------------
' Saves the names of the components claimed 'hosted raw components' (hr) to
' the Dictionary (dctHostedRaws).
' ----------------------------------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    
    If dctHostedRaws Is Nothing Then
        Set dctHostedRaws = New Dictionary
    Else
        dctHostedRaws.RemoveAll
    End If
    For Each v In Split(hr, ",")
        sComp = Trim$(v)
        If Not dctHostedRaws.Exists(sComp) Then
            dctHostedRaws.Add sComp, sComp
        End If
    Next v
    
End Property

Public Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub DisplayChanges(Optional ByVal fl_1 As String = vbNullString, _
                          Optional ByVal fl_2 As String = vbNullString)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "DisplayChanges"
    
    On Error GoTo eh
    Dim fl  As File
    Dim fso As New FileSystemObject
    
    If fl_1 = vbNullString Then
        mFso.FilePicked p_title:="Select the file regarded 'before the change' (displayed at the left)!" _
                      , p_file:=fl
        If Not fl Is Nothing Then fl_1 = fl.Path
    End If
    
    If fl_2 = vbNullString Then
        mFso.FilePicked p_title:="Select the file regarded 'the changed one' (displayed at the right)!" _
                   , p_file:=fl
        If Not fl Is Nothing Then fl_2 = fl.Path
    End If
    
    If Not fso.FileExists(fl_1) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid file specification provided with argument fl_1 or no fiie selected for fl_1!"
    If Not fso.FileExists(fl_2) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid file specification provided with argument fl_2 or no fiie selected for fl_2!"
                            
    mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=fl_1 _
                                   , fd_exp_file_left_title:=fl_1 & " ( b e f o r e  the changes)" _
                                   , fd_exp_file_right_full_name:=fl_2 _
                                   , fd_exp_file_right_title:=fl_2 & " ( a f t e r  the changes)"

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCompMan" & "." & es_proc
End Function

Public Sub EstablishExecTraceFile(ByVal etl_wbk_serviced As Workbook, _
                        Optional ByVal etl_append As Boolean = False)
' --------------------------------------------------------------------------
' Establishes a trace log file in the serviced Workbook's parent folder
' provided the Conditional Compile Argument ExecTrace = 1.
' --------------------------------------------------------------------------
#If ExecTrace = 1 Then
    
    Dim sFile As String
    sFile = Replace(etl_wbk_serviced.FullName, etl_wbk_serviced.name, "CompMan.Service.trc")

    '~~ Even when etl_append = False: When the file had been createde today etl_append will be set to True
    With New FileSystemObject
        If .FileExists(sFile) Then
            If Format(.GetFile(sFile).DateCreated, "YYYY-MM-DD") = Format(Now(), "YYYY-MM-DD") Then
                etl_append = True
            End If
        End If
    End With
    mTrc.LogFile(tl_append:=etl_append) = sFile
    mTrc.LogTitle = Log.Service
#End If
End Sub

Public Sub ExportAll(Optional ByRef ea_wbk_serviced As Workbook = Nothing)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    
    If Log Is Nothing Then Set Log = New clsLog
    Log.Service = "Export All"
    EstablishExecTraceFile ea_wbk_serviced
    
    mBasic.BoP ErrSrc(PROC)
    If ea_wbk_serviced Is Nothing _
    Then mService.WbkServiced = ActiveWorkbook _
    Else mService.WbkServiced = ea_wbk_serviced
    mExport.All
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub UpdateOutdatedCommonComponents(ByRef uo_wbk_serviced As Workbook, _
                                 Optional ByVal uo_hosted As String = vbNullString, _
                                 Optional ByVal uo_unused As Boolean)
' ------------------------------------------------------------------------------
' Presents the serviced Workbook's outdated components in a modeless dialog with
' two buttons for each component. One button executes Application.Run mRenew.Run
' for a component to update it, the other executes Application.Run
' mService.ExpFilesDiffDisplay to display the code changes.
' Note: uo_unused is for backwards compatibility only.
'
' Precondition: The service has been checked by the client to be able to run.
' ------------------------------------------------------------------------------
    Const PROC = "UpdateOutdatedCommonComponents"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim i           As Long
    Dim wbServiced  As Workbook
    Dim RawExpFile  As String
    Dim fUpdate     As fMsg
    Dim Msg         As TypeMsg
    Dim sTitle      As String
    Dim cll         As Collection
    Dim dctRunArgs  As Dictionary
    Dim RenewComp   As String
    Dim Comps       As New clsComps
    Dim cllBttns    As New Collection
    Dim sBttn1      As String
    Dim sBttn2      As String
    Dim Comp        As clsComp
    
    wsService.ClearDataAllServices
    wsService.ClearDataUpdateService
    wsService.ServicedWorkbookFullName = uo_wbk_serviced.FullName
    mService.WbkServiced = uo_wbk_serviced
    Set Log = New clsLog
    Log.Service(new_log:=True) = mCompManClient.SRVC_UPDATE_OUTDATED
    wsService.LogFileFullName = Log.FileFullName
    wsService.ServicedItemsMaxLenName = 0
    wsService.ServicedItemsMaxLenType = 0
    EstablishExecTraceFile uo_wbk_serviced
    mService.DsplyStatus Log.Service
    mBasic.BoP ErrSrc(PROC)
    mComCompRawsHosted.Manage uo_hosted
    
    mOutdated.Display uo_hosted ' Dialog to update/renew one by one
                     
xt: mBasic.EoP ErrSrc(PROC)
    Set Log = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ExportChangedComponents(ByRef ec_wbk_serviced As Workbook, _
                               Optional ByVal ec_hosted As String = vbNullString) As Variant
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
    wsService.ClearDataAllServices
    wsService.ServicedWorkbookFullName = ec_wbk_serviced.FullName
    mService.WbkServiced = ec_wbk_serviced
    Set Log = New clsLog
    Log.Service = mCompManClient.SRVC_EXPORT_CHANGED
    
    EstablishExecTraceFile ec_wbk_serviced
    mBasic.BoP ErrSrc(PROC)
    
    If mService.Denied Then GoTo xt
    mService.ExportChangedComponents ec_hosted
    ExportChangedComponents = True
    ExportChangedComponents = Application.StatusBar
    
xt: Application.EnableEvents = True
    mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Install()
    mService.Install ActiveWorkbook
End Sub

Public Function RunTest(ByVal rt_service As String, _
                        ByRef rt_serviced_wbk As Workbook) As Variant
' --------------------------------------------------------------------------
' Ensures the requested service (rt_service) is able to run or returns the
' reason why not. The function returns:
' AppErr(1): CompMan's current configuration does not support the requested
'            service, i.e. either the 'Synchronization-Folder' for the
'            'Synchronization' service is invalid or the 'Servicing-Folder'
'            for all other services is invalid.
' AppErr(2): The servicing Workbook is the available Add-in but the Add-in
'            is currently paused.
'            It requires the CompMan Workbook (mCompManClient.COMPMAN_DEVLP)
'            to turn the Add-in to a status continued. When the CompMan
'            Workbook (mCompManClient.COMPMAN_DEVLP) is open, it will provide
'            the service provided it is not also the serviced Workbook.
' ----------------------------------------------------------------------------
    Const PROC = "RunTest"
    
    On Error GoTo eh
    
    Select Case True
        Case rt_service = mCompManClient.SRVC_UPDATE_OUTDATED _
         And mMe.IsDevInstnc _
         And ((mAddin.IsOpen And mAddin.Paused) Or Not mAddin.IsOpen)
            RunTest = AppErr(3)
        
        Case rt_service = mCompManClient.SRVC_SYNCHRONIZE _
         And (Not wsConfig.FolderSyncTargetIsValid Or Not wsConfig.FolderSyncArchiveIsValid)
            RunTest = AppErr(1) ' The preconditions for the VB-Project Synchronization Service are not met
                    
        Case (rt_service = mCompManClient.SRVC_UPDATE_OUTDATED Or rt_service = mCompManClient.SRVC_EXPORT_CHANGED) _
         And Not wsConfig.FolderDevAndTestIsValid
            RunTest = AppErr(1) ' The serviced root folder is invalid (not configured or not existing)
    
        Case rt_service = mCompManClient.SRVC_SYNCHRONIZE _
         And Not rt_serviced_wbk.FullName Like wsConfig.FolderSyncTarget & "*"
            RunTest = AppErr(4)
        
        Case Else
            RunTest = 0
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function SyncSourceInDevFolder(ByVal ss_serviced As Workbook) As Boolean
    Stop ' impl pending

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
'   'ServicedDevAndTest' folder
' - CompMan's synchronization service is available. i.e. the opened Workbook
'   is able to be served either by the CompMan development instance or by the
'   Add-in instance
' ----------------------------------------------------------------------------
    Const PROC = "SynchronizeVBProjects"
    
    Dim cllResultFiles          As Collection
    Dim sSyncWbTargetWrkngCpy   As String
    Dim sSyncWbSourceFullName   As String
    Dim SyncOpenMsg             As TypeMsg
    Dim wbk                     As Workbook
    Dim sResult                 As String
        
    mService.WbkServiced = sync_wbk_opened
#If ExecTrace = 1 Then
    mTrc.LogFile = Replace(sync_wbk_opened.FullName, sync_wbk_opened.name, "Exec.trc")
#End If
    mBasic.BoP ErrSrc(PROC)
    wsService.ClearDataAllServices
    
    If mSync.SourceExists(sync_wbk_opened) Then
        '~~ - Keep records of the full name of all three Workbooks involved in this synchronization
        '~~   derived from the Workbook opened which may be the Sync-Target-Workbook or the
        '~~   Sync-Target-Worlkbook's working copy and
        '~~ - Display an open decision dialog
        wsService.ServicedWorkbookFullName = mSync.TargetOriginFullName(sync_wbk_opened)
        wsService.SyncTargetFullNameCopy = mSync.TargetCopyFullName(sync_wbk_opened)
        mSync.OpenDecision ' Display mode-less open decision dialog
    Else
        sync_wbk_opened.Close False
    End If
    
xt: Set Log = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


Public Function WbkGetOpen(ByVal go_wbk_full_name As String) As Workbook
' ----------------------------------------------------------------------------
' Returns an opened Workbook object named (go_wbk_full_name) or Nothing when a
' file named (go_wbk_full_name) not exists.
' ----------------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If fso.FileExists(go_wbk_full_name) Then
        If mCompMan.WbkIsOpen(io_name:=go_wbk_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wbk_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wbk_full_name)
    End If
    
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function WbkIsOpen( _
           Optional ByVal io_name As String = vbNullString, _
           Optional ByVal io_full_name As String) As Boolean
' ----------------------------------------------------------------------------
' When the full name is provided the check spans all Excel instances else only
' the current one.
' ----------------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
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
        io_name = Application.Workbooks(io_name).name
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

