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
Public Const README_URL                     As String = "https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md"
Public Const README_SYNC_CHAPTER            As String = "?#using-the-synchronization-service"
Public Const README_SYNC_CHAPTER_NAMES      As String = "?#names-synchronization"
Public Const README_FILES_AND_FOLDERS       As String = "?#files-and-folders"

Public Enum enKindOfComp            ' The kind of VBComponent in the sense of CompMan
    enUnknown = 0
    enCommCompHosted = 1
    enCommCompUsed = 2              ' The Component is a used raw, i.e. the raw is hosted by another Workbook
    enInternal = 3                  ' Neither a hosted nor a used Raw Common Component
End Enum

Public Enum siCounter
    sic_comps
    sic_comps_changed
    sic_used_comm_vbc_outdated
End Enum

Public lMaxCompLength   As Long
    
Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

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

Public Sub UpdateOutdatedCommonComponents(ByRef u_wbk_serviced As Workbook, _
                                 Optional ByVal u_hosted As String = vbNullString)
' ------------------------------------------------------------------------------
' Presents the serviced Workbook's outdated components in a modeless dialog with
' two buttons for each component. One button executes Application.Run mRenew.Run
' for a component to update it, the other executes Application.Run
' mService.ExpFilesDiffDisplay to display the code changes.
' Note: u_unused is for backwards compatibility only.
'
' Precondition: The service has been checked by the client to be able to run.
' ------------------------------------------------------------------------------
    Const PROC = "UpdateOutdatedCommonComponents"
    
    On Error GoTo eh
    EstablishExecTraceFile u_wbk_serviced
    
    mBasic.BoP ErrSrc(PROC)
    mService.Initiate mCompManClient.SRVC_UPDATE_OUTDATED, u_wbk_serviced
    mCommComps.Hskpng u_hosted
    mCompManDat.Hskpng u_hosted
    Set mCommComps.Qoutdated = Nothing
    mCommComps.OutdatedUpdate ' Dialog to update/renew one by one
    
xt: mBasic.EoP ErrSrc(PROC)
    mService.Terminate
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
    EstablishExecTraceFile e_wbk_serviced
        
    mBasic.BoP ErrSrc(PROC)
    mService.Initiate mCompManClient.SRVC_EXPORT_CHANGED, e_wbk_serviced
    If mService.Denied(mCompManClient.SRVC_EXPORT_CHANGED) Then GoTo xt
    
    mCommComps.Hskpng e_hosted
    mCompManDat.Hskpng e_hosted
    
    mService.ExportChangedComponents e_hosted
    ExportChangedComponents = True
    ExportChangedComponents = Application.StatusBar
    
xt: mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Application.EnableEvents = True
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function RunTest(ByVal rt_service As String, _
                        ByRef rt_serviced_wbk As Workbook) As Variant
' ----------------------------------------------------------------------------
' Ensures the requested service (rt_service) is able to run or returns the
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
    Select Case rt_service
        Case mCompManClient.SRVC_UPDATE_OUTDATED, mCompManClient.SRVC_EXPORT_CHANGED
            Select Case True
                Case Not wsConfig.FolderCompManRootIsValid:                              RunTest = AppErr(1) ' Configuration for the service is invalid
                Case Not rt_serviced_wbk.FullName Like wsConfig.FolderCompManRoot & "*": RunTest = AppErr(2) ' Denied because outside configured 'Dev and Test' folder
                Case rt_service = mCompManClient.SRVC_UPDATE_OUTDATED
                    If mMe.IsDevInstnc And ((mAddin.IsOpen And mAddin.Paused) Or Not mAddin.IsOpen) Then
                        RunTest = AppErr(3) ' Denied because serviced is the DevInstance but the Addin is paused or not open
                    End If
            End Select
        
        Case mCompManClient.SRVC_SYNCHRONIZE
            If wsConfig.FolderSyncTarget = vbNullString Or wsConfig.FolderSyncArchive = vbNullString Then
                RunTest = AppErr(1) ' Not configured
            ElseIf Not rt_serviced_wbk.FullName Like wsConfig.FolderSyncTarget & "*" Then
                RunTest = AppErr(2) ' Denied because not opened from within the configured 'Sync-Target' folder
            ElseIf Not mMe.IsDevInstnc Then
                RunTest = AppErr(3)
            End If
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    EstablishExecTraceFile sync_wbk_opened
    
    mBasic.BoP ErrSrc(PROC)
    mService.Initiate mCompManClient.SRVC_SYNCHRONIZE, sync_wbk_opened
    
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

Private Function WbkIsOpen(Optional ByVal io_name As String = vbNullString, _
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

Private Sub CheckForUnusedPublicItems()
' ----------------------------------------------------------------
' Attention! The service requires the VBPUnusedPublic.xlsb
'            Workbook open. When not open the service terminates
'            without notice.
' ----------------------------------------------------------------
    Const COMPS_EXCLUDED = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"
    Const LINES_EXCLUDED = "Select Case ErrMsg(ErrSrc(PROC))" & vbLf & _
                           "Case vbResume:*Stop:*Resume" & vbLf & _
                           "Case Else:*GoTo xt"
    On Error Resume Next
    '~~ Providing all (optional) arguments saves the Workbook selection dialog and the VBComponents selection dialog
    Application.Run "VBPUnusedPublic.xlsb!mUnused.Unused", ThisWorkbook, COMPS_EXCLUDED, LINES_EXCLUDED

End Sub
