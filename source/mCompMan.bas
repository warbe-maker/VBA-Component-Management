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
' Usage:   This Workbbok's services are available as 'CompMan-AddIn' when
'          - at least once the Renew service had been performed
'          - either a Workbbook referring to the Addin is opened
'          - or the Addin-Workbook-Development-Instance is opened and Renew
'            is performed again
'
'          For further detailed information see:
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
'                         - mFile
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
Public Const MAX_LEN_TYPE           As Long = 17
Public Const SRVC_UPDATE_OUTDATED   As String = "Update Outdated"
Public Const SRVC_SYNC_WORKBOOKS    As String = "Sync Target with Source Workbook"
Public Const SRVC_EXPORT_CHANGED    As String = "Export Changed"

Public Enum enKindOfComp       ' The kind of VBComponent in the sense of CompMan
    enUnknown = 0
    enCommCompHosted = 1
    enCommCompUsed = 2             ' The Component is a used raw, i.e. the raw is hosted by another Workbook
    enInternal = 3             ' Neither a hosted nor a used Raw Common Component
End Enum

' Distinguish the code of which Workbook is allowed to be updated
Public Enum vbcmType
    vbext_ct_StdModule = 1          ' .bas
    vbext_ct_ClassModule = 2        ' .cls
    vbext_ct_MSForm = 3             ' .frm
    vbext_ct_ActiveXDesigner = 11   ' ??
    vbext_ct_Document = 100         ' .cls
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
    sic_raw_comm_comp_changed
    sic_used_comm_comp_Outdated
    sic_used_comm_comp_updated
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

    
Private Property Get HostedRaws() As Variant:           Set HostedRaws = dctHostedRaws:                 End Property

Private Property Let HostedRaws(ByVal hr As Variant)
' ---------------------------------------------------
' Saves the names of the hosted raw components (hr)
' to the Dictionary (dctHostedRaws).
' ---------------------------------------------------
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

Public Sub DisplayChanges( _
           Optional ByVal fl_1 As String = vbNullString, _
           Optional ByVal fl_2 As String = vbNullString)
' --------------------------------------------------------
'
' --------------------------------------------------------
    Const PROC = "DisplayChanges"
    
    On Error GoTo eh
    Dim fl  As File
    Dim fso As New FileSystemObject
    
    If fl_1 = vbNullString Then
        mFile.Picked p_title:="Select the file regarded 'before the change' (displayed at the left)!" _
                   , p_file:=fl
        If Not fl Is Nothing Then fl_1 = fl.Path
    End If
    
    If fl_2 = vbNullString Then
        mFile.Picked p_title:="Select the file regarded 'the changed one' (displayed at the right)!" _
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

Private Sub EstablishTraceLogFile(ByVal dt_wb As Workbook, _
                         Optional ByVal dt_append As Boolean = False)
' --------------------------------------------------------------------------
' Establishes a trace log file in the serviced Workbook's parent folder.
' --------------------------------------------------------------------------
    Dim sFile As String
    sFile = Replace(dt_wb.FullName, dt_wb.Name, "CompMan.Service.trc")

    '~~ Even when dt_append = False: When the filke had been createde today dt_append will be set to True
    With New FileSystemObject
        If .FileExists(sFile) Then
            If Format(.GetFile(sFile).DateCreated, "YYYY-MM-DD") = Format(Now(), "YYYY-MM-DD") Then
                dt_append = True
            End If
        End If
    End With
    mTrc.LogFile(tl_append:=dt_append) = sFile
    mTrc.LogTitle = Log.Service

End Sub

Public Sub ExportAll(Optional ByRef ea_wb As Workbook = Nothing)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    
    If Log Is Nothing Then Set Log = New clsLog
    Log.Service = "Export All"
    EstablishTraceLogFile ea_wb
    
    mBasic.BoP ErrSrc(PROC)
    If ea_wb Is Nothing _
    Then Set mService.Serviced = ActiveWorkbook _
    Else Set mService.Serviced = ea_wb
    mExport.All
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ExportChangedComponents( _
                ByRef ec_wb As Workbook, _
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

' ----------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    Set mService.Serviced = ec_wb
    Set Log = New clsLog
    Log.Service = SRVC_EXPORT_CHANGED
    
    '~~ Determine any reason the service basically cannot be provided
    If Not mMe.FolderServicedIsValid Then
        '~~ The serviced root folder is invalid (not configured or not existing)
        ExportChangedComponents = AppErr(1)
    ElseIf Not ec_wb.FullName Like mConfig.FolderServiced & "*" Then
        '~~ The serviced Workbook is located outside the serviced folder
        ExportChangedComponents = AppErr(2)
    ElseIf mMe.IsAddinInstnc And mMe.CompManAddinIsPaused Then
        '~~ When the service is about to be provided by the Addin but the Addin is currently paused
        '~~ another try with the serviced provided by the open Development instance may do the job.
        ExportChangedComponents = AppErr(4)
    
    '~~ The very basic requirements are met
    Else
        EstablishTraceLogFile ec_wb
        mBasic.BoP ErrSrc(PROC)
        
        If mService.Denied Then GoTo xt
        
        mService.ExportChangedComponents ec_hosted
        ExportChangedComponents = True
        ExportChangedComponents = Application.StatusBar
        
        mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    End If

xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Install()
    mService.Install ActiveWorkbook
End Sub

Public Sub MaintainPropertiesOfHostedRawCommonComponents(ByVal mh_hosted As String)
' ----------------------------------------------------------------------------
' Manages all aspects of Raw/Hosted Common Components which includes the
' copies of the Export File in the Common Components Folder.
' - Registers a Workbook as 'Raw-Host' when it claims hosting at least one
'   Common Component (mh_wb)
' - Registers for each hosted 'Raw Common Component':
'   - in the local ComCompsHosted.dat the properties:
'     - Component Name
'     - Revision Number
'   - in the ComComps-RawsSaved.dat in the Common Components folder:
'     - Component Name
'     - Export File Full Name
'     - Host Base Name
'     - Host Full Name
'     - Host Name
'     - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "MaintainPropertiesOfHostedRawCommonComponents"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim Comp            As clsComp
    Dim sHostBaseName   As String
    
    mBasic.BoP ErrSrc(PROC)
    sHostBaseName = fso.GetBaseName(mService.Serviced.FullName)
    Set dctHostedRaws = New Dictionary
    HostedRaws = mh_hosted
                    
    If HostedRaws.Count <> 0 Then
        For Each v In HostedRaws
            If Not mComCompsRawsHosted.Exists(v) Then
                '~~ Initially register the component as a 'Hosted Raw Common Component'
                mComCompsRawsHosted.RawRevisionNumberIncrease v     ' this will initially set it
            End If
            If Not mComCompsRawsSaved.Exists(v) Then
                '~~ Initially registers the Common Component in the ComComps-RawsSaved.dat file
                '~~ Note: The Raw's Revision Number is updated whenever the raw is exported because it had been modified
                mComCompsRawsSaved.RawHostWbFullName(v) = mService.Serviced.FullName
                mComCompsRawsSaved.RawHostWbName(v) = mService.Serviced.Name
                mComCompsRawsSaved.RawHostWbBaseName(v) = fso.GetBaseName(mService.Serviced.FullName)
                Log.Entry = "Raw-Component '" & v & "' hosted in this Workbook registered"
            ElseIf StrComp(mComCompsRawsSaved.RawHostWbFullName(v), mService.Serviced.FullName, vbTextCompare) <> 0 _
                Or StrComp(mComCompsRawsSaved.RawHostWbName(v), mService.Serviced.Name, vbTextCompare) <> 0 Then
                '~~ Update the properties when they had changed - which may happen when the Raw Common Component's
                '~~ host has changed
                '~~ Note: The RevisionNumber is updated whenever the modified raw is exported
                mComCompsRawsSaved.RawHostWbFullName(v) = mService.Serviced.FullName
                mComCompsRawsSaved.RawHostWbName(v) = mService.Serviced.Name
                mComCompsRawsSaved.RawHostWbBaseName(v) = fso.GetBaseName(mService.Serviced.FullName)
                Log.Entry = "Raw Common Component '" & v & "' hosted changed properties updated"
            End If
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = mService.Serviced
                .CompName = v
                If Not fso.FileExists(.ExpFileFullName) Then
                    '~~ Initially export the new Common Component claimed being hosted in this Workbook
                    .VBComp.Export .ExpFileFullName
                End If
                mComCompsRawsHosted.RawExpFileFullName(v) = .ExpFileFullName ' in any case update the Export File name
                If mService.FilesDiffer(fd_exp_file_1:=.ExpFile _
                                          , fd_exp_file_2:=mComCompsRawsSaved.SavedExpFile(v)) Then
                    '~~ Attention! This is a cruical issue which shold never be the case. However, when different
                    '~~ computers/users are involved in the development process ...
                    '~~ Instead of simply updating the saved raw Export File better have carefully checked the case
                    If mComCompsRawsSaved.RawSavedRevisionNumber(v) = mComCompsRawsHosted.RawRevisionNumber(v) Then
                        If SavedRawInconsitencyWarning _
                           (sri_raw_exp_file_full_name:=.ExpFile.Path _
                          , sri_saved_exp_file_full_name:=mComCompsRawsSaved.SavedExpFile(v).Path _
                          , sri_diff_message:="While the Revision Number of the 'Hosted Raw'  " & mBasic.Spaced(v) & "  is identical with the " & _
                                              "'Saved Raw' their Export Files are different. Compared were:" & vbLf & _
                                              "Hosted Raw Export File = " & .ExpFile.Path & vbLf & _
                                              "Saved Raw Export File  = " & mComCompsRawsSaved.SavedExpFile(v).Path & vbLf & _
                                              "whereby any empty code lines and case differences had been ignored. " & _
                                              "The difference thus really matters!" _
                           ) Then
                            .CopyExportFileToCommonComponentsFolder
                        End If
                    ElseIf mComCompsRawsSaved.RawSavedRevisionNumber(v) <> mComCompsRawsHosted.RawRevisionNumber(v) Then
                        If SavedRawInconsitencyWarning _
                           (sri_raw_exp_file_full_name:=.ExpFile.Path _
                          , sri_saved_exp_file_full_name:=mComCompsRawsSaved.SavedExpFile(v).Path _
                          , sri_diff_message:="The 'Revision Number' of the 'Hosted Raw Common Component's Export File' and the " & _
                                              "the 'Saved Raw's Export File' differ:" & vbLf & _
                                              "Hosted Raw = " & mComCompsRawsHosted.RawRevisionNumber(v) & vbLf & _
                                              "Saved Raw  = " & mComCompsRawsSaved.RawSavedRevisionNumber(v) & vbLf & _
                                              "and also the Export Files differ. Compared were:" & vbLf & _
                                              "Hosted Raw = " & .ExpFile.Path & vbLf & _
                                              "Saved Raw  = " & mComCompsRawsSaved.SavedExpFile(v).Path & vbLf & _
                                              "whereby any empty code lines and case differences had been ignored. " & _
                                              "The difference thus really matters! Updating is not at all " & _
                                              "recommendable before the issue had been clarified." _
                           ) Then
                            .CopyExportFileToCommonComponentsFolder
                        End If
                    End If
                End If
            End With
            Set Comp = Nothing
        Next v
    Else
        '~~ When this Workbook not or no longer hosts any Common Component Raws the entries
        '~~ the ComCompsHosted.dat is deleted
        mFile.Delete mComCompsRawsHosted.ComCompsHostedFileFullName
        '~~ The destiny of the corresponding data in the ComComps-Saved.dat is un-clear
        '~~ The component may be now hosted in another Workbook (likely) or the life of the
        '~~ Common Component has ended. The entry will be removed when it still points to this
        '~~ Workbook. When it points to another one it appears to have been moved alrerady.
        For Each v In mComCompsRawsSaved.Components
            If StrComp(mComCompsRawsSaved.RawHostWbFullName(comp_name:=v), mService.Serviced.FullName, vbTextCompare) = 0 Then
                mComCompsRawsSaved.Remove comp_name:=v
                Log.Entry = "Component no longer hosted in '" & mService.Serviced.FullName & "' removed from '" & mComCompsRawsSaved.ComCompsSavedFileFullName & "'"
            End If
        Next v
    End If

xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function RunTest(ByVal rt_service As String, _
                        ByVal rt_serviced_wb As Workbook) As Variant
' --------------------------------------------------------------------------
' Ensures the requested service is able to run or returns the reason why not.
' The function returns:
' - AppErr(1) when CompMan is not configured properly
' - AppErr(2) when the service Workbook is not in the configures Folder
' - AppErr(3) when the requested service (rt_service) was
'             "UpdateOutdatedUsedCommonComponents" but the servicing and the
'             serviced Workbook are both CompMan.xlsb - and the Workbook
'             cannot update itself. This error only happens when the Addin
'             is not available.
' - AppErr(4) when the servicing Workbook is the Addin which is available
'             but paused. It would requires CompMan.xlsb to run the
'             service. When CompMan.xlsb is open, it will provide the
'             service.
' --------------------------------------------------------------------------
    Const PROC = "RunTest"
    
    On Error GoTo eh
    If Not mMe.FolderServicedIsValid Then
        RunTest = AppErr(1) ' The serviced root folder is invalid (not configured or not existing)
    ElseIf Not rt_serviced_wb.FullName Like mConfig.FolderServiced & "*" Then
        RunTest = AppErr(2) ' The serviced Workbook is located outside the serviced folder
    ElseIf mMe.IsAddinInstnc And mMe.CompManAddinIsPaused Then
        RunTest = AppErr(4) ' The service is about to be provided by the Addin but the Addin is currently paused
    ElseIf rt_service = "UpdateOutdatedCommonComponents" And mMe.IsDevInstnc And rt_serviced_wb.Name = mMe.DevInstncName Then
        RunTest = AppErr(3) ' The servicing and the serviced Workbook are both the 'CompMan Development Instance'
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function SavedRawInconsitencyWarning(ByVal sri_raw_exp_file_full_name, _
                                             ByVal sri_saved_exp_file_full_name, _
                                             ByVal sri_diff_message) As Boolean
' ----------------------------------------------------------------------------
' Displays an information about a modification of a Used Common Component.
' The disaplay offers the option to display the code difference.
' Returns TRUE only when the reply is "go ahead, update anyway"
' ----------------------------------------------------------------------------
    Const PROC = "SavedRawInconsitencyWarning"
    
    On Error GoTo eh
    Dim Msg         As mMsg.TypeMsg
    Dim cllBttns    As Collection
    Dim BttnDsply   As String
    Dim BttnSkip    As String
    Dim BttnAnyway  As String
    
    BttnDsply = "Display code difference" & vbLf & "between hosted and saved" & vbLf & "Export Files"
    BttnSkip = "Do not update!" & vbLf & "further investigation" & vbLf & "is required"
    BttnAnyway = "I know the reason!" & vbLf & "go ahead updating" & vbLf & "(not recommended!)"
    
    Set cllBttns = mMsg.Buttons(BttnDsply, vbLf, BttnAnyway, vbLf, BttnSkip)
    With Msg.Section(1)
        With .Label
            .Text = "Attention!"
            .FontColor = rgbRed
        End With
        With .Text
            .Text = sri_diff_message
            .FontColor = rgbRed
        End With
    End With
    With Msg.Section(2)
        .Label.Text = "Background:"
        .Text.Text = "When a Raw Common Component is modified within its hosting Workbook it is not only exported. " & _
                     "Its 'Revision Number' is increased and the 'Export File' is copied into the 'Common Components' " & _
                     "folder while the 'Revision Number' is updated in the 'ComComps-RawsSaved.dat' file in the " & _
                     "'Common Components' folder. Thus, the Raw's Export File and the copy of it as the 'Revision Number' " & _
                     "are always identical. In case not, something is seriously corrupted."
    End With
        
    Do
        If Not mMsg.IsValidMsgButtonsArg(cllBttns) Then Stop
        Select Case mMsg.Dsply(dsply_title:="Serious inconsistency warning!" _
                             , dsply_msg:=Msg _
                             , dsply_buttons:=cllBttns _
                              )
            Case BttnDsply
                mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=sri_raw_exp_file_full_name _
                                               , fd_exp_file_left_title:="Raw Common Component's Export File: (" & sri_raw_exp_file_full_name & ")" _
                                               , fd_exp_file_right_full_name:=sri_saved_exp_file_full_name _
                                               , fd_exp_file_right_title:="Saved Raw's Export File (" & sri_saved_exp_file_full_name & ")"
            Case BttnSkip:      SavedRawInconsitencyWarning = False:    Exit Do
            Case BttnAnyway:    SavedRawInconsitencyWarning = True:     Exit Do
        End Select
    Loop

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SynchTargetWbWithSourceWb( _
                               ByRef wb_target As Workbook, _
                               ByVal wb_source As String)
' ----------------------------------------------------------------------------
' Synchronizes the code of the open/ed Workbook (clone_project) with the code
' of the source Workbook (raw_project). The service is performed provided:
' - the Workbook is open/ed in the configured "Serviced Root"
' - the CompMan Addin is not paused
' - the open/ed Workbook is not a restored version
' ----------------------------------------------------------------------------
    Const PROC = "SynchTargetWbWithSourceWb"
    
    On Error GoTo eh
    
    Set mService.Serviced = wb_target
    If Log Is Nothing Then Set Log = New clsLog
    Log.Service = SRVC_SYNC_WORKBOOKS
    EstablishTraceLogFile wb_target
    
    mBasic.BoP ErrSrc(PROC)
    mService.SyncVBProjects wb_target:=wb_target, wb_source_name:=wb_source
    
xt: mBasic.EoP ErrSrc(PROC)
    Set Log = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function UpdateOutdatedCommonComponents(ByRef uo_wb As Workbook, _
                                      Optional ByVal uo_hosted As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
' Updates all outdated 'Used Common Components'.
'
' The function is terminated (returns FALSE) without further notice when:
' a) the serviced root folder is invalid (not configured or not existing)
' b) when the servicing and the serviced Workbook are both the 'CompMan
'    Development Instance'
'    Note: While the 'Development instance is able to export its modified
'          components it cannot update its own outdated 'Used Common
'          Components'. This is only possible by the 'Addin Instance'
'          which must be open and not 'paused'
' c) the serviced Workbook is located outside the serviced folder
'
' When the function returns vbNullString it is terminated "silent" which is
' the case when the seftviced Workbook does not reside within the 'Serviced
' Folder'. Any return string <> vbNullsString is displayed in the
' Application.StatusBar which may be caused by:
' a) the Workbook is one restored by Excel
' b) the serviced Workbook does not reside in a folder exclusivelyx (i.e. the
'    Workbook does not live in its own dedicated folder
' c) WinMerge is not installed
' ------------------------------------------------------------------------------
    Const PROC = "UpdateOutdatedCommonComponents"
    
    On Error GoTo eh
    
    Set mService.Serviced = uo_wb
    Set Log = New clsLog
    Log.Service(new_log:=True) = SRVC_UPDATE_OUTDATED
    
    '~~ Prevent any service is performed when not possible, applicable or any other reason
    If Not mMe.FolderServicedIsValid Then
        '~~ The serviced root folder is invalid (not configured or not existing)
        UpdateOutdatedCommonComponents = AppErr(1)
    ElseIf Not uo_wb.FullName Like mConfig.FolderServiced & "*" Then
        '~~ The serviced Workbook is located outside the serviced folder
        UpdateOutdatedCommonComponents = AppErr(2)
    ElseIf mMe.IsDevInstnc And uo_wb.Name = mMe.DevInstncName Then
        '~~ The servicing and the serviced Workbook are both the 'CompMan Development Instance'
        '~~ This is the case when either no CompMan-Addin-Instance is available or it is currently paused
        UpdateOutdatedCommonComponents = AppErr(3)
    ElseIf mMe.IsAddinInstnc And mMe.CompManAddinIsPaused Then
        '~~ When the service is about to be provided by the Addin but the Addin is currently paused
        '~~ another try with the serviced provided by the open Development instance may do the job.
        UpdateOutdatedCommonComponents = AppErr(4)
    
    Else
        EstablishTraceLogFile uo_wb
        mBasic.BoP ErrSrc(PROC)
        
        If mService.Denied Then GoTo xt
        
        mCompMan.MaintainPropertiesOfHostedRawCommonComponents uo_hosted
        Set Stats = New clsStats
        mUpdate.Outdated
        
        '~~ !!!! Saving the Workbook is likely to cause Excel to crash                             !!!!!
        '~~ !!!! Not saving the Workbook programmatically is stabilizing the process significantly !!!!
'        mService.SaveWbk uo_wb

        UpdateOutdatedCommonComponents = True
        
        mBasic.EoP ErrSrc(PROC)
    End If
    
xt: Set Log = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function WbkGetOpen(ByVal go_wb_full_name As String) As Workbook
' ----------------------------------------------------------------------------
' Returns an opened Workbook object named (go_wb_full_name) or Nothing when a
' file named (go_wb_full_name) not exists.
' ----------------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If fso.FileExists(go_wb_full_name) Then
        If mCompMan.WbkIsOpen(io_name:=go_wb_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wb_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wb_full_name)
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

