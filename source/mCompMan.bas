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
Public Const MAX_LEN_TYPE As Long = 17

Public Enum enKindOfComp       ' The kind of VBComponent in the sense of CompMan
    enUnknown = 0
    enCommCompHosted = 1
    enCommCompUsed = 2             ' The Component is a used raw, i.e. the raw is hosted by another Workbook
    enInternal = 3             ' Neither a hosted nor a used Raw Common Component
End Enum

Public Enum enUpdateReply
    enUpdateOriginWithUsed
    enUpdateUsedWithOrigin
    enUpdateNone
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

Public asNoSynch()      As String
Public lMaxCompLength   As Long
Public dctHostedRaws    As Dictionary
Public Stats            As clsStats
Public Log              As clsLog
Public TraceLog         As clsLog
Private lMaxLenComp     As Long
Private lMaxLenTypeItem As Long

    
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

'Public Sub CompareUsedWithRaw(ByVal cmp_comp_name As String)
'' -----------------------------------------------------------
''
'' -----------------------------------------------------------
'    Const PROC = "CompareUsedWithRaw"
'
'    On Error GoTo eh
'    Dim sExpFileRaw As String
'    Dim wb          As Workbook
'    Dim Comp        As New clsComp
'
'    Set wb = ActiveWorkbook
'    With Comp
'        Set .Wrkbk = wb
'        .CompName = cmp_comp_name
'        Set .VBComp = wb.VBProject.VBComponents(.CompName)
'        If .KindOfComp = enCommCompUsed Then
'            sExpFileRaw = .Raw.ExpFileFullName
'            mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=.ExpFileFullName _
'                                       , fd_exp_file_right_full_name:=sExpFileRaw _
'                                       , fd_exp_file_left_title:="Clone (used) Common Component code in Workbook/VBProject " & Comp.WrkbkBaseName & " (" & .ExpFileFullName & ")" _
'                                       , fd_exp_file_right_title:="Raw (hosted) Common Component code in Workbook/VBProject " & mBasic.BaseName(mComCompsRawsSaved.HostWbFullName(.CompName)) & " (" & sExpFileRaw & ")"
'
'        Else
'            mMsg.Box box_title:="Not a known 'Common Component'!" _
'                   , box_msg:="The provided component name '" & cmp_comp_name & "' is not registered/known as a 'Common Component'. " & _
'                              "To have this component been recognized by CompMan as a 'Common Component' one Workbook has to claim " & _
'                              "hosting it as the 'Raw Used Common Component'."
'        End If
'    End With
'    Set Comp = Nothing
'
'xt: Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Private Sub DetermineTraceLogFolder(ByVal dt_wb As Workbook)
' --------------------------------------------------------------------------
' Determines the location for the service execution trace log folder - which
' by the way suspends the display of the trace result.
' --------------------------------------------------------------------------
    If mMe.IsDevInstnc _
    Then mTrc.TraceLogFile = Replace(dt_wb.FullName, dt_wb.name, "CompMan.Trace.log") _
    Else mTrc.TraceLogFile = mConfig.FolderAddin & "\CompManAdmin\CompMan.Service.log"
End Sub

Public Sub DisplayChanges( _
           Optional ByVal fl_1 As String = vbNullString, _
           Optional ByVal fl_2 As String = vbNullString)
' --------------------------------------------------------
'
' --------------------------------------------------------
    Const PROC = "DisplayChanges"
    
    On Error GoTo eh
    Dim fl_left         As File
    Dim fl_right        As File
    
    If fl_1 = vbNullString _
    Then mFile.SelectFile sel_result:=fl_left, _
                          sel_title:="Select the file regarded 'before the change' (displayed at the left)!"
    
    If fl_2 = vbNullString _
    Then mFile.SelectFile sel_result:=fl_right, _
                          sel_title:="Select the file regarded 'the changed one' (displayed at the right)!"
    
    fl_1 = fl_left.Path
    fl_2 = fl_right.Path
    
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

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCompMan" & "." & es_proc
End Function


Public Sub ExportAll(Optional ByRef ea_wb As Workbook = Nothing)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    DetermineTraceLogFolder ea_wb
    
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
        mBasic.BoP ErrSrc(PROC)
        DetermineTraceLogFolder ec_wb
        Set Log = New clsLog
        Set mService.Serviced = ec_wb
        If mService.Denied(PROC) Then GoTo xt
        mCompMan.ManageRawCommonComponentsProperties ec_hosted
        
        mService.ExportChangedComponents ec_hosted
        ExportChangedComponents = True
        mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
        ExportChangedComponents = Application.StatusBar
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

Public Sub ManageRawCommonComponentsProperties(ByVal mh_hosted As String)
' ----------------------------------------------------------------------------
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
    Const PROC = "ManageRawCommonComponentsProperties"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim Comp            As clsComp
    Dim sHosted         As String
    Dim sHostBaseName   As String
    
    mBasic.BoP ErrSrc(PROC)
    sHostBaseName = fso.GetBaseName(mService.Serviced.FullName)
    Set dctHostedRaws = New Dictionary
    HostedRaws = mh_hosted
                    
    If HostedRaws.Count <> 0 Then
        For Each v In HostedRaws
            If Not mComCompsRawsHosted.Exists(raw_comp_name:=v) Then
                '~~ Register the component as a 'Hosted Common Component'
                mComCompsRawsHosted.RevisionNumberIncrease (v) ' this will initially set it
            End If
            If Not mComCompsRawsSaved.Exists(raw_comp_name:=v) Then
                '~~ Keep a record in the ComComps-RawsSaved.dat file for each of the VBComponents hosted by this Workbook
                '~~ Note: The RevisionNumber is updated whenever the modified raw is exported
                mComCompsRawsSaved.HostWbFullName(comp_name:=v) = mService.Serviced.FullName
                mComCompsRawsSaved.HostWbName(comp_name:=v) = mService.Serviced.name
                mComCompsRawsSaved.HostWbBaseName(comp_name:=v) = fso.GetBaseName(mService.Serviced.FullName)
                Log.Entry = "Raw-Component '" & v & "' hosted in this Workbook registered"
            ElseIf StrComp(mComCompsRawsSaved.HostWbFullName(comp_name:=v), mService.Serviced.FullName, vbTextCompare) <> 0 _
                Or StrComp(mComCompsRawsSaved.HostWbName(comp_name:=v), mService.Serviced.name, vbTextCompare) <> 0 Then
                '~~ Keep the hosted raw's properties in the ComComps-RawsSaved.dat up-to-date
                '~~ Note: The RevisionNumber is updated whenever the modified raw is exported
                mComCompsRawsSaved.HostWbFullName(comp_name:=v) = mService.Serviced.FullName
                mComCompsRawsSaved.HostWbName(comp_name:=v) = mService.Serviced.name
                mComCompsRawsSaved.HostWbBaseName(comp_name:=v) = fso.GetBaseName(mService.Serviced.FullName)
                Log.Entry = "Raw-Component '" & v & "' hosted in this Workbook registered"
            End If
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = mService.Serviced
                .CompName = v
                If Not fso.FileExists(.ExpFileFullName) Then
                    '~~ Just in case its a new VBComponent claimed by the host Workbook a Raw Common Component
                    .VBComp.Export .ExpFileFullName
                End If
                If mComCompsRawsSaved.ExpFileFullName(v) = vbNullString Then
                    '~~ The component is yet not registered or registered under an outdated location
                    mComCompsRawsSaved.ExpFileFullName(v) = .ExpFileFullName
                ElseIf mService.FilesDiffer(fd_exp_file_1:=.ExpFile _
                                          , fd_exp_file_2:=mComCompsRawsSaved.ExpFile(.CompName)) Then
                    '~~ Make sure the most up-to-date Export File has been copied to the Common Components Folder
                    .CopyExportFileToCommonComponentsFolder
                End If
            End With
            Set Comp = Nothing
        Next v
    Else
        '~~ When this Workbook not or no longer hosts any Common Component Raws the entries
        '~~ the ComCompsHosted.dat is deleted
        fso.DeleteFile mComCompsRawsHosted.ComCompsHostedFileFullName
        '~~ The destiny of the corresponding data in the ComComps-Saved.dat is un-clear
        '~~ The component may be now hosted in another Workbook (likely) or the life of the
        '~~ Common Component has ended. The entry will be removed when it still points to this
        '~~ Workbook. When it points to another one it appears to have been moved alrerady.
        For Each v In mComCompsRawsSaved.Components
            If StrComp(mComCompsRawsSaved.HostWbFullName(comp_name:=v), mService.Serviced.FullName, vbTextCompare) = 0 Then
                mComCompsRawsSaved.Remove comp_name:=v
                Log.Entry = "Component no longer hosted in '" & mService.Serviced.FullName & "' removed from '" & mComCompsRawsSaved.ComCompsSavedFileFullName & "'"
            End If
        Next v
    End If

xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
    mBasic.BoP ErrSrc(PROC)
    Set mService.Serviced = wb_target
    mService.SyncVBProjects wb_target:=wb_target, wb_source_name:=wb_source
    
xt: mBasic.EoP ErrSrc(PROC)
    Set Log = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function UpdateOutdatedCommonComponents( _
                ByRef uo_wb As Workbook, _
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
    Dim msg As TypeMsg
    
    If Not mMe.FolderServicedIsValid Then
        '~~ The serviced root folder is invalid (not configured or not existing)
        UpdateOutdatedCommonComponents = AppErr(1)
    ElseIf Not uo_wb.FullName Like mConfig.FolderServiced & "*" Then
        '~~ The serviced Workbook is located outside the serviced folder
        UpdateOutdatedCommonComponents = AppErr(2)
        
    ElseIf mMe.IsDevInstnc And uo_wb.name = mMe.DevInstncName Then
        '~~ The servicing and the serviced Workbook are both the 'CompMan Development Instance'
        '~~ This is the case when either no CompMan-Addin-Instance is available or it is currently paused
        UpdateOutdatedCommonComponents = AppErr(3)
    ElseIf mMe.IsAddinInstnc And mMe.CompManAddinIsPaused Then
        '~~ When the service is about to be provided by the Addin but the Addin is currently paused
        '~~ another try with the serviced provided by the open Development instance may do the job.
        UpdateOutdatedCommonComponents = AppErr(4)
    Else
        mBasic.BoP ErrSrc(PROC)
        DetermineTraceLogFolder uo_wb
        
        Set Log = New clsLog
        Set mService.Serviced = uo_wb
        If mService.Denied(PROC) Then GoTo xt
        
        mCompMan.ManageRawCommonComponentsProperties uo_hosted
        Set Stats = New clsStats
        mUpdate.Outdated mService.Serviced
        UpdateOutdatedCommonComponents = True
        
        Set Log = Nothing
        mBasic.EoP ErrSrc(PROC)
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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

Private Function WbkInServicedRoot() As Boolean
    WbkInServicedRoot = InStr(mService.Serviced.Path, mConfig.FolderServiced) <> 0
End Function

