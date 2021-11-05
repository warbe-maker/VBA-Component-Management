Attribute VB_Name = "mCompMan"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard Module mCompMan
'          Services for the management of VBComponents in Workbooks provided:
'          - stored within the 'ServicedRootFolder'
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
    enHostedRaw = 1
    enRawClone = 2             ' The Component is a used raw, i.e. the raw is hosted by another Workbook
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
    sic_clone_comps
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
    sic_raw_changed
    sic_clones_comps_updated
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

Private lMaxLenComp     As Long
Private lMaxLenTypeItem As Long

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

Public Sub CompareCloneWithRaw(ByVal cmp_comp_name As String)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "CompareCloneWithRaw"
    
    On Error GoTo eh
    Dim sExpFileRaw As String
    Dim wb          As Workbook
    Dim Comp        As New clsComp
    
    Set wb = ActiveWorkbook
    With Comp
        Set .Wrkbk = wb
        .CompName = cmp_comp_name
        Set .VBComp = wb.VBProject.VBComponents(.CompName)
        sExpFileRaw = mRawsHosted.ExpFileFullName(cmp_comp_name)
    
        mFile.Compare fc_file_left:=.ExpFileFullName _
                    , fc_file_right:=sExpFileRaw _
                    , fc_left_title:="The cloned raw's current code in Workbook/VBProject " & Comp.WrkbkBaseName & " (" & .ExpFileFullName & ")" _
                    , fc_right_title:="The remote raw's current code in Workbook/VBProject " & mBasic.BaseName(mRawsHosted.HostFullName(.CompName)) & " (" & sExpFileRaw & ")"

    End With
    Set Comp = Nothing

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
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
    
    mFile.Compare fc_file_left:=fl_1 _
                , fc_left_title:=fl_1 & " ( b e f o r e  the changes)" _
                , fc_file_right:=fl_2 _
                , fc_right_title:=fl_2 & " ( a f t e r  the changes)"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub DisplayCodeChange(ByVal cmp_comp_name As String)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "DisplayCodeChange"
    
    On Error GoTo eh
    Dim sTempExpFileFullName    As String
    Dim Comp                    As New clsComp
    Dim fso                     As New FileSystemObject
    Dim sTmpFolder              As String
    Dim flExpTemp               As File
    
    With Comp
        Set .Wrkbk = mService.Serviced
        .CompName = cmp_comp_name
        Set .VBComp = .Wrkbk.VBProject.VBComponents(.CompName)
    End With
    
    With fso
        sTmpFolder = mCompMan.ExpFileFolderPath(Comp.Wrkbk) & "\Temp"
        If Not .FolderExists(sTmpFolder) Then .CreateFolder sTmpFolder
        sTempExpFileFullName = sTmpFolder & "\" & Comp.CompName & Comp.ExpFileExt
        Comp.VBComp.Export sTempExpFileFullName
        Set flExpTemp = .GetFile(sTempExpFileFullName)
    End With

    With Comp
        mFile.Compare fc_file_left:=sTempExpFileFullName _
                    , fc_file_right:=.ExpFileFullName _
                    , fc_left_title:="The component's current code in Workbook/VBProject " & Comp.WrkbkBaseName & " ('" & sTempExpFileFullName & "')" _
                    , fc_right_title:="The component's currently exported code in '" & .ExpFileFullName & "'"

    End With
    
xt: If fso.FolderExists(sTmpFolder) Then fso.DeleteFolder (sTmpFolder)
    Set Comp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub DsplyProgress( _
          Optional ByVal p_result As String = vbNullString, _
          Optional ByVal p_total As Long = 0, _
          Optional ByVal p_done As Long = 0)

    Dim sService    As String
    Dim sDots       As String
    Dim sMsg        As String
    
    If p_total - p_done >= 0 Then sDots = VBA.String$(p_total - p_done, ".")
    On Error Resume Next
    sService = Log.Service
    sMsg = sService & sDots & p_result
    If Len(sMsg) > 250 Then sMsg = Left(sMsg, 246) & "...."
    If Right(sMsg, 1) = " " Then sMsg = Left(sMsg, Len(sMsg) - 1) & " ."
    Application.StatusBar = sMsg

End Sub

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCompMan" & "." & es_proc
End Function

Public Function ExpFileFolderPath(ByVal v As Variant) As String
' --------------------------------------------------------------------------
' ! This is the only source for the location of a Workbook's Export-Files !
' All Export-Files are placed in a '\source' sub-folder within the Workbook
' folder. This maitains a clear Workbook folder which matters (not only)
' when the Workbook-Folder is identical with a Github repo clone folder.
' (v) may be provided as a Workbook object or a Workbook's FullName string
' ----------------------------------------------------------------------------
    Const PROC                  As String = "ExpFileFolderPath"
    Const EXP_FILES_SUB_FOLDER  As String = "\source"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim wb  As Workbook
    Dim s   As String
    
    With fso
        Select Case TypeName(v)
            Case "Workbook"
                Set wb = v
                ExpFileFolderPath = wb.Path & EXP_FILES_SUB_FOLDER
            Case "String"
                s = v
                If Not .FileExists(s) _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "'" & s & "' is not the FullName of an existing Workbook!"
                ExpFileFolderPath = .GetParentFolderName(s) & EXP_FILES_SUB_FOLDER
            Case Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The required information about the concerned Workbook is neither provided as a Workbook object nor as a string identifying an existing Workbooks FullName"
        End Select
        If Not .FolderExists(ExpFileFolderPath) Then .CreateFolder ExpFileFolderPath
    End With
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Sub ExportAll(Optional ByRef ea_wb As Workbook = Nothing)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    
    If ea_wb Is Nothing _
    Then Set mService.Serviced = ActiveWorkbook _
    Else Set mService.Serviced = ea_wb
    mExport.All
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub ExportChangedComponents( _
                             ByRef ec_wb As Workbook, _
                    Optional ByVal ec_hosted As String = vbNullString)
' --------------------------------------------------------------------
' Exclusively performed/trigered by the Before_Save event:
' - Any code change (detected by the comparison of a temporary export
'   file with the current export file) is backed-up/exported
' - Outdated Export Files (components no longer existing) are removed
' - Clone code modifications update the raw code when confirmed by the
'   user
' --------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Set mService.Serviced = ec_wb
    
    mService.ExportChangedComponents ec_hosted
    
xt: mErH.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Install()
    mService.Install ActiveWorkbook
End Sub

Public Sub ManageHostHostedProperty(ByVal mh_hosted As String)
' ---------------------------------------------------------
' - Registers a Workbook as 'Raw-Host' when it has at least
'   one of the Workbook's (mh_wb) VBComponents indicated
'   hosted.
' - Registers each hosted 'Raw-Component' as such
' ---------------------------------------------------------
    Const PROC = "ManageHostHostedProperty"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim Comp           As clsComp
    Dim sHosted         As String
    Dim sHostBaseName   As String
    
    mErH.BoP ErrSrc(PROC)
    sHostBaseName = fso.GetBaseName(mService.Serviced.FullName)
    Set dctHostedRaws = New Dictionary
    HostedRaws = mh_hosted
                    
    If HostedRaws.Count <> 0 Then
        For Each v In HostedRaws
            '~~ Keep a record for each of the VBComponents hosted by this Workbook
            If Not mRawsHosted.Exists(raw_comp_name:=v) _
            Or mRawsHosted.HostFullName(comp_name:=v) <> mService.Serviced.FullName Then
                mRawsHosted.HostFullName(comp_name:=v) = mService.Serviced.FullName
                Log.Entry = "Raw-Component '" & v & "' hosted in this Workbook registered"
            End If
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = mService.Serviced
                .CompName = v
                If mRawsHosted.ExpFileFullName(v) = vbNullString _
                Or mRawsHosted.ExpFileFullName(v) <> .ExpFileFullName Then
                    '~~ The component is yet not registered or registered under an outdated location
                    mRawsHosted.ExpFileFullName(v) = .ExpFileFullName
                    '~~ Just in case its a new VBComponent - by the way
                    If Not fso.FileExists(.ExpFileFullName) Then .VBComp.Export .ExpFileFullName
                End If
            End With
            Set Comp = Nothing
        Next v
    Else
        '~~ Remove any raws still existing and pointing to this Workbook as host
        For Each v In mRawsHosted.Components
            If mRawsHosted.HostFullName(comp_name:=v) = mService.Serviced.FullName Then
                mRawsHosted.Remove comp_name:=v
                Log.Entry = "Component removed from '" & mRawsHosted.HostedRawsFile & "'"
            End If
        Next v
        If mRawHosts.Exists(fso.GetBaseName(mService.Serviced.FullName)) Then
            mRawHosts.Remove (fso.GetBaseName(mService.Serviced.FullName))
            Log.Entry = "Workbook no longer a host for at least one raw component removed from '" & mRawHosts.RawHostsFile & "'"
        End If
    End If

xt: Set fso = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub SynchTargetWbWithSourceWb( _
                               ByRef wb_target As Workbook, _
                               ByVal wb_source As String)
' -------------------------------------------------------------
' Synchronizes the code of the open/ed Workbook (clone_project)
' with the code of the source Workbook (raw_project).
' The service is performed provided:
' - the Workbook is open/ed in the configured "Serviced Root"
' - the CompMan Addin is not paused
' - the open/ed Workbook is not a restored version
' -------------------------------------------------------------
    Const PROC = "SynchTargetWbWithSourceWb"
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    Set mService.Serviced = wb_target
    mService.SyncVBProjects wb_target:=wb_target, wb_source_name:=wb_source
    
xt: mErH.EoP ErrSrc(PROC)
    Set Log = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub UpdateRawClones( _
                     ByRef uc_wb As Workbook, _
            Optional ByVal uc_hosted As String = vbNullString)
' ------------------------------------------------------------
' Updates a clone component with the Export-File of the remote
' raw component provided the raw's code has changed.
' ------------------------------------------------------------
    Const PROC = "UpdateRawClones"
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    
    Set Log = New clsLog
    Set mService.Serviced = uc_wb
    If mService.Denied(PROC) Then GoTo xt
    
    mCompMan.ManageHostHostedProperty uc_hosted
    Set Stats = New clsStats
    mUpdate.RawClones mService.Serviced
    
xt: Set Log = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function WbkGetOpen(ByVal go_wb_full_name As String) As Workbook
' ---------------------------------------------------------------------
' Returns an opened Workbook object named (go_wb_full_name) or Nothing
' when a file named (go_wb_full_name) not exists.
' ---------------------------------------------------------------------
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
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Function WbkIsOpen( _
           Optional ByVal io_name As String = vbNullString, _
           Optional ByVal io_full_name As String) As Boolean
' ------------------------------------------------------------
' When the full name is provided the check spans all Excel
' instances else only the current one.
' ------------------------------------------------------------
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

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Function WinMergeIsInstalled() As Boolean
    WinMergeIsInstalled = AppIsInstalled("WinMerge")
End Function

