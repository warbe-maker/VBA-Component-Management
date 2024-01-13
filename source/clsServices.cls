VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------------------------------
' Class Module clsServices: Common services used throughout this VB-Project.
' =========================

' Public services:
' - CurrentServiceStatusBar ...
' - Denied                  ...
' - DsplyStatus             ...
' - EstablishExecTraceFile  ...
' - ExpFilesDiffDisplay     ...
' - ExportChangedComponents ....
' - FilesDiffer             ...
' - FilesDifference         ...
' - Initiate                ...
' - IsRenamedByCompMan      ...
' - ServicedItemLogEntry    ...
' - LogServicingEntry       ...
' - MessageUnload           ...
' - Progress                ...
' - ProgressOld             ...
' - RemoveTempRenamed       ...
'
' Public Properties:
' - NonBreakingSpace        Global non-breaking space character
' - NoOfItemsTotal      r/w The total number of items relevant for the service
' - noOfObjectsOutdated
' - NoOfCommonComponents
' - NoOfItemsServiced   r/w The number of items Exported, Updated, Synchronized
' - NoOfItemsSkipped    r/w The number of items ignored/passed because irrelevant for the service
' - ServicedWbk         r/w
' - ServicedItem          w
' ------------------------------------------------------------------------------------
Public DialogLeft               As Long
Public DialogTop                As Long

Private lNoOfCommonComponents   As Long
Private lNoOfItemsOutdated      As Long
Private lNoOfItemsSkipped       As Long
Private lNoOfItemsServiced      As Long
Private lNoOfItemsTotal         As Long
Private sNoOfItemsServicedNames As String
Private sServicedItemLoged      As String
Private sServicedItemName       As String
Private sServicedItemType       As String
Private wbkServiced             As Workbook
Private sCurrentService         As String
Private sExecTraceFile          As String

Private Sub Class_Initialize()
    sNoOfItemsServicedNames = vbNullString
    Set LogServiced = New clsLog
    Set LogService = New clsLog
End Sub

Private Sub Class_Terminate()
    Set LogServiced = Nothing
    Set LogService = Nothing
    Set Comps = Nothing
    Set CompManDat = Nothing
End Sub

Friend Property Get CurrentService() As String:             CurrentService = sCurrentService:                   End Property

Friend Property Let CurrentService(ByVal s As String):      sCurrentService = s:                                End Property

Public Property Get ExecTraceFile() As String:              ExecTraceFile = sExecTraceFile:                     End Property

Private Property Let ExecTraceFile(ByVal s As String):      sExecTraceFile = s:                                 End Property

Friend Property Get NonBreakingSpace() As String:           NonBreakingSpace = Chr$(160):                       End Property

Friend Property Get NoOfCommonComponents() As Long:         NoOfCommonComponents = lNoOfCommonComponents:       End Property

Friend Property Let NoOfCommonComponents(ByVal l As Long):  lNoOfCommonComponents = l:                          End Property

Friend Property Get NoOfItemsSkipped() As Long:             NoOfItemsSkipped = lNoOfItemsSkipped:               End Property

Friend Property Let NoOfItemsSkipped(ByVal l As Long):      lNoOfItemsSkipped = l:                              End Property

Friend Property Get NoOfItemsOutdated() As Long:            NoOfItemsOutdated = lNoOfItemsOutdated:             End Property

Friend Property Let NoOfItemsOutdated(ByVal l As Long):     lNoOfItemsOutdated = l:                             End Property

Friend Property Get NoOfItemsServiced() As Long:            NoOfItemsServiced = lNoOfItemsServiced:             End Property

Friend Property Let NoOfItemsServiced(ByVal l As Long):     lNoOfItemsServiced = l:                             End Property

Public Property Get NoOfItemsServicedNames() As String:     NoOfItemsServicedNames = sNoOfItemsServicedNames:   End Property

Public Property Let NoOfItemsServicedNames(ByVal s As String)
    If sNoOfItemsServicedNames = vbNullString _
    Or InStr(sNoOfItemsServicedNames, s) <> 0 Then
        sNoOfItemsServicedNames = "(" & s & ")"
    Else
        sNoOfItemsServicedNames = Left(sNoOfItemsServicedNames, Len(sNoOfItemsServicedNames) - 1) _
                                & ", " _
                                & s _
                                & ")"
    End If
    sNoOfItemsServicedNames = Replace(sNoOfItemsServicedNames, ", ,", ",")
    
End Property

Friend Property Get NoOfItemsTotal() As Long:               NoOfItemsTotal = lNoOfItemsTotal:                   End Property

Friend Property Let NoOfItemsTotal(ByVal l As Long):        lNoOfItemsTotal = l:                                End Property

Friend Property Get ServicedWbk() As Workbook
    Const PROC = "ServicedWbk-Get"
    
    On Error GoTo eh
    If wbkServiced Is Nothing _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), _
         "A service Workbook has yet not been provided!" & _
         mErH.CONCAT & "Services.ServicedWbk = <serviced workbook> must have been executed before!"
    
    Set ServicedWbk = wbkServiced
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Let ServicedWbk(ByVal ws_wbk As Workbook)
    Const PROC = "ServicedWbk-Let"
    Set wbkServiced = ws_wbk
    wsService.CurrentServicedWorkbookFullName = wbkServiced.FullName
End Property

Friend Property Let ServicedItem(Optional ByRef s_max_len_name As Long, _
                                 Optional ByRef s_max_len_type As Long, _
                                 Optional ByVal s_type As String, _
                                          ByVal v As Variant)
' -----------------------------------------------------------------------------------
' Determines the serviced item's type and name string for logging and computes the
' maximum length for both.
' -----------------------------------------------------------------------------------
    Const PROC = "ServicedItem-Let"

    On Error GoTo eh
    Dim vbc As VBComponent
    Dim nme  As Name
    Dim wsh  As Worksheet
    Dim ref As Reference
    Dim shp As Shape
    Dim tbt As ToggleButton
    Dim tbx As TextBox
    Dim sbt As SpinButton
    Dim scb As ScrollBar
    Dim obt As OptionButton
    Dim lbx As ListBox
    Dim lbl As Label
    Dim img As Image
    Dim cbt As CommandButton
    Dim rng As Range

    Select Case TypeName(v)
        Case "VBComponent":     Set vbc = v:    Select Case vbc.Type
                                                    Case vbext_ct_ActiveXDesigner:  sServicedItemType = "ActiveX-Designer": sServicedItemName = vbc.Name
                                                    Case vbext_ct_ClassModule:      sServicedItemType = "Class-Module":     sServicedItemName = vbc.Name
                                                    Case vbext_ct_MSForm:           sServicedItemType = "UserForm":         sServicedItemName = vbc.Name
                                                    Case vbext_ct_StdModule:        sServicedItemType = "Standard-Module":  sServicedItemName = vbc.Name
                                                    Case vbext_ct_Document
                                                        If mComp.IsSheetDocMod(vbc, Me.ServicedWbk) Then
                                                                                    sServicedItemType = "Worksheet":        sServicedItemName = vbc.Name
                                                        Else
                                                                                    sServicedItemType = "Workbook":         sServicedItemName = vbc.Name
                                                        End If
                                                End Select
        Case "Name":            Set nme = v:                                        sServicedItemType = "Name":             sServicedItemName = Replace(nme.Name, nme.Parent.Name & "!", vbNullString) & "(" & nme.RefersTo & ")"
        Case "Reference":       Set ref = v:                                        sServicedItemType = TypeName(ref):      sServicedItemName = ref.Description
        Case "Shape":           Set shp = v:                                        sServicedItemType = s_type:             sServicedItemName = shp.Parent.Name & "." & ShapeNames(shp)
        Case "Worksheet":       Set wsh = v:                                        sServicedItemType = "Worksheet":        sServicedItemName = wsh.CodeName
        Case "String":                                                              sServicedItemType = Split(v, ";")(0):   sServicedItemName = Split(v, ";")(1)
        Case "CommandButton":   Set cbt = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "Image":           Set img = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "Label":           Set lbl = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "ListBox":         Set lbx = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "OptionButton":    Set obt = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "ScrollBar":       Set scb = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "SpinButton":      Set sbt = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "TextBox":         Set tbx = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "ToggleButton":    Set tbt = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name
        Case "Range":           Set rng = v:                                        sServicedItemType = TypeName(v):        sServicedItemName = v.Name.Name
        Case Else
            Debug.Print "TypeName '" & TypeName(v) & "' not yet considered as a serviced item"
    End Select

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get ServicedItemName() As String:   ServicedItemName = sServicedItemName:   End Property

Friend Property Get TempExportFolder() As String
    TempExportFolder = mExport.ExpFileFolderPath(ServicedWbk) & "\Temp"
End Property

Private Function CodeModuleIsEmpty(ByRef vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0
        If Not CodeModuleIsEmpty Then CodeModuleIsEmpty = .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Private Function CurrentServiceStatusBar() As String
' ------------------------------------------------------------------------------
' Returns the current services base status bar message.
' ------------------------------------------------------------------------------
    CurrentServiceStatusBar = wsService.CurrentServiceName & " (by "
    
    If ThisWorkbook.Name = mAddin.WbkName _
    Then CurrentServiceStatusBar = CurrentServiceStatusBar & "Add-in" _
    Else CurrentServiceStatusBar = CurrentServiceStatusBar & ThisWorkbook.Name
    
    CurrentServiceStatusBar = CurrentServiceStatusBar & ") for " & Me.ServicedWbk.Name

End Function

Public Function Denied(ByVal d_service As String) As Boolean
' --------------------------------------------------------------------------
' Returns TRUE when all preconditions for a service (d_service) are met.
' --------------------------------------------------------------------------
    Const PROC = "Denied"
    
    On Error GoTo eh
    Dim sStatus         As String
    Dim sDeniedService  As String
    
    Select Case d_service
        Case mCompManClient.SRVC_EXPORT_CHANGED:    sDeniedService = "The service """ & mCompManClient.SRVC_EXPORT_CHANGED_DSPLY & """ is denied! "
        Case mCompManClient.SRVC_SYNCHRONIZE:       sDeniedService = "The service """ & mCompManClient.SRVC_EXPORT_CHANGED_DSPLY & """ is denied! "
        Case mCompManClient.SRVC_UPDATE_OUTDATED:   sDeniedService = "The service """ & mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY & """ is denied! "
    End Select
    Select Case True
        Case WbkIsOpenedRegular = False
            sStatus = sDeniedService & "The serviced Workbook has apparently been restored by the system and yet not saved under its origin name!"
        Case mMe.IsAddinInstnc And mAddin.Paused
            '~~ When the service is about to be provided by the Addin this means that the CompMan.xlsb is not open.
            '~~ When the Addin is currently paused it requires the CompMan.xlsb to continue it and or to provide
            '~~ the service.
            sStatus = sDeniedService & "CompMan Addin is required but currently paused! Open CompMan.xlsb to continue it and re-open CompMan.xlsb to run the service."
        Case mMe.IsDevInstnc And d_service = mCompManClient.SRVC_UPDATE_OUTDATED And mAddin.Paused
            '~~ Note: The CompMan development instance is able to export its modified components but requires the
            '~~       Addin instance to update its own used Common Components when outdated
            sStatus = sDeniedService & "CompMan Addin is available but currently paused!"
        Case FolderNotVbProjectExclusive
            sStatus = sDeniedService & "The Workbook is not the only one in its parent folder!"
        Case Not mCompMan.WinMergeIsInstalled And d_service = mCompManClient.SRVC_UPDATE_OUTDATED
            sStatus = sDeniedService & "WinMerge is required but not installed!"
    End Select
    
    If sStatus <> vbNullString Then
        LogService.Entry sStatus
        With Me
            mCompManClient.Progress p_service_name:=CurrentService _
                                  , p_serviced_wbk_name:=ServicedWbk.Name _
                                  , p_by_servicing_wbk_name:=ThisWorkbook.Name _
                                  , p_service_info:=sStatus
        End With
        Denied = True
    End If
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

'Private Function DifferentWithExpHeaderExcluded(ByVal dhe_fle1 As File, _
'                                                ByVal dhe_fle2 As File) As Boolean
'' ----------------------------------------------------------------------------
''
'' ----------------------------------------------------------------------------
'    Dim dct1    As Dictionary
'    Dim dct2    As Dictionary
'
'    Set dct1 = ExportHeaderRemoved(dhe_fle1)
'    Set dct2 = ExportHeaderRemoved(dhe_fle2)
'
'    DifferentWithExpHeaderExcluded = _
'    mDct.DctDiffers(dd_dct1:=dct1 _
'                  , dd_dct2:=dct2 _
'                  , dd_diff_keys:=False _
'                  , dd_ignore_items_empty:=True _
'                  , dd_ignore_case:=True)
'
'End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "clsServices." & s
End Function

Public Sub EstablishExecTraceFile()
' --------------------------------------------------------------------------
' Establishes a trace log file in the serviced Workbook's parent folder
' provided either the Cond. Comp. Arg. `XcTrc_mTrc = 1` or XcTrc_clsTrc = 1`
' Note: A new trace-log-file is provided with the SRVC_UPDATE_OUTDATED which
'       is the service performed when a serviced Workbook is opened.
' --------------------------------------------------------------------------
    Const PROC = "EstablishExecTraceFile"
    Const TRACE_FILE_NAME   As String = "CompMan.Services.ExecTrace.log"

#If XcTrc_mTrc = 1 Then
    mTrc.FileName = TRACE_FILE_NAME
    mTrc.Path = ServicedWbk.Path
    If CurrentService = mCompManClient.SRVC_UPDATE_OUTDATED Then
        mTrc.NewFile
    End If
    sExecTraceFile = mTrc.FileFullName

#ElseIf XcTrc_clsTrc = 1 Then
    Set Trc = New clsTrc
    With Trc
        .FileName = TRACE_FILE_NAME
        .Path = ServicedWbk.Path
        .Title = CurrentServiceStatusBar
        If Services.CurrentService = mCompManClient.SRVC_UPDATE_OUTDATED Then
            .NewFile
        End If
        sExecTraceFile = .FileFullName
    End With
#End If
End Sub

Public Sub ExportChangedComponents(ByVal e_hosted As String)
' ------------------------------------------------------------------------------
' Exclusively called by mCompMan.ExportChangedComponents, triggered by the
' Before_Save event.
' Attention: When called directly by the user, e.g. via the 'Imediate Window' an
'            error will be raised because an 'Me.ServicedWbk' Workbook is
'            not set.
' ------------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If Me.ServicedWbk Is Nothing _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The procedure '" & ErrSrc(PROC) & "' has been called without a prior set of the 'ServicedWbk' Workbook. " & _
                                            "(it may have been called directly via the 'Immediate Window'"
    mExport.ChangedComponents e_hosted
    
xt: mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ExportHeaderRemoved(ByVal ehr_fle As File) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim bAttribReached  As Boolean
    Dim bAttribPassed   As Boolean
    Dim dct             As Dictionary
    Dim v               As Variant
    Dim s               As String
    
    Set dct = mFso.FileDict(ehr_fle)
    For Each v In dct
        s = dct(v) & "               "
        If Not bAttribReached Then bAttribReached = Left(s, 12) = "Attribute VB"
        If Not bAttribPassed Then bAttribPassed = bAttribReached And Left(s, 12) <> "Attribute VB"
        
        If Not bAttribPassed _
        Or (bAttribReached And bAttribPassed And Trim(s) = vbNullString) Then
            dct.Remove v
        End If
        If bAttribReached And bAttribPassed And Trim(s) <> vbNullString Then Exit For
    Next v
    
    Set ExportHeaderRemoved = dct
    Set dct = Nothing
    
End Function

'Public Function FilesDiffer(ByVal f_file_1 As File, _
'                            ByVal f_file_2 As File, _
'                   Optional ByVal f_ignore_export_header = False) As Boolean
'' ----------------------------------------------------------------------------
'' Returns TRUE when file 1 and file 2 are different whereby case differences
'' and empty lines are ignored. This function guarantees a uniform comparison of
'' export files throughout CompMan.
'' ----------------------------------------------------------------------------
'    Const PROC = "FilesDiffer"
'
'    On Error GoTo eh
'    Dim fl1 As File
'    Dim fl2 As File
'
'    With fso
'        If TypeName(f_file_1) = "File" Then
'            If Not .FileExists(f_file_1) _
'            Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The provided 'f_file_1' does not exist!"
'            Set fl1 = fso.GetFile(f_file_1)
'        ElseIf TypeName(f_file_1) = "Nothing" _
'            Then Err.Raise mBasic.AppErr(2), ErrSrc(PROC), "File 'f_file_1' is not provided!"
'        Else
'            Set fl1 = f_file_1
'        End If
'
'        If TypeName(f_file_2) = "File" Then
'            If Not .FileExists(f_file_2) _
'            Then Err.Raise mBasic.AppErr(2), ErrSrc(PROC), "The provided 'f_file_2' does not exist!"
'            Set fl2 = fso.GetFile(f_file_2)
'        ElseIf TypeName(f_file_2) = "Nothing" Then
'            FilesDiffer = True
'            GoTo xt
'        Else
'            Set fl2 = f_file_2
'        End If
'    End With
'
'    If f_ignore_export_header Then
'        FilesDiffer = DifferentWithExpHeaderExcluded(fl1, fl2)
'    Else
'        FilesDiffer = mFso.FileDiffers(f_file1:=fl1 _
'                                     , f_file2:=fl2 _
'                                     , f_stop_after:=1 _
'                                     , f_ignore_empty_records:=True _
'                                     , f_compare:=vbTextCompare).Count <> 0
'    End If
'xt: Exit Function
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function

Public Function FilesDifference(ByVal f_file_1 As File, _
                                ByVal f_file_2 As File) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with either 0 items when file 1 and file 2 are
' identical or with one item when the two files differ. Empty lines and case
' differences are ignored because the do not constitute a relevant code change
' ----------------------------------------------------------------------------
    Const PROC = "FilesDifference"
    
    On Error GoTo eh
    Dim fl1 As File
    Dim fl2 As File
    
    With fso
        If VarType(f_file_1) = vbString Then
            If Not .FileExists(f_file_1) _
            Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "The provided 'f_file_1' is a string not identifying an existing file!"
            Set fl1 = fso.GetFile(f_file_1)
        Else
            Set fl1 = f_file_1
        End If
        
        If VarType(f_file_2) = vbString Then
            If Not .FileExists(f_file_2) _
            Then Err.Raise mBasic.AppErr(2), ErrSrc(PROC), "The provided 'f_file_2' is a string not identifying an existing file!"
            Set fl2 = fso.GetFile(f_file_2)
        Else
            Set fl2 = f_file_2
        End If
    End With
    
    Set FilesDifference = mFso.FileDiffersFromFile(f_file_this:=fl1 _
                                                 , f_file_from:=fl2 _
                                                 , f_stop_after:=1 _
                                                 , f_exclude_empty:=True _
                                                 , f_compare:=vbTextCompare)
                            
xt: Exit Function
                            
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FolderNotVbProjectExclusive() As Boolean

    Dim fl  As File
    
    For Each fl In fso.GetFolder(Me.ServicedWbk.Path).Files
        If VBA.Left$(fso.GetFileName(fl.Path), 2) = "~$" Then GoTo next_fl
        If VBA.StrComp(fl.Path, Me.ServicedWbk.FullName, vbTextCompare) <> 0 Then
            Select Case fso.GetExtensionName(fl.Path)
                Case "xlsm", "xlam", "xlsb" ' may contain macros, a VB-Project repsectively
                    FolderNotVbProjectExclusive = True
                    Exit For
            End Select
        End If
next_fl:
    Next fl

End Function

Public Sub Initiate(ByVal i_service_proc As String, _
                    ByVal i_serviced_wbk As Workbook, _
           Optional ByVal i_hosted As String = vbNullString, _
           Optional ByRef i_ini As Boolean = True)
' ------------------------------------------------------------------------------
' Provides the initiation for all (regular) services.
'
' Precondition: mCompMan.EstablishExecTraceFile had been executed
' ------------------------------------------------------------------------------
    Const PROC = "Initiate"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    wsService.ClearDataAllServices
    ServicedWbk = i_serviced_wbk
        
    Set CompManDat = New clsCompManDat
    CompManDat.Hosted = i_hosted ' must be initialized first !
    Set CommComps = New clsCommComps
    CommComps.Hosted = i_hosted
    CommComps.Hskpng
    Set Comps = Nothing
    Set Comps = New clsComps ' initialization collect the VB-Component relevant for the current service
    
    '~~ The log-file for the serviced Workbook
    If LogServiced Is Nothing Then Set LogServiced = New clsLog
    With LogServiced
        .KeepDays = 2 ' a new log-file is created after 48 hours
        .WithTimeStamp
        .FileFullName = i_serviced_wbk.Path & "\" & fso.GetBaseName(ThisWorkbook.Name) & ".Services.log"
        .MaxItemLengths Comps.MaxLenServicedType, Comps.MaxLenServicedItem
        .AlignmentItems "|L|L.:|L|"
        mCompMan.LogFileService = .FileName
        wsService.CurrentServiceLogFileFullName = .FileFullName
    End With
    
    '~~ The log-file for the servicing Workbook
    If LogService Is Nothing Then Set LogService = New clsLog
    With LogService
        .WithTimeStamp
        .KeepDays = 10
        .FileFullName = ThisWorkbook.Path & "\" & fso.GetBaseName(ThisWorkbook.Name) & ".ServicesSummary.log"
        .NewLog True ' suppresses delimiter line
        .MaxItemLengths Comps.MaxLenServicedType, Comps.MaxLenServicedItem
        .AlignmentItems "|L|L.:|L|"
        mCompMan.LogFileServicedSummary = .FileName
    End With
    
    wsService.CurrentServiceName = mCompManClient.ServiceName(i_service_proc)
    Select Case i_service_proc
        Case mCompManClient.SRVC_SYNCHRONIZE:       wsService.ClearDataSynchService
                                                    wsSyncLog.Clear
                                                    If i_ini Then mSync.Initiate
        Case mCompManClient.SRVC_UPDATE_OUTDATED:   wsService.ClearDataUpdateService
    End Select
    
    LogServiced.Title CurrentServiceStatusBar
    LogService.Title CurrentServiceStatusBar

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function IsRenamedByCompMan(ByVal comp_name As String) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the component's name indicates that it is one which had been
' renamed by CompMan for an update (rename/import) service.
' ------------------------------------------------------------------------------
    IsRenamedByCompMan = InStr(comp_name, RENAMED_BY_COMPMAN) <> 0
End Function

Public Sub ServicedItemLogEntry(ByVal l_s As String)
' ----------------------------------------------------------------------------
' Writes an entry to the Services.log file by means of clsLog services.
' Supresses repetition of already loged items.
' ----------------------------------------------------------------------------
    With LogService
        If sServicedItemName = sServicedItemLoged Then
            .Entry " ", " ", l_s
            LogServiced.Entry " ", " ", l_s
        Else
            .Entry sServicedItemType, sServicedItemName, l_s
            LogServiced.Entry sServicedItemType, sServicedItemName, l_s
            sServicedItemLoged = sServicedItemName
        End If
    End With
        
End Sub

Public Sub LogEntrySummary(ByVal s As String)
    CompMan.LogService.Entry s, "  (see " & mCompMan.LogFileService & " for log details)"
End Sub

Public Sub MessageUnload(ByVal sm_title As String)
' ----------------------------------------------------------------------------
' Save current message window position and terminate the display of it.
' ----------------------------------------------------------------------------
    DialogTop = mMsg.MsgInstance(sm_title).Top
    DialogLeft = mMsg.MsgInstance(sm_title).Left
    mMsg.MsgInstance sm_title, True
End Sub

Public Function Progress(ByVal p_operation_performed As String) As String
' --------------------------------------------------------------------------
' Returns a progress message for the current service for being displayed in
' the Application.StatusBar in the form:
' <service> (by <servicing>) for <serviced>: <n> of <m> <op> [(component [, component] ..]
' <n> = Number of objects the service has been provided for (p_items_serviced)
' <m> = Total number of objects to be (ptentially) serviced
' <op> = The performed operation
' Whereby the progress is indicated in two ways: an increasing number of
' dots for the items collected for being serviced and a decreasing number
' of dots indication the items already serviced.
'
' Example:
' Export ... (by CompMan....) for ......: 1 of 50 exported (clsServices)
' --------------------------------------------------------------------------
    Const PROC          As String = "Progress"
    
    On Error GoTo eh
    Dim sMsg    As String
    Dim lTotalComps As Long
    
    If lNoOfItemsOutdated <> 0 Then
        lTotalComps = lNoOfItemsOutdated
    Else
        lTotalComps = lNoOfItemsTotal
    End If
    mCompManClient.Progress p_service_name:=CurrentService _
                          , p_serviced_wbk_name:=ServicedWbk.Name _
                          , p_by_servicing_wbk_name:=ThisWorkbook.Name _
                          , p_progress_figures:=True _
                          , p_no_comps_serviced:=lNoOfItemsServiced _
                          , p_no_comps_skipped:=lNoOfItemsSkipped _
                          , p_no_comps_total:=lTotalComps _
                          , p_service_op:=p_operation_performed _
                          , p_service_info:=sNoOfItemsServicedNames
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub RemoveTempRenamed()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "RemoveTempRenamed"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    With Me.ServicedWbk.VBProject
        For Each vbc In .VBComponents
            If Services.IsRenamedByCompMan(vbc.Name) Then
                .VBComponents.Remove vbc
            End If
        Next vbc
    End With

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function SelectServicedWrkbk(ByVal gs_service As String) As Workbook
    Dim fl As File
    
    If mFso.FilePicked(p_title:="Select the Workbook (may already be open, will be opened if not) to be served by the " & gs_service & " service" _
                  , p_filters:="Excel Workbook,*.xl*" _
                  , p_file:=fl) _
    Then Set SelectServicedWrkbk = mCompMan.WbkGetOpen(fl.Path) _
    Else: Set SelectServicedWrkbk = Nothing

End Function

Public Sub TempExportFolderRemove()
    With New FileSystemObject
        If .FolderExists(TempExportFolder) Then .DeleteFolder TempExportFolder
    End With
End Sub

Private Function WbkIsOpenedRegular() As Boolean
' ----------------------------------------------------------------------------
' Retrurns FALSE when the Workbook had been restored by Excel or opended as
' a version - which will cause the denial of any service.
' ----------------------------------------------------------------------------
    WbkIsOpenedRegular = ActiveWindow.Caption = Me.ServicedWbk.Name
End Function

