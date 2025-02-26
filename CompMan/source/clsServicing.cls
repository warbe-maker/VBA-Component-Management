VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServicing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------------------------------
' Class Module clsServicing: Common services used throughout this VB-Project.
' =========================

' Public services:
' ----------------
' - CurrentServiceStatusBar ...
' - Denied                  ...
' - DsplyStatus             ...
' - EstablishExecTraceFile  ...
' - ExpFilesDiffDisplay     ...
' - ExportChangedComponents ....
' - FilesDiffer             ...
' - FilesDifference         ...
' - Initiate                ...
' - Log    ...
' - LogServicingEntry       ...
' - MessageUnload           ...
' - Progress                ...
' - ProgressOld             ...
' - RemoveTempRenamed       ...
'
' Public Properties:
' - NonBreakingSpace        Global non-breaking space character
' - noOfObjectsOutdated
' - NoOfCommonComponents
' - NoOfItemsServicedIncrement The number of items Exported, Updated, Synchronized
' - NoOfItemsSkipped    r/w The number of items ignored/passed because irrelevant for the service
' - ServicedWbk         r/w
' - LogItem          w
' ------------------------------------------------------------------------------------
Public DialogLeft               As Long
Public DialogTop                As Long

Private lNoOfCommonComponents   As Long
Private lNoOfItemsOutdated      As Long
Private lNoOfItemsSkipped       As Long
Private lNoOfItemsServiced      As Long
Private lTotalComps             As Long
Private sNoOfItemsServicedNames As String
Private sInstanceName           As String
Private sServicedItemLogged      As String
Private sServicedItemName       As String
Private sServicedItemType       As String
Private wbkServiced             As Workbook
Private vCurrentService         As Variant

Private Sub Class_Initialize()
    sNoOfItemsServicedNames = vbNullString
    sInstanceName = ThisWorkbook.Name
    Set LogServiced = New clsLog
    lTotalComps = ActiveWorkbook.VBProject.VBComponents.Count
End Sub

Private Sub Class_Terminate()
    Set LogServiced = Nothing
    Set LogServicesSummary = Nothing
    Set CommonServiced = Nothing
End Sub

Friend Property Get InstanceName() As String:               InstanceName = sInstanceName:                       End Property

Friend Property Get CurrentService() As Variant:            CurrentService = vCurrentService:                   End Property

Friend Property Let CurrentService(ByVal v As Variant):     vCurrentService = v:                                End Property

Public Property Let Log(Optional ByVal s_Item As Variant, _
                                 ByVal s_log As String)
' ----------------------------------------------------------------------------
' Writes an entry to the Services.log file concerning an item (s_item) by
' using clsLog services. Supresses repetition of already loged items.
' When the item (s_item) is of type string it is regarded a VBProject
' component's name, else a VBProject component object.
' ----------------------------------------------------------------------------
    Static lMaxLenItems As Long
    
    If LogServicesSummary Is Nothing Then Set LogServicesSummary = New clsLog
    If TypeName(s_Item) = "String" _
    Then LogItem = Serviced.Wrkbk.VBProject.VBComponents(s_Item) _
    Else LogItem = s_Item
    If lMaxLenItems = 0 Then lMaxLenItems = Serviced.MaxLenItem
    
    With LogServicesSummary
        If sServicedItemName = sServicedItemLogged Then
            .ColsItems " ", " ", s_log
            LogServiced.ColsItems " ", " ", s_log
        Else
            .ColsItems sServicedItemType, sServicedItemName, s_log
            LogServiced.ColsItems sServicedItemType, sServicedItemName, s_log
            sServicedItemLogged = sServicedItemName
        End If
    End With
        
End Property

Friend Property Let LogItem(ByVal s_serviced As Variant)
' -----------------------------------------------------------------------------------
' Determines the serviced log item's type and name string and computes the
' maximum length for both.
' -----------------------------------------------------------------------------------
    Const PROC = "LogItem-Let"

    Dim vbc As VBComponent

    Select Case TypeName(s_serviced)
        Case "VBComponent":     Set vbc = s_serviced:    Select Case vbc.Type
                                                    Case vbext_ct_ActiveXDesigner:  sServicedItemType = "ActiveX-Designer":        sServicedItemName = vbc.Name
                                                    Case vbext_ct_ClassModule:      sServicedItemType = "Class-Module":            sServicedItemName = vbc.Name
                                                    Case vbext_ct_MSForm:           sServicedItemType = "UserForm":                sServicedItemName = vbc.Name
                                                    Case vbext_ct_StdModule:        sServicedItemType = "Standard-Module":         sServicedItemName = vbc.Name
                                                    Case vbext_ct_Document
                                                        If mComp.IsSheetDocMod(vbc, Me.ServicedWbk) Then
                                                                                    sServicedItemType = "Worksheet":               sServicedItemName = vbc.Name
                                                        Else
                                                                                    sServicedItemType = "Workbook":                sServicedItemName = vbc.Name
                                                        End If
                                                End Select
        Case "String":                                                              sServicedItemType = Split(s_serviced, ";")(0): sServicedItemName = Split(s_serviced, ";")(1)
        Case Else
            Debug.Print ErrSrc(PROC) & ": " & "TypeName '" & TypeName(s_serviced) & "' not yet considered as a serviced item"
    End Select

End Property

Friend Property Get NonBreakingSpace() As String:                   NonBreakingSpace = Chr$(160):                       End Property

Friend Property Get NoOfCommonComponents() As Long:                 NoOfCommonComponents = lNoOfCommonComponents:       End Property

Friend Property Let NoOfCommonComponents(ByVal l As Long):          lNoOfCommonComponents = l:                          End Property

Friend Property Get NoOfItemsOutdated() As Long:                    NoOfItemsOutdated = lNoOfItemsOutdated:             End Property

Friend Property Let NoOfItemsOutdated(ByVal l As Long):             lNoOfItemsOutdated = l:                             End Property

Public Property Get NoOfItemsServicedNames() As String:             NoOfItemsServicedNames = sNoOfItemsServicedNames:   End Property

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

Friend Property Get NoOfItemsSkipped() As Long:             NoOfItemsSkipped = lNoOfItemsSkipped:                       End Property

Friend Property Let NoOfItemsSkipped(ByVal l As Long):      lNoOfItemsSkipped = l:                                      End Property

Friend Property Get ServicedItemName() As String:           ServicedItemName = sServicedItemName:                       End Property

Friend Property Get ServicedWbk() As Workbook:              Set ServicedWbk = wbkServiced:                              End Property

Friend Property Let ServicedWbk(ByVal s_wbk As Workbook):   Set wbkServiced = s_wbk:                                    End Property

Public Function TempExportFolder() As String
    TempExportFolder = mEnvironment.ExportServiceFolderPath & "\Temp"
End Function

Private Function CodeModuleIsEmpty(ByRef vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0
        If Not CodeModuleIsEmpty Then CodeModuleIsEmpty = .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
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
        LogServicesSummary.Entry sStatus
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
    ErrSrc = "clsServicing." & s
End Function

Public Sub ExportChangedComponents()
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
    mExport.ChangedComponents
    
xt: mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

Public Sub Initiate(ByVal i_wbk_serviced As Workbook)
' ------------------------------------------------------------------------------
' Provides the initiation for all (regular) services.
'
' Precondition: mCompMan.EstablishExecTraceFile had been executed
' ------------------------------------------------------------------------------
    Const PROC = "Initiate"
    
    On Error GoTo eh
    ServicedWbk = i_wbk_serviced
    LogServicesSummary.Title CurrentServiceStatusBar

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub LogEntrySummary(ByVal l_s As String)
    Dim s As String
    s = l_s & "  (see " & mCompMan.LogFileService & " for log details)"
    s = Replace(s, " ", Chr$(160))
    CompMan.LogServicesSummary.Entry s
End Sub

Public Sub MessageUnload(ByVal sm_title As String)
' ----------------------------------------------------------------------------
' Save current message window position and terminate the display of it.
' ----------------------------------------------------------------------------
    DialogTop = mMsg.Instance(sm_title).Top
    DialogLeft = mMsg.Instance(sm_title).Left
    mMsg.Instance sm_title, True
End Sub

Public Sub NoOfItemsServicedIncrement():                            lNoOfItemsServiced = lNoOfItemsServiced + 1:        End Sub

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
' Export ... (by CompMan....) for ......: 1 of 50 exported (clsServicing)
' --------------------------------------------------------------------------
    Const PROC = "Progress"
    
    On Error GoTo eh
    Dim l As Long
    
    If lNoOfItemsOutdated <> 0 _
    Then l = lNoOfItemsOutdated _
    Else l = lTotalComps
    
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

Private Function SelectServicedWrkbk(ByVal gs_service As String) As Workbook
    Dim fl As File
    
    If mFso.FilePicked(p_title:="Select the Workbook (may already be open, will be opened if not) to be served by the " & gs_service & " service" _
                  , p_filters:="Excel Workbook,*.xl*" _
                  , p_file:=fl) _
    Then Set SelectServicedWrkbk = mCompMan.WbkGetOpen(fl.Path) _
    Else: Set SelectServicedWrkbk = Nothing

End Function

Private Function WbkIsOpenedRegular() As Boolean
' ----------------------------------------------------------------------------
' Retrurns FALSE when the Workbook had been restored by Excel or opended as
' a version - which will cause the denial of any service.
' ----------------------------------------------------------------------------
    WbkIsOpenedRegular = ActiveWindow.Caption = Me.ServicedWbk.Name
End Function

