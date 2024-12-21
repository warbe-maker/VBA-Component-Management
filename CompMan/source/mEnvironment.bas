Attribute VB_Name = "mEnvironment"
Option Explicit
' ------------------------------------------------------------------------------
' Class Module clsEnvironment: Provides an up-to-date file/folder environment
' ============================ by considering the any initial state may be
' forwarded to a current state when specified folder/file locations/names have
' changed.
' ------------------------------------------------------------------------------
'~~ Current environment file and folder names
Public Const FILE_NAME_COMMCOMPS_PENDING_PRIVPROF  As String = "PendingReleases.dat"
Public Const FILE_NAME_COMMCOMPS_PUBLIC_PRIVPROF   As String = "CommComps.dat"
Public Const FILE_NAME_COMMCOMPS_SERVICED_PRIVPROF As String = "CommComps.dat"
Public Const FILE_NAME_COMPMAN_LOCAL_CONFIG        As String = "CompMan.cfg"
Public Const FILE_NAME_EXEC_TRACE                  As String = "ExecTrace.log"
Public Const FILE_NAME_SERVICES_LOG                As String = "Services.log"
Public Const FILE_NAME_SERVICES_SUMMARY_LOG        As String = "ServicesSummary.log"
    
Public Const FLDR_NAME_ADDIN                       As String = "CompManAddin"
Public Const FLDR_NAME_COMMON_COMPONENTS           As String = "Common-Components"
Public Const FLDR_NAME_COMMON_COMPONENTS_PENDING   As String = "PendingReleases"
Public Const FLDR_NAME_EXPORT                      As String = "source"
Public Const FLDR_NAME_SERVICED_COMPMAN_SERVICES   As String = "CompMan"

Private sAddInFolderPath                        As String
Private sCommCompsFolderPath                    As String
Private sCommCompsPendingFolderPath             As String
Private sCommCompsPendingPrivProfFileFullName   As String
Private sCommCompsPublicPrivProfFileFullName    As String
Private sCommCompsServicedPrivProfFileFullName  As String
Private sCompManServiceFolderPath               As String
Private sExportServiceFolderPath                As String
Private sServicedExecTraceLogFileFullName       As String
Private sServicesLogFileFullName                As String
Private sServicesSummaryLogFileFullName         As String
Private sCompManLocalConfigFileFullName         As String
Private sCompManServicedRoot                    As String
Private sCompManDedicatedFolder                 As String

Public Property Get CompManDedicatedFolder() As String:                 CompManDedicatedFolder = sCompManDedicatedFolder:                               End Property

Public Property Get CompManLocalConfigFileFullName() As String:         CompManLocalConfigFileFullName = sCompManLocalConfigFileFullName:               End Property

Public Property Get CompManServicedRootFolder() As String:              CompManServicedRootFolder = sCompManServicedRoot:                               End Property

Public Property Get AddInFolderPath() As String:                        AddInFolderPath = sAddInFolderPath:                                             End Property

Public Property Get CommCompsPath() As String:                          CommCompsPath = sCommCompsFolderPath:                                           End Property
    
Public Property Get CommCompsPendingPath() As String:                   CommCompsPendingPath = sCommCompsPendingFolderPath:                             End Property

Public Property Get CommCompsPendingPrivProfFileFullName() As String:   CommCompsPendingPrivProfFileFullName = sCommCompsPendingPrivProfFileFullName:   End Property

Public Property Get CommCompsPublicPrivProfFileFullName() As String:    CommCompsPublicPrivProfFileFullName = sCommCompsPublicPrivProfFileFullName:     End Property

Public Property Get CompManServiceFolder() As String:                   CompManServiceFolder = sCompManServiceFolderPath:                               End Property

Public Property Get ExportServiceFolderPath() As String:                ExportServiceFolderPath = sExportServiceFolderPath:                             End Property

Public Property Get CommCompsServicedPrivProfFileFullName() As String:  CommCompsServicedPrivProfFileFullName = sCommCompsServicedPrivProfFileFullName: End Property

Public Property Get ServicedExecTraceLogFileFullName() As String:       ServicedExecTraceLogFileFullName = sServicedExecTraceLogFileFullName:           End Property

Public Property Get ServicesLogFileFullName() As String:                ServicesLogFileFullName = sServicesLogFileFullName:                             End Property

Public Property Get ServicesSummaryLogFileFullName() As String:         ServicesSummaryLogFileFullName = sServicesSummaryLogFileFullName:               End Property

Public Function ThisComputersName() As String:                          ThisComputersName = Environ("COMPUTERNAME"):                                    End Function

Public Function ThisComputersUser() As String:                          ThisComputersUser = Environ("USERNAME"):                                        End Function

Public Function PrivateProfileFileFooter() As String
' --------------------------------------------------------------------------
' Returns the footer written into any Private Profile file maintained by
' CompMan
' --------------------------------------------------------------------------
    PrivateProfileFileFooter = "This Private Profile file is maintained by services provided by the class module ""clsPrivProf""" & vbCrLf & _
                               "(see the public Common Component in https://github.com/warbe-maker/VBA-Private-Profile)." & vbCrLf & _
                               "Main services are: Section separation, optional file header/footer, sections and value-names" & vbCrLf & _
                               "are maintained in alphabetical ascending order, optional section-/value headers/comments."
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsEnvironment" & "." & sProc
End Function

Private Sub EstablishServicedExecTraceLogFile()
' --------------------------------------------------------------------------
' - Establishes the execution trace log file for the serviced Workbook in
'   the serviced Workbook's CompMan.ServicesFolder.
' - Considers the mTrc Standard Module is used (Cond. Comp. Arg. `mTrc = 1`)
'   or the clsTrc class module is used (Cond. Comp. Arg. `clsTrc = 1`)
' - Establishes a new log file along with the update service only (else the
'   log is appended).
' --------------------------------------------------------------------------

#If mTrc = 1 Then
    mTrc.FileFullName = mEnvironment.ServicedExecTraceLogFileFullName
    mTrc.KeepLogs = 5
    mTrc.Title = Servicing.CurrentServiceStatusBar
    If sCurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY _
    Then mTrc.NewFile
    sExecTraceFile = mTrc.FileFullName

#ElseIf clsTrc = 1 Then
    Set Trc = Nothing: Set Trc = New clsTrc
    With Trc
        .FileFullName = mEnvironment.ServicedExecTraceLogFileFullName
        .KeepLogs = 5
        .Title = mCompMan.CurrentServiceStatusBar
        If mCompMan.CurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY Then
            .NewFile
        End If
    End With
#End If

End Sub

Public Sub EstablishServicedServicesLog(ByVal e_max_len_type As Long, _
                                        ByVal e_max_len_item As Long)
' --------------------------------------------------------------------------
' Establish the serviced Workbook's services log file.
' --------------------------------------------------------------------------
    
    Set LogServiced = New clsLog
    With LogServiced
        .KeepLogs = 5
        .WithTimeStamp
        .FileFullName = mEnvironment.ServicesLogFileFullName
        .Title CurrentServiceStatusBar
        .ColsSpecs = "L" & e_max_len_type & ", L.:" & e_max_len_item & ", L60"
        mCompMan.LogFileService = .FileFullName
    End With
    
End Sub

Public Sub EstablishServicesSummaryLog(ByVal e_max_len_type As Long, _
                                       ByVal e_max_len_item As Long)
' --------------------------------------------------------------------------
' Establish the services summary log file.
' --------------------------------------------------------------------------
    Set LogServicesSummary = New clsLog
    With LogServicesSummary
        .WithTimeStamp
        .KeepLogs = 10
        .FileFullName = mEnvironment.ServicesSummaryLogFileFullName
        .NewLog
        .ColsSpecs = "L" & e_max_len_type & ", L.:" & e_max_len_item & ", L60"
        mCompMan.LogFileServicesSummary = .FileFullName
    End With

End Sub

Private Function History(ByVal e_lctn As String, _
                         ByVal e_name As String) As Collection
' ------------------------------------------------------------------------------
' Returns a collection of history files or folders (e_lctn) & "\" & (e_name)
' with the current as first and history items as subsequent items.
' ------------------------------------------------------------------------------
    Dim cll   As New Collection
    Dim vLctn As Variant
    Dim vName As Variant
    Dim s     As String
    
    With FSo
        For Each vLctn In ItemHistory(e_lctn)
            For Each vName In ItemHistory(e_name)
                '~~ The first location and the first name indicate the current item (folder or file)
                s = vLctn & "\" & vName
                cll.Add s
'                Debug.Print "History " & cll.Count & " = " & cll(cll.Count)
                If .FolderExists(s) Or .FileExists(s) Then Exit For
            Next vName
            If .FolderExists(s) Or .FileExists(s) Then
                Exit For
            End If
        Next vLctn
    End With
    
    Set History = cll
    Set cll = Nothing
    
End Function

Private Function HistoryForwarded(ByVal e_lctn As String, _
                                  ByVal e_name As String, _
                         Optional ByVal e_create As Boolean = True, _
                         Optional ByVal e_folder As Boolean = True) As String
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim cll         As Collection
    Dim sCurrent    As String
    Dim sHistory    As String

    Set cll = History(e_lctn, e_name)
    '~~ The collection contains at least one item which represents the - to be - current file or folder.
    sCurrent = cll(1)
    HistoryForwarded = sCurrent
    If cll.Count = 2 Then sHistory = cll(2)
    With FSo
        If .FolderExists(sCurrent) Or .FileExists(sCurrent) Then GoTo xt
        Select Case True
            Case .FolderExists(sHistory)               ' Forward the folder's location and/or name
                .MoveFolder sHistory, sCurrent
                GoTo xt
            Case .FileExists(sHistory)               ' Forward the file's location and/or name
                .MoveFile sHistory, sCurrent
                GoTo xt
            Case Else
                '~~ There was no history able to be forwarded for vecomin up-to-date
                If e_folder And e_create Then FSo.CreateFolder sCurrent
        End Select
    End With
    
xt: Exit Function
End Function

Private Function ItemHistory(ByVal e_hist As String) As Collection
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------

    Dim a   As Variant
    Dim cll As New Collection
    Dim i   As Long
    Dim v   As Variant
    
    If InStr(e_hist, ">") <> 0 Then
        a = Split(e_hist, ">")
        For i = UBound(a) To LBound(a) Step -1
            cll.Add a(i)
        Next i
    Else
        a = Split(e_hist, "<")
        For Each v In a
            cll.Add v
        Next v
    End If
    Set ItemHistory = cll
    Set cll = Nothing

End Function

Public Sub Provide(Optional ByVal p_create As Boolean = False)
' ------------------------------------------------------------------------------
' Maintains and provides an up-to-date CompMan environment for CompMan itself
' and the serviced Workbook.
'
' Item/structure                            Current name/location
' ----------------------------------------- ------------------------------------
' <compman-serviced-root-folder>            ThisWorkbook.Path.Parent
'  +--Common-Components                     public Common Components
'  |  +--CommComps.dat                      public Common Components' properties
'  |  +--PendingReleases.dat                only exists when there are pendings
'  |  +--PendingReleases                    only exists when there are pendings
'  |  +--<common-component-export-files>
'  |  +--...
'  +--CompMan
'  |  +--CompMan.xlsb
'  |  +--CompMan
'  |     +--source                          Export folder
'  |     +--ExecTrace.log
'  |     +--CommComps.dat                   Used/hoste Common Components properties
'  |     +--Service.log
'  |     +--ServiceSummary.log
'  |
'  +--CompManAddin                          when configured
'  |  +--CompMan.xlam                       when configured
'  |
'  +--<any-serviced-workbooks-parent-folder>
'     +--CompMan                                CompMan-Service-Folder
'        +--source                              Export folder
'        +--ExecTrace.log
'        +--CommComps.dat
'        +--Service.log
' ------------------------------------------------------------------------------
    Const PROC = "Provide"
        
    On Error GoTo eh
    Dim sLctn   As String
    Dim sName   As String
    
    '~~ CompMan-serviced root folder
    '~~ Attention! This is the only folder which has no fixed configuration. By definition it is
    '~~            the CurrentWorkbook's parent parent folder. I.e. this folder may be renamed
    '~~            or moved to another location at any point in time. Also moving the whole content
    '~~            to another folder is possible.
    sCompManServicedRoot = FSo.GetFolder(ThisWorkbook.Path).ParentFolder.Path
    
    '~~ CompManServicedRoot: AddIn folder
    sLctn = sCompManServicedRoot
    sName = FLDR_NAME_ADDIN
    sAddInFolderPath = HistoryForwarded(sLctn, sName, p_create)
    
    '~~  CompManServicedRoot: CompMan Workbook's own dedicated folder
    sLctn = sCompManServicedRoot
    sName = FSo.GetBaseName(ThisWorkbook.Name)
    sCompManDedicatedFolder = HistoryForwarded(sLctn, sName, p_create)
    
    '~~  CompManServicedRoot: Common Components folder
    sLctn = sCompManServicedRoot
    sName = FLDR_NAME_COMMON_COMPONENTS
    sCommCompsFolderPath = HistoryForwarded(sLctn, sName, p_create)
    
    '~~ Common Components folder: Common Components Public Private Profile file
    sLctn = sCommCompsFolderPath
    sName = FILE_NAME_COMMCOMPS_PUBLIC_PRIVPROF
    sCommCompsPublicPrivProfFileFullName = HistoryForwarded(sLctn, sName, False)
    Set CommonPublic = New clsCommonPublic
    
    '~~ Common Components folder: Common Components Peding folder
    sLctn = sCommCompsFolderPath
    sName = FLDR_NAME_COMMON_COMPONENTS_PENDING
    sCommCompsPendingFolderPath = HistoryForwarded(sLctn, sName, False)
    
    '~~ Common Components folder: Common Components Pending Release Private Profile file
    sLctn = sCommCompsFolderPath
    sName = FILE_NAME_COMMCOMPS_PENDING_PRIVPROF
    sCommCompsPendingPrivProfFileFullName = HistoryForwarded(sLctn, sName, False)
    Set CommonPending = New clsCommonPending

    '~~ Serviced Workbook's related environment resources
    If ActiveWorkbook.Path <> vbNullString Then
        '~~ Serviced Workbook's CompMan service folder
        sLctn = ActiveWorkbook.Path
        sName = FLDR_NAME_SERVICED_COMPMAN_SERVICES
        sCompManServiceFolderPath = HistoryForwarded(sLctn, sName, p_create)
    
        '~~ CompMan.cfg file
        sLctn = sCompManServiceFolderPath
        sName = FILE_NAME_COMPMAN_LOCAL_CONFIG
        sCompManLocalConfigFileFullName = HistoryForwarded(sLctn, sName, False)
              
        '~~ Serviced Workbook's Export folder
        sLctn = ActiveWorkbook.Path & ">" & sCompManServiceFolderPath
        sName = "source"
        sExportServiceFolderPath = HistoryForwarded(sLctn, sName, p_create)
        
        '~~ Services Summary log file
        sLctn = sCompManServiceFolderPath
        sName = FILE_NAME_SERVICES_SUMMARY_LOG
        sServicesSummaryLogFileFullName = HistoryForwarded(sLctn, sName, False, False)
    
        '~~ Serviced Workbook's Execution Trace log file
        sLctn = ActiveWorkbook.Path & ">" & sCompManServiceFolderPath
        sName = FILE_NAME_EXEC_TRACE & "<CompMan.Service.trc"
        sServicedExecTraceLogFileFullName = HistoryForwarded(sLctn, sName, False, False)
        EstablishServicedExecTraceLogFile
        
        '~~ Serviced Workbook's Common Components Private Profile file
        sLctn = ActiveWorkbook.Path & ">" & sCompManServiceFolderPath
        sName = FILE_NAME_COMMCOMPS_SERVICED_PRIVPROF & "<CompMan.dat"
        sCommCompsServicedPrivProfFileFullName = HistoryForwarded(sLctn, sName, False, False)
        Set CommonServiced = New clsCommonServiced
        
        '~~ Serviced Workbook's Services log file
        sLctn = ActiveWorkbook.Path & ">" & sCompManServiceFolderPath
        sName = FILE_NAME_SERVICES_LOG & "<CompMan.Services.log"
        sServicesLogFileFullName = HistoryForwarded(sLctn, sName, False, False)
    
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RenameMove(ByVal e_path_old As String, _
                      ByVal e_path_new As String)
' ------------------------------------------------------------------------------
' Note: Does nothing when the old item (e_path_old) neither identifies an
'       existing folder nor an existing file.
' ------------------------------------------------------------------------------
    
    With New FileSystemObject
        If .FileExists(e_path_old) _
        Then .MoveFile e_path_old, e_path_new _
        Else .MoveFolder e_path_old, e_path_new
    End With

End Sub



