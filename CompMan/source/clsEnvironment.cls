VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------------------------
' Class Module clsEnvironment: Provides an up-to-date file/folder environment
' ============================ by considering any changes of an initial state
' forwarded to current facts regarding location and name.
'
' ------------------------------------------------------------------------------
Private sCommCompsFolderPath                    As String   ' final after environment housekeeping
Private sCommCompsPendingFolderPath             As String   ' final after environment housekeeping
Private sCommCompsPendingPrivProfFileFullName   As String
Private sCommCompsServicedPrivProfFileFullName  As String
Private sCompManServiceFolderPath               As String   ' final after environment housekeeping
Private sExecTraceLogFileFullName               As String   ' final after environment housekeeping
Private sExportServiceFolderPath                As String   ' final after environment housekeeping
Private sServicedExecTraceLogFileFullName       As String
Private sServicesLogFileFullName                As String   ' final after environment housekeeping

Private Sub Class_Initialize()
    Manage
End Sub

Public Property Get CommCompsPath() As String:                          CommCompsPath = sCommCompsFolderPath:                                           End Property
    
Public Property Get CommCompsPendingPath() As String:                   CommCompsPendingPath = sCommCompsPendingFolderPath:                             End Property

Public Property Get CommCompsPendingPrivProfFileFullName() As String:   CommCompsPendingPrivProfFileFullName = sCommCompsPendingPrivProfFileFullName:   End Property

Public Property Get CompManServiceFolder() As String:                   CompManServiceFolder = sCompManServiceFolderPath:                               End Property

Public Property Get ExportServiceFolderPath() As String:                ExportServiceFolderPath = sExportServiceFolderPath:                             End Property

Public Property Get ServicedExecTraceLogFileFullName() As String:       ServicedExecTraceLogFileFullName = sServicedExecTraceLogFileFullName:           End Property

Public Property Get ServicesLogFileFullName() As String:                ServicesLogFileFullName = sServicesLogFileFullName:                             End Property

Public Function ThisComputersName() As String:                          ThisComputersName = Environ("COMPUTERNAME"):                                    End Function

Public Function ThisComputersUser() As String:                          ThisComputersUser = Environ("USERNAME"):                                        End Function

Private Function CommCompsServicedPrivProfFileHeader() As String
    CommCompsServicedPrivProfFileHeader = _
    "Used or hosted Common Components in the corresponding serviced Workbook """ & Serviced.Wrkbk.Name & """." & vbLf & _
    "The values are provided when a Common Component's code has been modified in the serviced Workbook or" & vbLf & _
    "when a used or hosted Common Component has been updated with code modified elsewhere." & vbLf & _
    "- LastModAt           : Date/time of the last modification date/time (the export file's creation date repectively)" & vbLf & _
    "- LastModBy           : User, which had made the last modification" & vbLf & _
    "- LastModExpFileOrigin: Indicates the 'origin'! of the export file (may point to an export file not available on or not accessable by the used compunter)" & vbLf & _
    "- LastModIn           : The Workbook/VB-Project in which the last code modification had been made (may point to a Workbook om another computer)" & vbLf & _
    "- LastModOn           : The computer on which the last modification had been made in the above Workbook."

End Function


Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsEnvironment" & "." & sProc
End Function

Private Sub EstablishCommCompsServicedPrivProfFile(ByVal e_full_name)
    
    Set CommonServiced = New clsCommonServiced
    With CommonServiced
        Set .PPFile = New clsPrivProf
        With .PPFile
            .FileName = e_full_name
        
            ' Used housekeeping names syntax:
            ' [current-name]<outdated-name<outdated-name...
            ' < indicates the history sequence
            ' > followed by nothing indicate an obsolete name
            .HskpngNames .FileName _
                       , VALUE_NAME_LAST_MOD_AT & "<LastUpdateAtDatTime<LastModAtDatTime<LastModifiedAtDatTime<RawRevisionNumber" _
                       , VALUE_NAME_LAST_MOD_EXP_FILE_ORIG & "<RawExpFileFullName" _
                       , VALUE_NAME_LAST_MOD_IN & "<LastModInWrkbook<LastUpdateModIn<LastUpdateModIn<LastModInWrkbkFullName" _
                       , VALUE_NAME_LAST_MOD_ON & "<LastUpdateModOn<LastModOnMachine" _
                       , "DueModificationWarning>" _
                       , "UsedExportFolder>" _

            If .FileHeader <> CommCompsServicedPrivProfFileHeader Then
                .FileHeader = CommCompsServicedPrivProfFileHeader
            End If
        End With
    End With
    
End Sub

Private Sub EstablishServicedExecTraceLogFile(ByVal e_full_name As String)
' --------------------------------------------------------------------------
' - Establishes the execution trace log file for the serviced Workbook in
'   the serviced Workbook's CompMan.ServicesFolder.
' - Considers the mTrc Standard Module is used (Cond. Comp. Arg. `mTrc = 1`)
'   or the clsTrc class module is used (Cond. Comp. Arg. `clsTrc = 1`)
' - Establishes a new log file along with the update service only (else the
'   log is appended).
' --------------------------------------------------------------------------

#If mTrc = 1 Then
    mTrc.FileFullName = ServicesTraceLogFileFullName
    mTrc.KeepLogs = 5
    mTrc.Title = Services.CurrentServiceStatusBar
    If sCurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY _
    Then mTrc.NewFile
    sExecTraceFile = mTrc.FileFullName

#ElseIf clsTrc = 1 Then
    Set Trc = Nothing: Set Trc = New clsTrc
    With Trc
        .FileFullName = e_full_name
        .KeepLogs = 5
        .Title = mCompMan.CurrentServiceStatusBar
        If mCompMan.CurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY Then
            .NewFile
        End If
    End With
#End If

End Sub

Public Sub EstablishServicesLog(ByVal e_max_len_type As Long, _
                                ByVal e_max_len_item As Long)
' --------------------------------------------------------------------------
' Establish the serviced Workbook's services log file.
' --------------------------------------------------------------------------
    
    Set LogServiced = New clsLog
    With LogServiced
        .KeepLogs = 5
        .WithTimeStamp
        .FileFullName = Environment.ServicesLogFileFullName
        .Title CurrentServiceStatusBar
        .ColsSpecs = "L" & e_max_len_type & ", L.:" & e_max_len_item & ", L60"
        mCompMan.LogFileService = .FileName
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
        .FileFullName = mCompMan.ServicesFolder & "\ServicesSummary.log"
        .NewLog
        .ColsSpecs = "L" & e_max_len_type & ", L.:" & e_max_len_item & ", L60"
        mCompMan.LogFileServicesSummary = .FileName
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
                Debug.Print "History " & cll.Count & " = " & cll(cll.Count)
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

Public Sub Manage()
' ------------------------------------------------------------------------------
' Maintains and provides an up-to-date CompMan environment for CompMan itself
' and the serviced Workbook.
'
' Item/structure                            Current name/location
' ----------------------------------------- ------------------------------------
' <service-folder>                          <compman-root>\CompMan
'  +- <service-trace-log-file>              <compman-root>\CompMan\ExecTrace.log
'  +- <commm-comps-private-profile-file>    <compman-root>\CompMan\CompMan.dat
'  +- <export-folder>                       <compman-root>\source
'  +- <service-log-file>                    <compman-root>\CompMan\Service.log
' ------------------------------------------------------------------------------
    Const PROC = "Manage"
    
    '~~ Current environment file and folder names
    Const FILE_NAME_COMMCOMPS_PENDING_PRIVPROF  As String = "PendingReleases.dat"
    Const FILE_NAME_COMMCOMPS_SERVICED_PRIVPROF As String = "CommComps.dat"
    Const FILE_NAME_EXEC_TRACE                  As String = "ExecTrace.log"
    Const FILE_NAME_SERVICES_LOG                As String = "Services.log"
    Const FLDR_NAME_COMMCOMPS                   As String = "Common-Components"
    Const FLDR_NAME_COMMCOMPS_PENDING           As String = "PendingReleases"
    Const FLDR_NAME_EXPORT                      As String = "source"
    Const FLDR_NAME_SERVICED_COMPMAN_SERVICES   As String = "CompMan"
    
    On Error GoTo eh
    Dim sLctn   As String
    Dim sName   As String
    
    '~~ 1. Common Component folder
    sLctn = wsConfig.FolderCompManRoot
    sName = FLDR_NAME_COMMCOMPS
    sCommCompsFolderPath = HistoryForwarded(sLctn, sName, True)
    
    '~~ 2. Common Components Peding folder
    sName = FLDR_NAME_COMMCOMPS_PENDING
    sLctn = sCommCompsFolderPath
    sCommCompsPendingFolderPath = HistoryForwarded(sLctn, sName, False)
    
    '~~ 3. Common Components Pending Release Private Profile file
    sLctn = sCommCompsFolderPath
    sName = FILE_NAME_COMMCOMPS_PENDING_PRIVPROF
    sCommCompsPendingPrivProfFileFullName = HistoryForwarded(sLctn, sName, False)

    '~~ 4. Serviced Workbook's CompMan service folder
    sLctn = mCompMan.ServicedWrkbk.Path
    sName = FLDR_NAME_SERVICED_COMPMAN_SERVICES
    sCompManServiceFolderPath = HistoryForwarded(sLctn, sName, True)

    '~~ 5. Serviced Workbook's Export folder
    sLctn = mCompMan.ServicedWrkbk.Path & ">" & sCompManServiceFolderPath
    sName = "source"
    sExportServiceFolderPath = HistoryForwarded(sLctn, sName, True)
    
    '~~ 6. Serviced Workbook's Execution Trace log file
    sLctn = mCompMan.ServicedWrkbk.Path & ">" & sCompManServiceFolderPath
    sName = FILE_NAME_EXEC_TRACE & "<CompMan.Service.trc"
    sServicedExecTraceLogFileFullName = HistoryForwarded(sLctn, sName, False, False)
    EstablishServicedExecTraceLogFile sServicedExecTraceLogFileFullName
    
    '~~ 7. Serviced Workbook's Common Components Private Profile file
    sLctn = mCompMan.ServicedWrkbk.Path & ">" & sCompManServiceFolderPath
    sName = FILE_NAME_COMMCOMPS_SERVICED_PRIVPROF & "<CompMan.dat"
    sCommCompsServicedPrivProfFileFullName = HistoryForwarded(sLctn, sName, False, False)
    EstablishCommCompsServicedPrivProfFile sCommCompsServicedPrivProfFileFullName
    
    '~~ 8. Serviced Workbook's Services log file
    sLctn = mCompMan.ServicedWrkbk.Path & ">" & sCompManServiceFolderPath
    sName = FILE_NAME_SERVICES_LOG & "<CompMan.Services.log"
    sServicesLogFileFullName = HistoryForwarded(sLctn, sName, False, False)

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RenameMove(ByVal e_path_old As String, _
                                 ByVal e_path_new As String, _
                        Optional ByVal e_folder As Boolean = True)
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

