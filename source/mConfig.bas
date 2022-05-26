Attribute VB_Name = "mConfig"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mConfig
' Read/Write CompMan configuration properties from /to the Registry
' ----------------------------------------------------------------------------
Private Const CONFIG_BASE_KEY               As String = "HKCU\SOFTWARE\CompManVBP\BasicConfig\"
Private Const VNAME_ADDIN_IS_PAUSED         As String = "AddinIsPaused"
Private Const VNAME_FOLDER_ADDIN            As String = "FolderAddin"
Private Const VNAME_FOLDER_EXPORT           As String = "FolderExport"
Private Const VNAME_FOLDER_SERVICED         As String = "FolderServiced"
Private Const VNAME_FOLDER_SYNCED           As String = "FolderSynced"

Private FolderAddinIsValid                  As Boolean
Private FolderExportIsValid                 As Boolean
Private FolderServicedIsValid               As Boolean
Private FolderSyncedIsValid                 As Boolean

Public Property Get AddinPaused() As Boolean
    AddinPaused = CBool(RegValue(VNAME_ADDIN_IS_PAUSED))
End Property

Public Property Let AddinPaused(ByVal b As Boolean):   RegValue(VNAME_ADDIN_IS_PAUSED) = Abs(CInt(b)): End Property

Public Property Get FolderAddin() As String:                    FolderAddin = WsValue(VNAME_FOLDER_ADDIN):      End Property

Public Property Let FolderAddin(ByVal s As String):             WsValue(VNAME_FOLDER_ADDIN) = s:                End Property

Public Property Get FolderExport() As String
' ----------------------------------------------------------------------------
' When no Export Folder name is available 'source' is returned as the default.
' Any changes of the default name appear as names ammended comma separeated.
' The prefix \ indicates that the names are names of a sub-folder of the
' serviced Workbook's parent folder. When there is a change history any still
' existing old folder is renamed to the last (most current) specified name.
' ----------------------------------------------------------------------------
    Dim sHistory    As String
    Dim Names()     As String
    Dim iUBound     As Long
    
    sHistory = WsValue(VNAME_FOLDER_EXPORT)
    If sHistory = vbNullString Then
        FolderExport = "source" ' The default
    Else
        Names = Split(sHistory, ",")
        iUBound = UBound(Names)
        FolderExport = Names(iUBound)
        
        If Not mService.Serviced Is Nothing Then
            If iUBound > 0 Then
                '~~ When the Export folder name is obtained for a certain serviced Workbook
                '~~ and not just for the configuration of its name - and there is a history
                '~~ of names - any outdated old Export Folder name is renamed along with
                '~~ this provision of the most current name
                ForwardFolderName sHistory, mService.Serviced.Path
            End If
        End If
    End If
    
End Property

Public Property Let FolderExport(ByVal s As String)
' ----------------------------------------------------------------------------
' When the most current folder name is not = s the new name is added comma
' separated in order to have the history of names maintained.
' Note: It may take an unknown time until all Workbooks serviced by CompMan
'       had been switched to a new Export Folder name. The history allows
'       forwarding outdated folder names to the most current configured name.
' ----------------------------------------------------------------------------
    Dim sHistory    As String
    Dim Names()     As String
    Dim iUBound     As Long
    
    sHistory = WsValue(VNAME_FOLDER_EXPORT)
    If sHistory = vbNullString Then
        WsValue(VNAME_FOLDER_EXPORT) = s
    Else
        Names = Split(sHistory, ",")
        iUBound = UBound(Names)
        If Names(iUBound) <> s Then
            WsValue(VNAME_FOLDER_EXPORT) = sHistory & "," & s
        End If
    End If
End Property

Public Property Get FolderServiced() As String:             FolderServiced = WsValue(VNAME_FOLDER_SERVICED):        End Property

Public Property Let FolderServiced(ByVal s As String):      WsValue(VNAME_FOLDER_SERVICED) = s:                     End Property

Public Property Get FolderSynced() As String:               FolderSynced = WsValue(VNAME_FOLDER_SYNCED):            End Property

Public Property Let FolderSynced(ByVal s As String):        WsValue(VNAME_FOLDER_SYNCED) = s:                       End Property

' ---------------------------------------------------------------------------
' Interfaces to the wsBasicConfig Worksheet
' ---------------------------------------------------------------------------
Private Property Get WsValue(Optional ByVal v_value_name As String) As Variant
    WsValue = mWbk.Value(v_ws:=wsBasicConfig, v_name:=v_value_name)
End Property

Private Property Let WsValue(Optional ByVal v_value_name As String, _
                                    ByVal v_value As Variant)
    mWbk.Value(v_ws:=wsBasicConfig, v_name:=v_value_name) = v_value
End Property

' ---------------------------------------------------------------------------
' Interfaces to the Registry
' ---------------------------------------------------------------------------
Private Property Get RegValue(Optional ByVal v_value_name As String) As Variant
    RegValue = mReg.Value(reg_key:=CONFIG_BASE_KEY, reg_value_name:=v_value_name)
End Property

Private Property Let RegValue(Optional ByVal v_value_name As String, _
                                    ByVal v_value As Variant)
    mReg.Value(reg_key:=CONFIG_BASE_KEY, reg_value_name:=v_value_name) = v_value
End Property

Public Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mConfig." & sProc
End Function

Private Function WsExists(ByVal value_name As String) As Boolean
    WsExists = mWbk.Exists(ex_wb:=ThisWorkbook, ex_range_name:=value_name)
End Function

Public Sub ForwardFolderName(ByRef ff_history As String, _
                             ByVal ff_wb_parent_folder As String)
' ----------------------------------------------------------------------------
' Forwards (renames) any outdatede Export Folder name in a Workbook's parent
' folder (ff-wb_parent_folder) to the current name.
' ----------------------------------------------------------------------------
    Const PROC = "ForwardFolderName"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim Names() As String
    Dim iUBound As Long
    Dim i       As Long
    Dim NameNow As String
    Dim sFolder As String
    
    Names = Split(ff_history, ",")
    iUBound = UBound(Names)
    If iUBound = 0 Then GoTo xt
    NameNow = Trim(Names(iUBound))
    
    With fso
        For i = iUBound - 1 To 0 Step -1
            sFolder = ff_wb_parent_folder & "\" & Names(i)
            sFolder = Replace(sFolder, "\\", "\")
            If .FolderExists(sFolder) Then
                .GetFolder(sFolder).Name = NameNow
            End If
        Next i
    End With
    
xt: Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

