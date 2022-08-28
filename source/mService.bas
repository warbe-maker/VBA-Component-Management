Attribute VB_Name = "mService"
Option Explicit

Public Serviced As Workbook     ' The  s e r v i c e d  Workbook throughout all services

' ---------------------------------------------------------------------------
' Interfaces to the wsService Worksheet
' ---------------------------------------------------------------------------
Private Property Get WsValue(Optional ByVal v_value_name As Variant) As Variant
    WsValue = mWbk.Value(v_ws:=wsService, v_name:=v_value_name)
End Property

Private Property Let WsValue(Optional ByVal v_value_name As Variant, _
                                      ByVal v_value As Variant)
    mWbk.Value(v_ws:=wsService, v_name:=v_value_name) = v_value
End Property

Public Sub AddAscByKey(ByRef add_dct As Dictionary, _
                        ByVal add_key As Variant, _
                        ByVal add_item As Variant)
' ------------------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by the key
' (add_key). When the key is an object with no Name property an error is raisede.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be copied
'       into any VBProject's module in order to have it independant from this
'       Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ------------------------------------------------------------------------------------
    Const PROC = "AddAscByKey"
    
    On Error GoTo eh
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItem           As Variant
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim vValueTarget    As Variant ' the add before/after add_key/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    bOrderByItem = False
    bOrderByKey = True
    bSeqAscending = True
    bCaseIgnored = False
    bCaseSensitive = True
    bStayWithFirst = True
    bEntrySequence = False
    
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not bStayWithFirst Then
            If .Exists(add_key) Then
                If VarType(add_item) = vbObject Then Set .item(add_key) = add_item Else .item(add_key) = add_item
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If VarType(add_key) = vbObject Then
            On Error Resume Next
            add_key.Name = add_key.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If VarType(add_item) = vbObject Then
            On Error Resume Next
            add_item.Name = add_item.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = AddAscByKeyValue(add_key)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = AddAscByKeyValue(.Keys()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If VarType(add_dct.item(vKeyExisting)) = vbObject _
        Then Set vItemExisting = add_dct.item(vKeyExisting) _
        Else vItemExisting = add_dct.item(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = AddAscByKeyValue(vKeyExisting)
            
                If vValueExisting = vValueNew And bOrderByItem And bSeqAscending And Not .Exists(add_key) Then
                    If bStayWithFirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and bStayWithFirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AddAscByKeyValue(ByVal add_key As Variant) As Variant
' ----------------------------------------------------------------------------
' When add_key is an object its name becomes the sort order value else the
' the value is returned as is.
' ----------------------------------------------------------------------------
    If VarType(add_key) = vbObject Then
        On Error Resume Next ' the object may not have a Name property
        AddAscByKeyValue = add_key.Name
        If Err.Number <> 0 Then Set AddAscByKeyValue = add_key
    Else
        AddAscByKeyValue = add_key
    End If
End Function

Public Function AllComps(ByVal ac_wbk As Workbook, _
                Optional ByVal ac_service As String = vbNullString) As Dictionary
' ---------------------------------------------------------------------------
' Returns a Dictionary with all VBComponents in ascending order thereby
' calculating the max lengths for vthe log entries.
' ---------------------------------------------------------------------------
    Dim vbc     As VBComponent
    Dim lDone   As Long
    
    If Log Is Nothing Then Set Log = New clsLog
    
    Set AllComps = New Dictionary
    For Each vbc In ac_wbk.VBProject.VBComponents
        Log.ServicedItem = vbc
        AddAscByKey AllComps, vbc.Name, vbc
        lDone = lDone + 1
        mService.DsplyStatus _
        mService.Progress(p_service:=ac_service _
                        , p_of:=lDone _
                        , p_dots:=lDone _
                         )
    Next vbc

End Function

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

Private Function CodeModuleIsEmpty(ByRef vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0
        If Not CodeModuleIsEmpty Then CodeModuleIsEmpty = .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Public Sub CompManAddinContinue()
' -------------------------------------------
' Continues the paused CompMan Addin Services
' -------------------------------------------
    If mMe.IsDevInstnc Then
        mConfig.AddinPaused = False
        mMe.DisplayStatus
    End If
End Sub

Public Sub CompManAddinPause()
' -------------------------------------------
' Continues the paused CompMan Addin Services
' -------------------------------------------
    If mMe.IsDevInstnc Then
        mConfig.AddinPaused = True
        mMe.DisplayStatus
    End If
End Sub

Public Function Denied() As Boolean
' --------------------------------------------------------------------------
' Returns TRUE when all preconditions for a service execution are fulfilled.
' --------------------------------------------------------------------------
    Const PROC = "Denied"
    
    On Error GoTo eh
    Dim sStatus As String
        
    If mMe.IsAddinInstnc And mMe.CompManAddinIsPaused Then
        '~~ When the service is about to be provided by the Addin but the Addin is currently paused
        '~~ another try with the serviced provided by the open Development instance may do the job.
        sStatus = "The CompMan Addin is currently paused. Open the development instance and retry."
    ElseIf WbkIsRestoredBySystem Then
        sStatus = "Service denied! Workbook appears restored by the system!"
    ElseIf mConfig.AddinPaused And mMe.IsAddinInstnc And Log.Service Like mCompManClient.SRVC_UPDATE_OUTDATED & "*" Then
        '~~ Note: The CompMan development instance is able to export its modified components but requires the
        '~~       Addin to upodate its outdated Used Common Components
        sStatus = "Service denied! The CompMan Addin is currently paused!"
    ElseIf FolderNotVbProjectExclusive Then
        sStatus = "Service denied! The Workbook is not the only one in its parent folder!"
    ElseIf Not mCompMan.WinMergeIsInstalled Then
        sStatus = "Service denied! WinMerge is required but not installed!"
    End If
    
xt: If sStatus <> vbNullString Then
        Log.Entry = sStatus
        mService.DsplyStatus Log.Service & sStatus
        Denied = True
    End If
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DsplyStatus(ByVal s As String)
    With Application
        .StatusBar = vbNullString
        .StatusBar = Trim(s) & vbLf
    End With
'    DoEvents
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
    If err_source = vbNullString Then err_source = Err.Source
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
            Else: ErrType = "VB Runtime Error "
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

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mService." & s
End Function

Public Sub EstablishServiceLog(ByVal sync_wbk_target As Workbook, _
                               ByVal sync_service As String)
     If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = sync_service
    End If
End Sub

Public Function ExpFilesDiffDisplay( _
          ByVal fd_exp_file_left_full_name As String, _
          ByVal fd_exp_file_right_full_name As String, _
          ByVal fd_exp_file_left_title As String, _
          ByVal fd_exp_file_right_title As String) As Boolean
' ----------------------------------------------------------------------------
' Displays the differences between export file 1 and 2 by means of WinMerge!
' Note: CompMan ignores any differences caused by empty code lines or case.
'       When a difference is displayed it is thus not because of this kind of
'       differneces but of others. Unfortunately it depends on the installed
'       WinMerge's set option wether or not these kind of differences are
'       displayed.
' ----------------------------------------------------------------------------
    Const PROC                          As String = "ExpFilesDiffDisplay"
    Const WINMERGE_SETTINGS_BASE_KEY    As String = "HKEY_CURRENT_USER\SOFTWARE\Thingamahoochie\WinMerge\Settings\"
    Const WINMERGE_BLANK_LINES          As String = "IgnoreBlankLines"
    Const WINMERGE_IGNORE_CASE          As String = "IgnoreCase"
    
    On Error GoTo eh
    Dim waitOnReturn        As Boolean: waitOnReturn = True
    Dim windowStyle         As Integer: windowStyle = 1
    Dim sCommand            As String
    Dim fso                 As New FileSystemObject
    Dim wshShell            As Object
    Dim sIgnoreBlankLines   As String ' 1 = True, 0 = False
    Dim sIgnoreCase         As String ' 1 = True, 0 = False
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="WinMerge is obligatory for the Compare service of this module but not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(fd_exp_file_left_full_name) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & fd_exp_file_left_full_name & """ does not exist!"
    
    If Not fso.FileExists(fd_exp_file_right_full_name) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & fd_exp_file_right_full_name & """ does not exist!"
    
    '~~ Save WinMerge configuration items and set them for CompMan
    sIgnoreBlankLines = mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_BLANK_LINES)
    sIgnoreCase = mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_IGNORE_CASE)
    mReg.Value(WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_BLANK_LINES) = "1"
    mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_IGNORE_CASE) = "1"
    
    '~~ Prepare command line
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & fd_exp_file_left_title & DQUOTE & _
               " /dr " & DQUOTE & fd_exp_file_right_title & DQUOTE & " " & _
               """" & fd_exp_file_left_full_name & """" & " " & _
               """" & fd_exp_file_right_full_name & """" ' & sIniFile doesn't work

    
    '~~ Execute command line
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        ExpFilesDiffDisplay = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
    '~~ Restore WinMerge configuration items
    mReg.Value(WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_BLANK_LINES) = sIgnoreBlankLines
    mReg.Value(reg_key:=WINMERGE_SETTINGS_BASE_KEY, reg_value_name:=WINMERGE_IGNORE_CASE) = sIgnoreCase

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ExportChangedComponents(ByVal hosted As String)
' -----------------------------------------------------------
' Exclusively called by mCompMan.ExportChangedComponents,
' triggered by the Before_Save event.
' - Any VBComponent the code has changed (temporary
'   Export-File differs from the current Export-File or no
'   Export-File exists) exported
' - Outdated Export-Files (components no longer existing) are
'   removed
' - Modified 'Clone-Components' require a confirmation by the
'   user.
'
' Note: For the case the Export-Folder may have changed any
'       Export-File found within the Workbook-Folder outside
'       the specified Export-Folder is removed.
' Attention: This procedure is called exclusively by
'            mCompMan.UpdateOutdatedCommonComponents! When called directly
'            by the user, e.g. via the 'Imediate Window' an
'            error will be raised because an 'mService.Serviced'
'            Workbook is not set.
' --------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If mService.Serviced Is Nothing _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The procedure '" & ErrSrc(PROC) & "' has been called without a prior set of the 'Serviced' Workbook. " & _
                                                 "(it may have been called directly via the 'Immediate Window'"
    If mService.Denied Then GoTo xt
    mCompMan.MaintainPropertiesOfHostedRawCommonComponents hosted
    Set Log = New clsLog
    Log.Service = PROC
    mExport.ChangedComponents
        
xt: Set dctHostedRaws = Nothing
    Set Log = Nothing
    mBasic.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function FilesDiffer(ByVal fd_exp_file_1 As File, _
                            ByVal fd_exp_file_2 As File) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when file 1 and file 2 are different whereby case differences
' and empty lines are ignored. This function guarantees a uniform comparison of
' export files throughout CompMan.
' ----------------------------------------------------------------------------
    Const PROC = "FilesDiffer"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim fl1 As File
    Dim fl2 As File
    
    With fso
        If TypeName(fd_exp_file_1) = "File" Then
            If Not .FileExists(fd_exp_file_1) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided 'fd_exp_file_1' does not exist!"
            Set fl1 = fso.GetFile(fd_exp_file_1)
        ElseIf TypeName(fd_exp_file_1) = "Nothing" _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "File 'fd_exp_file_1' is not provided!"
        Else
            Set fl1 = fd_exp_file_1
        End If
        
        If TypeName(fd_exp_file_2) = "File" Then
            If Not .FileExists(fd_exp_file_2) _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided 'fd_exp_file_2' does not exist!"
            Set fl2 = fso.GetFile(fd_exp_file_2)
        ElseIf TypeName(fd_exp_file_2) = "Nothing" _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "File 'fd_exp_file_2' is not provided!"
        Else
            Set fl2 = fd_exp_file_2
        End If
    End With
    
    FilesDiffer = mFile.Differs(fd_file1:=fl1 _
                              , fd_file2:=fl2 _
                              , fd_stop_after:=1 _
                              , fd_ignore_empty_records:=True _
                              , fd_compare:=vbTextCompare).Count <> 0
xt: Set fso = Nothing
    Exit Function
                            
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FilesDifference(ByVal fd_exp_file_1 As File, _
                                ByVal fd_exp_file_2 As File) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with either 0 items when file 1 and file 2 are
' identical or with one item when the two files differ. Empty lines and case
' differences are ignored because the do not constitute a relevant code change
' ----------------------------------------------------------------------------
    Const PROC = "FilesDifference"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim fl1 As File
    Dim fl2 As File
    
    With fso
        If VarType(fd_exp_file_1) = vbString Then
            If Not .FileExists(fd_exp_file_1) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided 'fd_exp_file_1' is a string not identifying an existing file!"
            Set fl1 = fso.GetFile(fd_exp_file_1)
        Else
            Set fl1 = fd_exp_file_1
        End If
        
        If VarType(fd_exp_file_2) = vbString Then
            If Not .FileExists(fd_exp_file_2) _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided 'fd_exp_file_2' is a string not identifying an existing file!"
            Set fl2 = fso.GetFile(fd_exp_file_2)
        Else
            Set fl2 = fd_exp_file_2
        End If
    End With
    
    Set FilesDifference = mFile.Differs(fd_file1:=fl1 _
                                      , fd_file2:=fl2 _
                                      , fd_stop_after:=1 _
                                      , fd_ignore_empty_records:=True _
                                      , fd_compare:=vbTextCompare)
                            
xt: Set fso = Nothing
    Exit Function
                            
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FolderNotVbProjectExclusive() As Boolean

    Dim fso As New FileSystemObject
    Dim fl  As File
    
    For Each fl In fso.GetFolder(mService.Serviced.Path).Files
        If VBA.Left$(fso.GetFileName(fl.Path), 2) = "~$" Then GoTo next_fl
        If VBA.StrComp(fl.Path, mService.Serviced.FullName, vbTextCompare) <> 0 Then
            Select Case fso.GetExtensionName(fl.Path)
                Case "xlsm", "xlam", "xlsb" ' may contain macros, a VB-Project repsectively
                    FolderNotVbProjectExclusive = True
                    Exit For
            End Select
        End If
next_fl:
    Next fl

End Function

Public Sub Install(Optional ByRef in_wbk As Workbook = Nothing)
    Const PROC = "Install"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    If in_wbk Is Nothing Then Set in_wbk = SelectServicedWrkbk(PROC)
    If in_wbk Is Nothing Then GoTo xt
    mInstall.CommonComponents in_wbk

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

Public Function Progress(ByVal p_service As String, _
                Optional ByVal p_result As Long = 0, _
                Optional ByVal p_of As Long = 0, _
                Optional ByVal p_op As String = vbNullString, _
                Optional ByVal p_comps As String = vbNullString, _
                Optional ByVal p_dots As Long = 0) As String
' --------------------------------------------------------------------------
' Displays a services progress in the Application.StatusBar in the form:
' service for serviced: n of m op [(component [, component] ..]
' --------------------------------------------------------------------------
    Const MSG_SCHEEME = "<s><n> of <m> <op> <comps> <dots>"
    
    Dim sMsg        As String
    
    sMsg = Replace(MSG_SCHEEME, "<s>", p_service)
    sMsg = Replace(sMsg, "<n>", p_result)
    sMsg = Replace(sMsg, "<m>", p_of)
    sMsg = Replace(sMsg, "<op>", p_op)
    If p_comps <> vbNullString Then
        If Right(p_comps, 2) = ", " Then p_comps = Left(p_comps, Len(p_comps) - 2)
        If Right(p_comps, 1) = "," Then p_comps = Left(p_comps, Len(p_comps) - 1)
        sMsg = Replace(sMsg, "<comps>", "(" & p_comps & ")")
    Else
        sMsg = Replace(sMsg, "<comps>", vbNullString)
    End If
    If p_dots > 0 Then
        sMsg = Replace(sMsg, "<dots>", VBA.String(p_dots, "."))
    Else
        sMsg = Replace(sMsg, "<dots>", vbNullString)
    End If
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 250) & " ..."
        
    Progress = sMsg
    
End Function

Public Sub RemoveTempRenamed()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "RemoveTempRenamed"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim vbc As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    With mService.Serviced.VBProject
        For Each vbc In .VBComponents
            If mService.IsRenamedByCompMan(vbc.Name) Then
                .VBComponents.Remove vbc
            End If
        Next vbc
    End With

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SaveWbk(ByRef rs_wbk As Workbook)
    Const PROC = "SaveWbk"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    Application.EnableEvents = False
    '~~ This is the action where the update process may lead to the effect that Excel closes the Workbook
    '~~ without having deleted the renamed components!
    mBasic.TimedDoEvents (ErrSrc(PROC))
    rs_wbk.Save
    mBasic.TimedDoEvents (ErrSrc(PROC))
    Application.EnableEvents = True

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function SelectServicedWrkbk(ByVal gs_service As String) As Workbook
    Dim fl As File
    
    If mFile.Picked(p_title:="Select the Workbook (may already be open, will be opened if not) to be served by the " & gs_service & " service" _
                  , p_filters:="Excel Workbook,*.xl*" _
                  , p_file:=fl) _
    Then Set SelectServicedWrkbk = mCompMan.WbkGetOpen(fl.Path) _
    Else: Set SelectServicedWrkbk = Nothing

End Function

Private Function SyncSourceAndTargetSelected( _
                     Optional ByRef sync_wbk_target As Workbook = Nothing, _
                     Optional ByVal sync_source_wbk_name As String = vbNullString, _
                     Optional ByRef sync_source_wbk As Workbook = Nothing, _
                     Optional ByVal sync_confirm_info As Boolean = False) As Boolean
' ---------------------------------------------------------------------------
' Returns True when the Sync-Target_VB-Project and the Sync-Source-VB-Project
' are valid. When bc_sync_confirm_info is True a confirmation dialog is
' displayed. The dialog is also displayed when function is called with
' invalid arguments.
' ---------------------------------------------------------------------------
    Const PROC                  As String = "SyncSourceAndTargetSelected"
    Const TARGET_PROJECT        As String = "Target-Workbook/VBProject"
    Const SOURCE_PROJECT        As String = "Source-Workbook/VBProject"
    
    On Error GoTo eh
    Dim bWbSource               As Boolean
    Dim bWbTarget               As Boolean
    Dim Buttons                 As Collection
    Dim fl                      As File
    Dim fso                     As New FileSystemObject
    Dim sBttnSourceTargetCnfrmd As String
    Dim sBttnSourceProject      As String
    Dim sBttnTargetProject      As String
    Dim sBttnTerminate          As String
    Dim sMsg                    As TypeMsg
    Dim sReply                  As String
    Dim sWbSource               As String ' either a full name or a registered raw project's basename
    Dim sWbTarget               As String
    
    sBttnSourceTargetCnfrmd = "Selected Source- and" & vbLf & _
                            "Target-Workbook/VBProject" & vbLf & _
                            "Confirmed"
    sBttnTargetProject = "Select/change the" & vbLf & vbLf & TARGET_PROJECT & vbLf & " "
    sBttnSourceProject = "Configure/change the" & vbLf & vbLf & SOURCE_PROJECT & vbLf & " "
    sBttnTerminate = "Terminate providing a " & vbLf & _
                     "Source- and Target-Workbook/VBProject"
    
    If Not sync_wbk_target Is Nothing Then sWbTarget = sync_wbk_target.FullName
    sWbSource = sync_source_wbk_name
    
    While (Not bWbTarget Or Not bWbSource) Or (sync_confirm_info And sReply <> sBttnSourceTargetCnfrmd)
        If sWbTarget = vbNullString Then
            sWbTarget = "n o t  p r o v i d e d !"
        ElseIf Not fso.FileExists(sWbTarget) Then
            sWbSource = sWbSource & ": i n v a l i d ! (does not exist)"
        Else
            sWbTarget = Split(sWbTarget, ": ")(0)
            bWbTarget = True
        End If
        
        If sWbSource = vbNullString Then
            sWbSource = "n o t  p r o v i d e d !"
        ElseIf Not fso.FileExists(sWbSource) Then
            sWbSource = sWbSource & ": i n v a l i d ! (does not exist)"
        Else
            sWbSource = Split(sWbSource, ": ")(0)
            bWbSource = True
        End If
    
        If bWbSource And bWbTarget And Not sync_confirm_info Then GoTo xt
        
        With sMsg
            .Section(1).Label.Text = TARGET_PROJECT & ":"
            .Section(1).Text.Text = sWbTarget
            .Section(1).Text.MonoSpaced = True
            .Section(2).Label.Text = SOURCE_PROJECT & ":"
            .Section(2).Text.Text = sWbSource
            .Section(2).Text.MonoSpaced = True
            
            If sync_confirm_info _
            Then .Section(3).Text.Text = "Please confirm the above current 'Basic CompMan Configuration'." _
            Else: .Section(3).Text.Text = "Please provide/complete the 'Basic CompMan Configuration'."
            
            .Section(3).Text.Text = .Section(3).Text.Text & vbLf & vbLf & _
                                "Attention!" & vbLf & _
                                "1. The '" & TARGET_PROJECT & "' must not be identical with the '" & SOURCE_PROJECT & "' and the two Workbooks must not have the same name." & vbLf & _
                                "2. Both VB-Projects/Workbook must exclusively reside in their parent Workbook" & vbLf & _
                                "3. Both Workbook folders must be subfolders of the configured '" & FOLDER_SERVICED & "'."

        End With
        
        '~~ Buttons preparation
        If Not bWbTarget Or Not bWbSource _
        Then Set Buttons = mMsg.Buttons(sBttnSourceProject, sBttnTargetProject, vbLf, sBttnTerminate) _
        Else: Set Buttons = mMsg.Buttons(sBttnSourceTargetCnfrmd, vbLf, sBttnSourceProject, sBttnTargetProject)
        
        If Not mMsg.IsValidMsgButtonsArg(Buttons) Then Stop
        sReply = mMsg.Dsply(dsply_title:="Basic configuration of the Component Management (CompMan Addin)" _
                          , dsply_msg:=sMsg _
                          , dsply_buttons:=Buttons _
                           )
        Select Case sReply
            Case sBttnTargetProject
                Do
                    If mFile.Picked(p_title:="Select the '" & TARGET_PROJECT & " to be synchronized with the '" & SOURCE_PROJECT & "'" _
                                  , p_filters:="Excel Workbook,*.xl*" _
                                  , p_file:=fl) _
                    Then
                        sWbTarget = fl.Path
                        Exit Do
                    End If
                Loop
                sync_confirm_info = True
                '~~ The change of the VB-Clone-Project may have made the VB-Raw-Project valid when formerly invalid
                sWbSource = Split(sWbSource, ": ")(0)
            Case sBttnSourceProject
                Do
                    If mFile.Picked(p_title:="Select the '" & SOURCE_PROJECT & " as the synchronization source for the '" & TARGET_PROJECT & "'" _
                                  , p_filters:="Excel Workbook,*.xl*" _
                                  , p_file:=fl) _
                    Then
                        sWbSource = fl.Path
                        Exit Do
                    End If
                Loop
                sync_confirm_info = True
                '~~ The change of the VB-Raw-Project may have become valid when formerly invalid
                sWbTarget = Split(sWbTarget, ": ")(0)
            
            Case sBttnSourceTargetCnfrmd: sync_confirm_info = False
            Case sBttnTerminate: GoTo xt
                
        End Select
        
    Wend ' Loop until the confirmed or configured basic configuration is correct
    
xt: If bWbTarget Then
       Set sync_wbk_target = mCompMan.WbkGetOpen(sWbTarget)
    End If
    If bWbSource Then
        Application.EnableEvents = False
        Set sync_source_wbk = mCompMan.WbkGetOpen(sWbSource)
        Application.EnableEvents = True
        sync_source_wbk_name = fso.GetBaseName(sWbSource)
    End If
    SyncSourceAndTargetSelected = bWbTarget And bWbSource
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function UsedCommonComponents(ByRef cl_wbk As Workbook) As Dictionary
' ---------------------------------------------------------------------------
' Returns a Dictionary of all Used Common Components with its VBComponent
' object as key and its name as item.
' ---------------------------------------------------------------------------
    Const PROC = "UsedCommonComponents"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim dct     As New Dictionary
    Dim fso     As New FileSystemObject
    Dim Comp    As clsComp
    Dim RawComp As clsRaw
    
    mBasic.BoP ErrSrc(PROC)
        
    For Each vbc In cl_wbk.VBProject.VBComponents
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = cl_wbk
            .CompName = vbc.Name
            Log.ServicedItem = .VBComp
            If .KindOfComp = enCommCompUsed Then
                If .Changed Then
                    dct.Add vbc, vbc.Name
                Else
                    Log.Entry = "Code un-changed."
                End If
            End If
        End With
        Set Comp = Nothing
        Set RawComp = Nothing
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)
    Set UsedCommonComponents = dct
    Set fso = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsRestoredBySystem() As Boolean
    WbkIsRestoredBySystem = InStr(ActiveWindow.Caption, "(") <> 0 _
                         Or InStr(mService.Serviced.FullName, "(") <> 0
End Function

