Attribute VB_Name = "mMe"
Option Explicit
Option Private Module
' ---------------------------------------------------------------------------
' Standard Module mMe       Services for the self management like the relation
'                           between the Component Management AddIn instance
'                           and the development instance.
'
' Public services:
' - AddInInstncWrkbkIsOpen  Returns True when the AddIn instance Workbook is
'                           open
' - CompManAddinPath        Get/Let the configured path for the AddIn instance
'                           of this Workbook
' - ConfigAsserted          Returns True when the required properties (paths)
'                           are configured and exist
' - ConfirmConfig           Get the current configured paths confirmed before
'                           the AddIn instance of this Workbook is established
'                           or renewed respectively
' - ControlItemRenewAdd     Used when the development and test instance
'                           Workbook is opened to add a RenewAddIn control
'                           item to the "Add-Ins" poupup menu.
' - ControlItemRenewRemove  Used when the development and test instance
'                           Workbook is closed.
' - RenewAddIn              Called via the RenewAddIn control item in the
'                           "Add-Ins" popup menu or executed via the
'                           VBE when the code of the Addin had been modified
'                           in its Development instance Workbook.
'                           + -------------------------------------------------- +
' - SaveAsDev               || Saves the Addin Workbook as Development instance ||
'                           || Workbook. Used exclusively through the immediate ||
'                           || Window when the code in the Addin has been       ||
'                           || modified directly - instead of through the       ||
'                           || Development instance Workbook                    ||
'                           + -------------------------------------------------- +
' - UpdateRawClones     Exclusively used by the Development instance
'                           Workbook to update its own used Common Components
'                           of which the origin code had changed - via the
'                           Addin - provided it is open - by means of the
'                           the corresponding method.
' - RootServicedByCompMan       Configured path for any Workbook serviced by the
'                           AddIn instance of this Workbook
' Uses Common Components:
' - mFile                   Get/Let PrivateProperty value service
' - mWrkbk                  GetOpen and Opened service
' - mMsg                    Dsply, Box, and Buttons service used by the
'                           RenewAddin,  ConfirmConfig service
'
' Requires:
' W. Rauschenberger, Berlin Nov 2020
' ---------------------------------------------------------------------------
Private Const CONTROL_CAPTION_RENEW As String = "Renew Addin"
Private Const CONTROL_CAPTION_PAUSE As String = "Pause Addin"
Private Const SECTION_BASE_CONFIG   As String = "BaseConfiguration"
Private Const ROOT_COMPMAN_SERVICED  As String = "VBDevProjectsRoot"
Private Const COMPMAN_ADDIN_LCTN    As String = "CompManAddInPath"

Private Const ADDIN_WORKBOOK        As String = "CompMan.xlam"      ' Extension adjusted when the above is saved as addin
Private Const ADDIN_VERSION         As String = "5.0"               ' Allows to check the success of an Addin renew
Private Const DEVLP_WORKBOOK        As String = "CompManDev.xlsb"   ' Extension depends on Excel versions
Private cllLog                      As Collection                   ' log records of the RenewAddIn and/or SaveAsDev process
Private wbDevlp                     As Workbook
Private wbSource                    As Workbook                     ' This development instance as the renew source
Private wbTarget                    As Workbook                     ' The Addin instance as renew target
Private bSucceeded                  As Boolean
Private bAllRemoved                 As Boolean
Private dctAddInRefs                As Dictionary
Private lStep                       As Long
Private bAddInPaused                As Boolean

Private Property Get DEVLP_FORMAT() As XlFileFormat  ' = .xlsb ! may require adjustment when the above is changed
    DEVLP_FORMAT = xlExcel12
End Property

Private Property Get ADDIN_FORMAT() As XlFileFormat ' = ... needs adjustment when the above is changed
    ADDIN_FORMAT = xlOpenXMLAddIn
End Property

Public Property Get ServicesLogFile(Optional ByVal slf_serviced_wb As Workbook) As String
    Dim fso As New FileSystemObject
    ServicesLogFile = slf_serviced_wb.Path & "\" & fso.GetBaseName(mMe.AddInInstanceName) & ".Services.log"
    Set fso = Nothing
End Property

Public Property Get PendingServicesFile(Optional ByVal slf_serviced_wb As Workbook) As String
    Dim fso As New FileSystemObject
    PendingServicesFile = slf_serviced_wb.Path & "\" & fso.GetBaseName(mMe.AddInInstanceName) & ".Services.Pending.imp"
    Set fso = Nothing
End Property

Public Property Get AddInInstanceFullName() As String:  AddInInstanceFullName = AddInPath & DBSLASH & AddInInstanceName:    End Property

Public Property Get AddInInstanceName() As String:      AddInInstanceName = ADDIN_WORKBOOK:                                 End Property

Private Property Get AddInPath() As String:             AddInPath = CompManAddinPath:                                  End Property

Public Property Get DevInstncFullName() As String
Dim fso As New FileSystemObject
    DevInstncFullName = RootServicedByCompMan & DBSLASH _
                          & fso.GetBaseName(DevInstncName) & DBSLASH _
                          & DevInstncName
End Property

Public Property Get DevInstncName() As String:          DevInstncName = DEVLP_WORKBOOK:                                 End Property

Public Property Get IsAddinInstnc() As Boolean:         IsAddinInstnc = ThisWorkbook.name = AddInInstanceName:          End Property

Public Property Get IsDevInstnc() As Boolean:           IsDevInstnc = ThisWorkbook.name = DevInstncName:                End Property

Public Property Get xlAddInFormat() As Long:            xlAddInFormat = ADDIN_FORMAT:                                       End Property

Public Property Get xlDevlpFormat() As Long:            xlDevlpFormat = DEVLP_FORMAT:                                       End Property

Public Function AddInVersion(Optional ByRef sVersion As String) As String
    sVersion = ADDIN_VERSION
    AddInVersion = sVersion
End Function

Private Property Get CFG_CHANGE_ADDIN_LCTN() As String
    CFG_CHANGE_ADDIN_LCTN = "Change CompMan" & vbLf & "AddIn Path"
End Property

Private Property Get CFG_CHANGE_DEVELOPMENT_ROOT() As String
    CFG_CHANGE_DEVELOPMENT_ROOT = "Change development" & vbLf & "root folder"
End Property

Private Property Get CFG_FILENAME() As String: CFG_FILENAME = ThisWorkbook.Path & "\CompMan.cfg": End Property

Public Property Get CompManAddinPath() As String
    CompManAddinPath = Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMPMAN_ADDIN_LCTN)
End Property

Public Property Let CompManAddinPath(ByVal s As String)
    Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMPMAN_ADDIN_LCTN) = s
End Property

Private Property Get Step() As Long
    If lStep = 0 Then lStep = 1 Else lStep = lStep + 1
    Step = lStep
End Property

Private Property Get Value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
    
    Value = mFile.Value(vl_file:=CFG_FILENAME _
                      , vl_section:=vl_section _
                      , vl_value_name:=vl_value_name _
                       )
End Property

Private Property Let Value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
    mFile.Value(vl_file:=CFG_FILENAME _
              , vl_section:=vl_section _
              , vl_value_name:=vl_value_name _
               ) = vl_value
End Property

Public Property Get RootServicedByCompMan() As String
    RootServicedByCompMan = Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=ROOT_COMPMAN_SERVICED)
End Property

Public Property Let RootServicedByCompMan(ByVal s As String)
    Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=ROOT_COMPMAN_SERVICED) = s
End Property

Private Sub CloseAddinInstncWorkbook()
    Const PROC = "CloseAddinInstncWorkbook"
    
    Dim lStep As Long

    On Error Resume Next
    wbSource.Activate
    wbTarget.Close False
    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Close the Addin instance Workbook"

    If Err.Number <> 0 _
    Then cllLog.Add "RenewAddIn " & lStep & ". Failed! Closing the Addin instance Workbook '" & AddInInstanceName & "' failed with:" & vbLf & _
                 "  (" & Err.Description & ")" _
    Else cllLog.Add "RenewAddIn " & lStep & ". Passed! Addin Workbook instance successfully closed"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub DeleteAddInInstanceWorkbook()
    Const PROC = "DeleteAddInInstanceWorkbook"
    
    On Error GoTo eh
    Dim lStep   As Long
    Dim fso     As New FileSystemObject

    On Error Resume Next
    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Delete the Addin instance Workbook"

    With fso
        If .FileExists(AddInInstanceFullName) Then
            .DeleteFile AddInInstanceFullName
            If Err.Number = 0 _
            Then cllLog.Add "RenewAddIn " & lStep & ". Passed! The Addin instance Workbook '" & AddInInstanceName & "' had successfully been removed." _
            Else cllLog.Add "RenewAddIn " & lStep & ". Failed! Deleting the Addin instance Workbook '" & AddInInstanceName & "' failed with:" & vbLf & _
                         "  (" & Err.Description & ")"
        Else
            cllLog.Add "RenewAddIn " & lStep & ". Passed! The Addin instance Workbook '" & AddInInstanceName & "' was found already deleted"
        End If
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function AddInInstncWrkbkExists() As Boolean
    Dim fso As New FileSystemObject
    AddInInstncWrkbkExists = fso.FileExists(AddInInstanceFullName)
End Function

Public Function AddInInstncWrkbkIsOpen() As Boolean
    Const PROC = "AddInInstncWrkbkIsOpen"
    
    On Error GoTo eh
    Dim i As Long
    
    AddInInstncWrkbkIsOpen = False
    For i = 1 To Application.AddIns2.Count
        If Application.AddIns2(i).name = AddInInstanceName Then
            On Error Resume Next
            Set wbTarget = Application.Workbooks(AddInInstanceName)
            AddInInstncWrkbkIsOpen = Err.Number = 0
            GoTo xt
        End If
    Next i
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Sub OpenAddinInstncWorkbook()
    Const PROC = "OpenAddinInstncWorkbook"
    
    On Error GoTo eh
    Dim lStep           As Long
    Dim wb              As Workbook
    Dim sBaseAddinName  As String
    Dim sBaseDevName    As String
    Dim fso             As New FileSystemObject

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Open the Addin instance Workbook"
    
    If Not AddInInstncWrkbkIsOpen Then
        If AddInInstncWrkbkExists Then
            On Error Resume Next
            Set wb = Application.Workbooks.Open(AddInInstanceFullName)
            If Err.Number = 0 Then
                With fso
                    sBaseAddinName = .GetBaseName(wb.name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.name)
                    wb.VBProject.name = sBaseAddinName
                End With
                cllLog.Add "RenewAddIn " & lStep & ". Passed! The Addin Workbook instance '" & AddInInstanceName & "' has successfully been (re)opened."
            Else
                cllLog.Add "RenewAddIn " & lStep & ". Failed! (Re)opening the Addin Workbook '" & AddInInstanceName & "' failed with:" & _
                         vbLf & "  (" & Err.Description & ")"
            End If
        End If
    End If

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub SaveAddinInstncWorkbookAsDevlp()
    Const PROC = "SaveAddinInstncWorkbookAsDevlp"

    On Error GoTo eh
    Dim lStep   As Long
    Dim fso     As New FileSystemObject

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Save the Addin instance Workbook as Development instance"
    
    With Application
        If Not DevInstncWorkbookExists Then
            '~~ At this point the Development instance Workbook must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbAddIn.SaveAs DevInstncFullName, FileFormat:=xlDevlpFormat, ReadOnlyRecommended:=False
            
            If Not mCompMan.WbkIsOpen(io_name:=DevInstncName) _
            Then Stop _
            Else wbDevlp.VBProject.name = fso.GetBaseName(DevInstncName)
            
            If Err.Number <> 0 Then
                cllLog.Add "RenewAddIn " & lStep & ". Failed! Saving the Addin instance Workbook '" & AddInInstanceName & "' (version " & AddInVersion & " as Development instance Workbook '" & DevInstncName & "' failed!"
            Else
                cllLog.Add "RenewAddIn " & lStep & ". Passed! The Addin instance Workbook '" & AddInInstanceName & "' (version " & AddInVersion & " has successfully been saved as Development instance Workbook '" & DevInstncName & "'"
            End If
            .EnableEvents = True
        Else ' file still exists
            cllLog.Add "RenewAddIn " & lStep & ". Failed! Saving the Addin instance Workbook '" & AddInInstanceName & "' (version " & AddInVersion & " as Development instance Workbook '" & DevInstncName & "' failed!"
        End If
    End With

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function ConfigAsserted() As Boolean
' ----------------------------------------------------
' Assert that an existing Common folder is configured
' and that it contains a subfolder "CommComponents".
' Attention! This function must not run in the AddIn
' instance of this Workbook!
' ----------------------------------------------------
    Const PROC = "Assert"
    
    On Error GoTo eh
    Dim sPathCompMan    As String
    Dim fso             As New FileSystemObject
    
    With fso
        If .FolderExists(CompManAddinPath) _
        And .FolderExists(RootServicedByCompMan) _
        Then
            ConfigAsserted = True
            .CopyFile Source:=CFG_FILENAME, Destination:=CompManAddinPath & "\CompMan.cfg", OverWriteFiles:=True
            GoTo xt
        End If
                
        '~~ Assert the folder for the CompMan AddIn
        sPathCompMan = CompManAddinPath
        If sPathCompMan = vbNullString Then
            sPathCompMan = mBasic.SelectFolder( _
                           sTitle:="Select the folder for the AddIn instance of the CompManDev Workbook (escape to use the Application.UserLibraryPath)")
            If sPathCompMan = vbNullString Then
                sPathCompMan = Application.UserLibraryPath ' Default because user escaped the selection
            Else
                '~~ Assure trust in this location and save it to the CompMan.cfg file
                TrustThisFolder FolderPath:=sPathCompMan
                CompManAddinPath = sPathCompMan
            End If
        Else
            While Not .FolderExists(sPathCompMan)
                '~~ Configured but folder not or no longer exists
                sPathCompMan = mBasic.SelectFolder( _
                               sTitle:="The configured CompMan AddIn folder does not exist. Select another one or escape for the default '" & Application.UserLibraryPath & "' path")
                If sPathCompMan = vbNullString Then
                    sPathCompMan = Application.UserLibraryPath
                Else
                    '~~ Assure trust in this location and save it to the CompMan.cfg file
                    TrustThisFolder FolderPath:=sPathCompMan
                    CompManAddinPath = sPathCompMan
                End If
            Wend
        End If
                   
        '~~ Assert the root for VB-Development-Projects
        If RootServicedByCompMan = vbNullString Then
            RootServicedByCompMan = mBasic.SelectFolder(sTitle:="Select the root folder for any VB development/maintenance project which is to be supported by CompMan")
        End If
    
    End With

xt: Set fso = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function AssertUpToDateVersion() As Boolean
' ---------------------------------------------------------------------------------------------------------------------
' Returns TRUE when the version of the Addin instance Workbook is identical with the delopment instance Workbook.
' The version of the Addin instance is obtained via the call of "Version" by Application.Run. Because
' ByRef parameter is not supported by this method the value is returned via the Version class object.
' See: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1022_ByRef_Argument_with_the_Application_Run_method.shtml
' ---------------------------------------------------------------------------------------------------------------------
    Const PROC = "AssertUpToDateVersion"
    
    On Error GoTo eh
    Dim lStep       As Long
    Dim cVersion    As New clsAddinVersion

    If cllLog Is Nothing Then Set cllLog = New Collection
    Set cVersion = New clsAddinVersion
    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Assert the up-to-date version"
    
    Application.StatusBar = "RenewAddIn " & lStep & ". Assert the up-to-date version"
    If AddInInstncWrkbkIsOpen Then
        Application.Run wbTarget.name & "!mCompMan.Version", cVersion
        If cVersion.Version = AddInVersion Then
            cllLog.Add "RenewAddIn " & lStep & ". Passed! Renew of the Addin (version " & cVersion.Version & ") successfull! Addin is open again and active :-)"
            bSucceeded = True
        Else
            cllLog.Add "RenewAddIn " & lStep & ". Failed! Renew of the Addin failed (is still version " & cVersion.Version & ":-("
            bSucceeded = False
        End If
    End If

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub ConfirmConfig()
    Const PROC = "ConfirmConfig"
    Const CFG_CONFIRMED = "Confirmed"
    
    On Error GoTo eh
    Dim sMsg            As tMsg
    Dim sReply          As String
    
    With sMsg
        .section(1).sText = "Confirm or (re-)configure the location (path) for the CompMan AddIn instance Workbook which currently is:"
        .section(2).sText = CompManAddinPath
        .section(2).bMonspaced = True
        .section(3).sText = "Confirm or (re-)configure the root folder for any VB-Project in status development/maintenance. " & _
                            "Only Workbooks(VB-Projects in any subfolder will be supported by CompMan!" & vbLf & _
                            "The current root folder is:"
        .section(4).sText = RootServicedByCompMan
        .section(4).bMonspaced = True
    End With
    
    While sReply <> CFG_CONFIRMED
        sReply = mMsg.Dsply(msg_title:="Confirm or change the current Component Management's basic configuration" _
                          , msg:=sMsg _
                          , msg_buttons:=mMsg.Buttons(CFG_CONFIRMED, vbLf, CFG_CHANGE_ADDIN_LCTN, CFG_CHANGE_DEVELOPMENT_ROOT) _
                           )
        Select Case sReply
            Case CFG_CHANGE_ADDIN_LCTN
                Do
                    CompManAddinPath = mBasic.SelectFolder("Select the ""obligatory!"" folder for the AddIn instance of the CompManDev Workbook")
                    If CompManAddinPath <> vbNullString Then Exit Do
                Loop
                TrustThisFolder FolderPath:=CompManAddinPath
            
            Case CFG_CHANGE_DEVELOPMENT_ROOT
                Do
                    RootServicedByCompMan = mBasic.SelectFolder("Select the ""obligatory!"" root folder for VB development projects about to be supported by CompMan")
                    If RootServicedByCompMan <> vbNullString Then Exit Do
                Loop
        End Select
    Wend
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ControlItemsAdd()
'------------------------------------
' Add control to "Add_Ins" popup menu
'------------------------------------
    Const PROC = "ControlItemsAdd"
    
    On Error GoTo eh
    Dim cmb   As CommandBar
    Dim cmbb  As CommandBarButton
    
    ControlItemsRemove
    Set cmb = Application.CommandBars("Worksheet Menu Bar")
    Set cmbb = cmb.Controls.Add(Type:=msoControlButton, id:=2950)
    With cmbb
        .caption = CONTROL_CAPTION_RENEW
        .Style = msoButtonCaption
        .TooltipText = "Saves the development instance '" & DevInstncName & "' as '" & AddInInstanceName & "' Addin"
        .OnAction = "RenewAddIn"
        .Visible = True
    End With
    
    Set cmbb = cmb.Controls.Add(Type:=msoControlButton, id:=2950)
    With cmbb
        .caption = CONTROL_CAPTION_PAUSE
        .Style = msoButtonCaption
        .TooltipText = "Pauses (temporarily removes) the " & AddInInstanceName & " Addin"
        .OnAction = "PauseAddIn"
        .Visible = True
    End With
    
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ControlItemsRemove()
'--------------------------------------------------------------
' Remove the "RenewAddIn" control item from the "Add_Ins" popup menu
'--------------------------------------------------------------
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(CONTROL_CAPTION_RENEW).Delete
    Application.CommandBars("Worksheet Menu Bar").Controls(CONTROL_CAPTION_PAUSE).Delete
End Sub

Private Sub DevInstncWorkbookClose()
    Const PROC = "DevInstncWorkbookClose"
    
    On Error GoTo eh
    Dim lStep As Long

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Close the Development instance Workbook"

    On Error Resume Next
    wbDevlp.Activate
    wbDevlp.Close False
    If Err.Number <> 0 _
    Then cllLog.Add "RenewAddIn " & lStep & ". Failed! Closing the Development instance Workbook '" & DevInstncName & "' failed with:" & vbLf & _
                    "  (" & Err.Description & ")" _
    Else cllLog.Add "RenewAddIn " & lStep & ". Passed! Development instance Workbook '" & DevInstncName & "' successfully closed"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub DevInstncWorkbookDelete()
    Const PROC = "DevInstncWorkbookDelete"
    
    On Error GoTo eh
    Dim lStep   As Long
    Dim fso     As New FileSystemObject
    
    On Error Resume Next
    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Delete the Development instance Workbook"
    
    With fso
        If .FileExists(DevInstncFullName) Then
            .DeleteFile DevInstncFullName
            If Err.Number = 0 _
            Then cllLog.Add "RenewAddIn " & lStep & ". Passed! Development instance Workbook '" & DevInstncName & "' successfully deleted" _
            Else cllLog.Add "RenewAddIn " & lStep & ". Failed! Deleting the Development instance Workbook '" & DevInstncName & "' failed with:" & vbLf & _
                         "  (" & Err.Description & ")"
        Else
            cllLog.Add "RenewAddIn " & lStep & ". Passed! Development instance Workbook '" & DevInstncName & "' found already deleted"
        End If
    End With

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function DevInstncWorkbookExists() As Boolean
Dim fso As New FileSystemObject
    DevInstncWorkbookExists = fso.FileExists(DevInstncFullName)
End Function

Private Function DevInstncWorkbookIsOpen() As Boolean
    On Error Resume Next
    Set wbDevlp = Application.Workbooks(DevInstncName)
    DevInstncWorkbookIsOpen = Err.Number = 0
End Function

Private Sub DevInstncWorkbookSave()
    Const PROC = "DevInstncWorkbookSave"
    
    On Error GoTo eh
    Dim lStep As Long

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Save the Development instance Workbook"
    
    Set wbSource = Application.Workbooks(DevInstncName)
    Application.EnableEvents = False
    wbSource.Save
    Application.EnableEvents = True
    wbSource.Activate
    cllLog.Add "RenewAddIn " & lStep & ". Passed! The Development instance Workbook '" & DevInstncName & "' has successfully been svaed"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub SaveDevInstncWorkbookAsAddin()
    Const PROC = "SaveDevInstncWorkbookAsAddin"
    
    On Error GoTo eh
    Dim lStep   As Long

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Save the Development instance Workbook as Addin"
    
    With Application
        If Not AddInInstncWrkbkExists Then
            '~~ At this point the Addin must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbSource.SaveAs AddInInstanceFullName, FileFormat:=xlAddInFormat
            If Err.Number <> 0 Then
                cllLog.Add "RenewAddIn " & lStep & ". Failed! Save Development instance (version " & AddInVersion & ") as Addin failed!"
            Else
                cllLog.Add "RenewAddIn " & lStep & ". Passed! The Development instance (version " & AddInVersion & ") has successfully been saved as Addin instance Workbook '" & AddInInstanceName & "'"
            End If
            .EnableEvents = True
'            mCompMan.ExportChangedComponents wbDevlp
        Else ' file still exists
            cllLog.Add "RenewAddIn " & lStep & ". Failed! Setup/renew of the Addin with version " & AddInVersion & " of the development instance failed"
        End If
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub DisplayRenewResult()

    Dim sMsg        As String
    Dim i           As Long
    Dim sTitle      As String
    
    If bSucceeded And Not bAddInPaused Then
        sTitle = "Successful! The Addin '" & AddInInstanceName & "' has been renewed by the development instance '" & DevInstncName & "' (see details below)"
    ElseIf bSucceeded And bAddInPaused Then
        sTitle = "Successful! The Addin is now paused (removed!) until the next renew"
    ElseIf Not bSucceeded And Not bAddInPaused Then
        sTitle = "Failed! Renewing the Addin '" & AddInInstanceName & "' by the development instance failed (see details below)"
    ElseIf Not bSucceeded And bAddInPaused Then
        sTitle = "Failed! Pausing the Addin instance of '" & AddInInstanceName & "' failed!"
    End If
    
    If cllLog.Count > 0 Then
        sMsg = cllLog.Item(1)
        For i = 2 To cllLog.Count
            sMsg = sMsg & vbLf & cllLog.Item(i)
        Next i
        mMsg.Box msg_title:=sTitle _
               , msg:=sMsg _
               , msg_monospaced:=True
    End If
    Application.StatusBar = vbNullString
    lStep = 0

End Sub

Private Sub DisplaySaveAsDevResult()

    Dim sMsg        As String
    Dim i           As Long

    If cllLog.Count > 0 Then
        sMsg = cllLog.Item(1)
        For i = 2 To cllLog.Count
            sMsg = sMsg & vbLf & cllLog.Item(i)
        Next i
        If bSucceeded _
        Then mMsg.Box msg_title:="Successful! The Addin '" & AddInInstanceName & "' has been saved as Development instance Workbook '" & DevInstncName & "' (see details below)" _
                    , msg:=sMsg _
                    , msg_monospaced:=True _
        Else mMsg.Box msg_title:="Failed! Saving the Addin " & AddInInstanceName & " as Development instance Workbook failed (see details below)" _
                    , msg:=sMsg _
                    , msg_monospaced:=True
    End If
    Application.StatusBar = vbNullString
    lStep = 0
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Private Sub RestoreReferencesToAddIn()
    Const PROC = "RestoreReferencesToAddIn"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim wb              As Workbook
    Dim sWbs            As String
    Dim bOneRestored    As Boolean
    Dim lStep           As Long
    
    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Restore references to the Addin for open Workbooks"
    
    For Each v In dctAddInRefs
        Set wb = v
        wb.VBProject.References.AddFromFile AddInInstanceFullName
        sWbs = wb.name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        cllLog.Add "RenewAddIn " & lStep & ". Passed! Saved references to the Addin in open Workbooks have been restored"
        cllLog.Add "   (" & sWbs & ")"
    Else
        cllLog.Add "RenewAddIn " & lStep & ". Passed! Restoring references did not find any saved to restore"
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub SaveAndRemoveReferencesToAddIn()
' ----------------------------------------------------------------
' - Allows the user to close any open Workbook which refers to the
'   Addin, which definitly hinders the Addin from being re-newed.
' - Returns TRUE when the user closed all open Workbboks.
' ----------------------------------------------------------------
    Const PROC = "SaveAndRemoveReferencesToAddIn"
    
    On Error GoTo eh
    Dim lStep       As Long
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim wb          As Workbook
    Dim ref         As Reference
    Dim sWbs        As String
    Dim bOneRemoved As Boolean
    Dim fso         As New FileSystemObject

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Save and remove references to the Addin from open Workbooks"
    
    With Application
        Set dct = mWrkbk.Opened ' Returns a Dictionary with all open Workbooks in any application instance
        Set dctAddInRefs = New Dictionary
        For Each v In dct
            Set wb = dct.Item(v)
            For Each ref In wb.VBProject.References
                If InStr(ref.name, fso.GetBaseName(AddInInstanceName)) <> 0 Then
                    dctAddInRefs.Add wb, ref
                    sWbs = wb.name & ", " & sWbs
                End If
            Next ref
        Next v
        
        For Each v In dctAddInRefs
            Set wb = v
            Set ref = dctAddInRefs(v)
            wb.VBProject.References.Remove ref
            bOneRemoved = True
        Next v
        bAllRemoved = True
    End With
    
    If bOneRemoved Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        cllLog.Add "RenewAddIn " & lStep & ". Passed! References to the Addin from open Workbooks saved and removed"
        cllLog.Add "   (" & sWbs & ")"
    Else
        cllLog.Add "RenewAddIn " & lStep & ". Passed! None of the open Workbooks referred to the Addin '" & AddInInstanceName & "'"
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub PauseAddIn()
    RenewAddIn bPause:=False
End Sub

Public Sub RenewAddIn(Optional ByVal bPause As Boolean = False)
' -----------------------------------------------------------
' Renews the code of the Addin instance of this Workbook with
' this Workbook's code by displaying a detailed result of the
' whole RenewAddIn process.
' Note: It cannot be avoided that this procedure is available
'       also in the Addin instance. However, its execution is
'       limited to this Workbook's development instance.
' -----------------------------------------------------------
    Const PROC = "RenewAddIn"
    
    On Error GoTo eh
    Dim lStep   As Long
    
    Set cllLog = New Collection
    
    '~~ Assert ThisWorkbook is the development instance of the CompMan Addin
    lStep = Step
    If Not IsDevInstnc() Then
        cllLog.Add "RenewAddIn " & lStep & ". Failed! The 'RenewAddIn' service had not been executed from within the development instance Workbook (" & DevInstncName & ")!"
        GoTo xt
    Else
        cllLog.Add "RenewAddIn " & lStep & ". Passed! The 'RenewAddIn service has been initiated from within the Development instance Workbook (" & DevInstncName & ") asserted"
    End If
                     
    If Not bAddInPaused Then
        '~~ Get the current CompMan's base configuration confirmed or changed
        ConfirmConfig
                         
        '~~ Assert no Workbooks are open referring to the Addin
        SaveAndRemoveReferencesToAddIn
        If Not bAllRemoved Then GoTo xt
    
        '~~ Assure the current version of the AddIn's development instance has been saved
        '~~ Note: Unconditionally saving the Workbook does an incredible trick:
        '~~       The un-unstalled and IsAddin=False Workbook is released from the Application
        '~~       and no longer considered "used"
        DevInstncWorkbookSave
        wbSource.Activate
              
        '~~ Attempt to turn Addin to "IsAddin=False", uninstall and close it
        If AddInInstncWrkbkIsOpen Then
            Set_IsAddin_ToFalse wbTarget
            CloseAddinInstncWorkbook
        End If
        
        '~~ Attempt to delete the Addin Workbook file
        DeleteAddInInstanceWorkbook
    End If
    
    If bPause Then
        bAddInPaused = True
        If Not AddInInstncWrkbkIsOpen And Not bSucceeded Then bSucceeded = True
        GoTo xt
    End If
    
    bAddInPaused = False
    '~~ Attempt to save the development instance as Addin
    SaveDevInstncWorkbookAsAddin
    
    '~~ Saving the development instance as Addin may also open the Addin.
    '~~ So if not already open it is re-opened and thus re-activated
    OpenAddinInstncWorkbook
    
    '~~ Assert the correct version has been renewed/re-opened
    AssertUpToDateVersion
    
    '~~ Re-instate references to the Addin which had been removed
    RestoreReferencesToAddIn

xt: DisplayRenewResult
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub SaveAsDev()
' ------------------------------------------------------------
' Saves the Addin instance Workbook as Developement instance.
' Note:
' This undertaking is a bit delicate since the code is executed
' in a Workbook which is saved as another one!
'
' Essential: The save location is the configured Common
'            Component Workbooks folder
' ------------------------------------------------------------
    Const PROC = "SaveAsDev"
    
    On Error GoTo eh
    Dim lStep   As Long

    Set cllLog = New Collection
        
    '~~ Assert ThisWorkbook is the development instance of the CompMan Addin
    lStep = Step
    If Not IsAddinInstnc() Then
        cllLog.Add "RenewAddIn " & lStep & ". Failed! Saving the '" & AddInInstanceName & "' as Development instance failed: It may have not been executed from within the Addin instance Workbook!"
        GoTo xt
    Else
        cllLog.Add "RenewAddIn " & lStep & ". Passed! The Addin instance '" & AddInInstanceName & "' has been saved as Development instance '" & DevInstncName & "'"
    End If
    Set wbSource = Application.Workbooks(AddInInstanceName)
    
    '~~ Save and remove references to the Addin from any open Workbook referring to it
    SaveAndRemoveReferencesToAddIn
    If Not bAllRemoved Then GoTo xt

    '~~ Attempt to close the Development instance Workbook when open
    If DevInstncWorkbookIsOpen Then
        If MsgBox(Prompt:="The Development instance Workbook of the Addin is still open. Saving the Addin as Development instance will cause any code modifications get lost. Reply with Ok or Cancel" _
                , Buttons:=vbOKCancel _
                , Title:="Development instance Workbook still open" _
                 ) = vbOK Then
            DevInstncWorkbookClose
        Else
            GoTo xt
        End If
    End If
    
    '~~ Attempt to delete the Addin Workbook file
    DevInstncWorkbookDelete
    
    '~~ Turn the Workbook property "IsAddin" to "False"
    Set_IsAddin_ToFalse wb:=wbSource
    
    '~~ Attempt to save the development instance as Addin
    SaveAddinInstncWorkbookAsDevlp
    
    '~~ Re-instate references to the Addin which had been removed
    RestoreReferencesToAddIn
    bSucceeded = True

xt: DisplaySaveAsDevResult
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Set_IsAddin_ToFalse(ByVal wb As Workbook)
    Const PROC = "Set_IsAddin_ToFalse"
    
    On Error GoTo eh
    Dim lStep As Long

    lStep = Step
    Application.StatusBar = "RenewAddIn " & lStep & ". Set the IsAddin property to False"
    If wb.IsAddin = True Then
        wb.IsAddin = False
        cllLog.Add "RenewAddIn " & lStep & ". Passed! The 'IsAddin' property of the Addin Workbook has been set to FALSE"
    Else
        cllLog.Add "RenewAddIn " & lStep & ". Passed! The 'IsAddin' property of the Addin Workbook was already set to FALSE"
    End If
    
xt:     Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub TrustThisFolder(Optional ByVal FolderPath As String, _
                            Optional ByVal TrustNetworkFolders As Boolean = False, _
                            Optional ByVal sDescription As String)
' ---------------------------------------------------------------------------
' Add a folder to the 'Trusted Locations' list so that your project's VBA can
' open Excel files without raising errors like "Office has detected a problem
' with this file. To help protect your computer this file cannot be opened."
' Ths function has been implemented to fail silently on error: if you suspect
' that users don't have permission to assign 'Trusted Location' status in all
' locations, reformulate this as a function returning True or False
'
' Nigel Heffernan January 2015
'
' Based on code published by Daniel Pineault in DevHut.net on June 23, 2010:
' www.devhut.net\2010\06\23\vbscript-createset-trusted-location-using-vbscript\
' **** **** **** ****  THIS CODE IS IN THE PUBLIC DOMAIN  **** **** **** ****
' UNIT TESTING:
'
' 1:    Reinstate the commented-out line 'Debug.Print sSubKey & vbTab & sPath
' 2:    Open the Immediate Window and run this command:
'           TrustThisFolder "Z:\", True, True, "The user's home directory"
' 3:    If  "Z:\"  is already in the list, choose another folder
' 4:    Repeat step 2 or 3: the folder should be listed in the debug output
' 5:    If it isn't listed, disable the error-handler and record any errors
' -----------------------------------------------------------------------------
    Const PROC = "TrustThisFolder"
    Const HKEY_CURRENT_USER = &H80000001
    
    On Error GoTo eh
    Dim sKeyPath            As String
    Dim oRegistry           As Object
    Dim sSubKey             As String
    Dim oSubKeys            As Variant   ' type not specified. After it's populated, it can be iterated
    Dim oSubKey             As Variant   ' type not specified.
    Dim bSubFolders         As Boolean
    Dim bNetworkLocation    As Boolean
    Dim iTrustNetwork       As Long
    Dim sPath               As String
    Dim i                   As Long
    Dim fso                 As New FileSystemObject
    
    bSubFolders = True
    bNetworkLocation = False

    With fso
        If FolderPath = "" Then
            FolderPath = .GetSpecialFolder(2).Path
            If sDescription = "" Then
                sDescription = "The user's local temp folder"
            End If
        End If
    End With

    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If

    sKeyPath = ""
    sKeyPath = sKeyPath & "SOFTWARE\Microsoft\Office\"
    sKeyPath = sKeyPath & Application.Version
    sKeyPath = sKeyPath & "\Excel\Security\Trusted Locations\"
     
    Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
    '~~ Note: not the usual \root\cimv2  for WMI scripting: the StdRegProv isn't in that folder
    oRegistry.EnumKey HKEY_CURRENT_USER, sKeyPath, oSubKeys
    
    For Each oSubKey In oSubKeys
        sSubKey = CStr(oSubKey)
        oRegistry.GetStringValue HKEY_CURRENT_USER, sKeyPath & "\" & sSubKey, "Path", sPath
        If sPath = FolderPath Then
            Exit For
        End If
    Next oSubKey
    
    If sPath <> FolderPath Then
        If IsNumeric(Replace(sSubKey, "Location", "")) _
        Then i = CLng(Replace(sSubKey, "Location", "")) + 1 _
        Else i = UBound(oSubKeys) + 1
        
        sSubKey = "Location" & CStr(i)
        
        If TrustNetworkFolders Then
            iTrustNetwork = 1
            oRegistry.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, "AllowNetworkLocations", iTrustNetwork
            If iTrustNetwork = 0 Then
                oRegistry.SetDWORDValue HKEY_CURRENT_USER, sKeyPath, "AllowNetworkLocations", 1
            End If
        End If
        
        oRegistry.CreateKey HKEY_CURRENT_USER, sKeyPath & "\" & sSubKey
        oRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath & "\" & sSubKey, "Path", FolderPath
        oRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath & "\" & sSubKey, "Description", sDescription
        oRegistry.SetDWORDValue HKEY_CURRENT_USER, sKeyPath & "\" & sSubKey, "AllowSubFolders", 1

    End If

xt: Set fso = Nothing
    Set oRegistry = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub UpdateRawClones()
' -------------------------------------------------
' Update Common Components used by this VBProject
' via the CompMan AddIn required to change any code
' in this VBProject. Performed at Workbook_Open.
' Note:
' This procedure only runs when
' - The Development instance Workbook is opened
' - and the Addin Instance Workbook is also open
'   (the latter will not be the case when Excel
'   is opend only for the Development Instance!)
' -------------------------------------------------
    Const PROC = "UpdateRawClones"
    
    On Error GoTo eh
    
    If IsDevInstnc Then
        If AddInInstncWrkbkIsOpen Then
            Application.Run AddInInstanceName & "!mCompMan.UpdateRawClones", ActiveWorkbook, wbAddIn.HOSTED_RAWS
        End If
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub
