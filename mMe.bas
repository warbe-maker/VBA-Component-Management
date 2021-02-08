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
' - CfgAsserted          Returns True when the required properties (paths)
'                           are configured and exist
' - Renew_1_ConfirmConfig   Get the current configured paths confirmed before
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
' - UpdateRawClones         Exclusively performed fro within the Development
'                           instance Workbook to update its own used Clone
'                           Components of which the Raw code has changed. This
'                           service only runs provided the Addin instance
'                           Workbook is open (see RenewAddIn service).
' - RootServicedByCompMan   When the AddIn is open and not  p a u s e d the
'                           Workbook_Open service 'UpdateRawClones' and the
'                           Workbook_BeforeSave service 'ExportChangedComponents'
'                           will be executed for Workbooks calling them under
'                           the two additional preconditions:
'                           1. The Conditional Compile Argument CompMan = 1"
'                           2. The Workbook is located in the configured
'                              'Serviced Development Root'
' Uses Common Components:
' - mFile                   Get/Let PrivateProperty value service
' - mWrkbk                  GetOpen and Opened service
' - mMsg                    Dsply, Box, and Buttons service used by the
'                           RenewAddin,  Renew_1_ConfirmConfig service
'
' Requires:
' W. Rauschenberger, Berlin Nov 2020
' ---------------------------------------------------------------------------
Private Const CONTROL_CAPTION_RENEW As String = "Renew Addin"
Private Const CONTROL_CAPTION_PAUSE As String = "Pause Addin"
Private Const SECTION_BASE_CONFIG   As String = "BaseConfiguration"
Private Const ROOT_COMPMAN_SERVICED As String = "VBDevProjectsRoot"
Private Const COMPMAN_ADDIN_LCTN    As String = "CompManAddInPath"
Private Const COMPMAN_ADDIN_PAUSED  As String = "CompManAddInPaused"

Private Const ADDIN_WORKBOOK        As String = "CompMan.xlam"      ' Extension adjusted when the above is saved as addin
Private Const ADDIN_VERSION         As String = "5.0"               ' Allows to check the success of an Addin renew
Private Const DEVLP_WORKBOOK        As String = "CompManDev.xlsb"   ' Extension depends on Excel versions
Private wbDevlp                     As Workbook
Private wbSource                    As Workbook                     ' This development instance as the renew source
Private wbTarget                    As Workbook                     ' The Addin instance as renew target
Private bSucceeded                  As Boolean
Private bAllRemoved                 As Boolean
Private dctAddInRefs                As Dictionary
Private lStep                       As Long
Private sRenewLogFile               As String
Private lRenewStep                  As Long
Private sRenewAction                As String
Private bRenewLogFileAppend         As Boolean

'Public Property Get RenewAction() As String:            RenewAction = sRenewAction:                                         End Property
Public Property Get AddInInstanceFullName() As String:  AddInInstanceFullName = AddInPath & DBSLASH & AddInInstanceName:    End Property

Public Property Get AddInInstanceName() As String:      AddInInstanceName = ADDIN_WORKBOOK:                                 End Property

Private Property Get AddInPath() As String:             AddInPath = CompManAddinPath:                                  End Property

Public Property Get AddInPaused() As Boolean
    Dim s As String
    s = Value(pp_section:=SECTION_BASE_CONFIG, pp_value_name:=COMPMAN_ADDIN_PAUSED)
    If s = vbNullString Then
        AddInPaused = False
    Else
        AddInPaused = VBA.CBool(s)
    End If
End Property

Public Property Let AddInPaused(ByVal b As Boolean)
    Value(pp_section:=SECTION_BASE_CONFIG, pp_value_name:=COMPMAN_ADDIN_PAUSED) = b
    With New FileSystemObject
        .CopyFile CFG_FILENAME, CompManAddinPath & "\" & .GetFileName(CFG_FILENAME)
    End With
End Property

Private Property Get ADDIN_FORMAT() As XlFileFormat ' = ... needs adjustment when the above is changed
    ADDIN_FORMAT = xlOpenXMLAddIn
End Property

Private Property Get BTTN_CHANGE_ADDIN_LCTN() As String
    BTTN_CHANGE_ADDIN_LCTN = "Change CompMan" & vbLf & "AddIn Path"
End Property

Private Property Get BTTN_CHANGE_DEVELOPMENT_ROOT() As String
    BTTN_CHANGE_DEVELOPMENT_ROOT = "Change development" & vbLf & "root folder"
End Property

Private Property Get CFG_FILENAME() As String: CFG_FILENAME = ThisWorkbook.Path & "\CompMan.cfg": End Property

Public Property Get CompManAddinPath() As String
    CompManAddinPath = Value(pp_section:=SECTION_BASE_CONFIG, pp_value_name:=COMPMAN_ADDIN_LCTN)
End Property

Public Property Let CompManAddinPath(ByVal s As String)
    Const PROC = "CompManAddinPath-Let"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    Value(pp_section:=SECTION_BASE_CONFIG, pp_value_name:=COMPMAN_ADDIN_LCTN) = s
    fso.CopyFile CFG_FILENAME, VBA.Replace(s & "\", "\\", "\")
    
xt: Set fso = Nothing
    Exit Property

eh: ErrMsg ErrSrc(PROC)
End Property

Public Property Get DevInstncFullName() As String
Dim fso As New FileSystemObject
    DevInstncFullName = RootServicedByCompMan & DBSLASH _
                          & fso.GetBaseName(DevInstncName) & DBSLASH _
                          & DevInstncName
End Property

Public Property Get DevInstncName() As String:          DevInstncName = DEVLP_WORKBOOK:                                 End Property

Private Property Get DEVLP_FORMAT() As XlFileFormat  ' = .xlsb ! may require adjustment when the above is changed
    DEVLP_FORMAT = xlExcel12
End Property

Public Property Get IsAddinInstnc() As Boolean:         IsAddinInstnc = ThisWorkbook.name = AddInInstanceName:          End Property

Public Property Get IsDevInstnc() As Boolean:           IsDevInstnc = ThisWorkbook.name = DevInstncName:                End Property

Private Property Get RenewFinalResult() As String
    If bSucceeded _
    Then RenewFinalResult = "Successful! The Addin '" & AddInInstanceName & "' has been renewed by the development instance '" & DevInstncName & "'" _
    Else RenewFinalResult = "Failed! Renewing the Addin '" & AddInInstanceName & "' by the development instance '" & DevInstncName & "' failed!"
End Property

Public Property Let RenewLogAction(ByVal la_action As String)
    lRenewStep = lRenewStep + 1
    If lRenewStep = 1 Then
        sRenewLogFile = mFile.Temp(tmp_extension:=".log")
        bRenewLogFileAppend = False
    Else
        bRenewLogFileAppend = True
    End If
    sRenewAction = la_action
    wsAddIn.LogRenewStep rn_action:=sRenewAction
End Property

Public Property Let RenewLogResult( _
                    Optional ByVal la_result_text As String = vbNullString, _
                             ByVal la_result As String)
' ---------------------------------------------------------------------------
'
' ---------------------------------------------------------------------------
    Dim s As String
    
    If la_result_text = vbNullString _
    Then s = lRenewStep & " " & la_result & " " & sRenewAction _
    Else s = lRenewStep & " " & la_result & " " & la_result_text
    
    mFile.Txt(ft_file:=sRenewLogFile _
            , ft_append:=bRenewLogFileAppend _
             ) = s
    wsAddIn.LogRenewStep rn_result:=la_result, rn_action:=la_result_text

End Property

Public Property Get RenewStep() As Long:    RenewStep = lRenewStep: End Property

Public Property Let RenewStep(ByVal l As Long)
    If l = 0 Then
        wsAddIn.Range("rngRenewLog").ClearContents
    End If
    lRenewStep = l
End Property

Public Property Get RootServicedByCompMan() As String
    RootServicedByCompMan = Value(pp_section:=SECTION_BASE_CONFIG, pp_value_name:=ROOT_COMPMAN_SERVICED)
End Property

Public Property Let RootServicedByCompMan(ByVal s As String)
    Value(pp_section:=SECTION_BASE_CONFIG, pp_value_name:=ROOT_COMPMAN_SERVICED) = s
    With New FileSystemObject
        .CopyFile CFG_FILENAME, CompManAddinPath & "\" & .GetFileName(CFG_FILENAME)
    End With
End Property

Public Property Get ServicesLogFile(Optional ByVal slf_serviced_wb As Workbook) As String
    Dim fso As New FileSystemObject
    ServicesLogFile = slf_serviced_wb.Path & "\" & fso.GetBaseName(mMe.AddInInstanceName) & ".Services.log"
    Set fso = Nothing
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
    
    Value = mFile.Value(pp_file:=CFG_FILENAME _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
End Property

Private Property Let Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
    mFile.Value(pp_file:=CFG_FILENAME _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value
End Property

Public Property Get xlAddInFormat() As Long:            xlAddInFormat = ADDIN_FORMAT:                                       End Property

Public Property Get xlDevlpFormat() As Long:            xlDevlpFormat = DEVLP_FORMAT:                                       End Property

Public Sub AddInContinue()
    mMe.AddInPaused = False
    wsAddIn.CompManAddInStatus = _
        "The AddIn is currently  a c t i v e ! The Workbook_Open service 'UpdateRawClones' " & _
        "and the Workbook_BeforeSave service 'ExportChangedComponents' will be executed for " & _
        "Workbooks calling them under the two additional preconditions: " & vbLf & _
        "1. The Conditional Compile Argument CompMan = 1" & vbLf & _
        "2. The Workbook is located in the configured 'Serviced Development Root' which currently is:" & vbLf & _
        "   '" & mMe.RootServicedByCompMan & "'"
 
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

Public Sub AddInPause()
    mMe.AddInPaused = True
    wsAddIn.CompManAddInStatus = _
        "The AddIn is currently  p a u s e d ! The Workbook_Open service 'UpdateRawClones' " & _
        "and the Workbook_BeforeSave service 'ExportChangedComponents' will be bypassed " & _
        "until the Addin is 'continued' again!"
End Sub

Public Function AddInVersion(Optional ByRef sVersion As String) As String
    sVersion = ADDIN_VERSION
    AddInVersion = sVersion
End Function

Public Function CfgAsserted() As Boolean
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
            TrustThisFolder FolderPath:=CompManAddinPath
            CfgAsserted = True
            .CopyFile Source:=CFG_FILENAME, Destination:=CompManAddinPath & "\CompMan.cfg", OverWriteFiles:=True
            GoTo xt
        End If
                
        '~~ Assert the folder for the CompMan AddIn
        sPathCompMan = CompManAddinPath
        If sPathCompMan = vbNullString Then
            While sPathCompMan = vbNullString
                '~~ Selecting a folder is obligatory !
                sPathCompMan = mBasic.SelectFolder( _
                            sTitle:="Select the folder dedicated to the AddIn instance of the CompManDev Workbook")
            Wend
            '~~ Assure trust in this location and save it to the CompMan.cfg file
            TrustThisFolder FolderPath:=sPathCompMan
            CompManAddinPath = sPathCompMan
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
            RootServicedByCompMan = _
            mBasic.SelectFolder(sTitle:="Select the root folder for any VB-Project where it will exclusively be serviced by the CompMan Addin")
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

Private Sub DevInstncWorkbookClose()
    Const PROC = "DevInstncWorkbookClose"
    
    On Error GoTo eh

    mMe.RenewLogAction = "Close Development instance Workbook '" & DevInstncName & "'"
    
    On Error Resume Next
    wbDevlp.Activate
    wbDevlp.Close False
    If Err.Number <> 0 _
    Then mMe.RenewLogResult("Closing the Development instance Workbook '" & DevInstncName & "' failed with:" & vbLf & _
                            "(" & Err.Description & ")" _
                           ) = "Failed" _
    Else mMe.RenewLogResult() = "Passed"
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
    Dim fso As New FileSystemObject
    
    mMe.RenewLogAction = "Delete Development instance Workbook '" & DevInstncName & "'"
    
    With fso
        If .FileExists(DevInstncFullName) Then
            On Error Resume Next
            .DeleteFile DevInstncFullName
            If Err.Number = 0 _
            Then mMe.RenewLogResult() = "Passed" _
            Else mMe.RenewLogResult("Deleting the Development instance Workbook '" & DevInstncName & "' failed with:" & vbLf & _
                                    "(" & Err.Description & ")" _
                                   ) = "Failed"
        Else
            mMe.RenewLogResult() = "Passed"
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Public Sub RenewAddIn()
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
    RenewStep = 0

    Application.EnableEvents = False
    bSucceeded = False
       
    mMe.RenewLogAction = "Assert the Renew service is executed from the development instance Workbook"
    If Not IsDevInstnc() Then
        mMe.RenewLogResult("The 'RenewAddIn' service had not been executed from within the development instance Workbook (" & DevInstncName & ")!" _
                          ) = "Failed"
        GoTo xt
    Else
        mMe.RenewLogResult = "Asserted"
    End If
                     
    '~~ Get the current CompMan's base configuration confirmed or changed
    Renew_1_ConfirmConfig
                         
    '~~ Assert no Workbooks are open referring to the Addin
    Renew_2_SaveAndRemoveAddInReferences
    If Not bAllRemoved Then GoTo xt

    '~~ Assure the current version of the AddIn's development instance has been saved
    '~~ Note: Unconditionally saving the Workbook does an incredible trick:
    '~~       The un-unstalled and IsAddin=False Workbook is released from the Application
    '~~       and no longer considered "used"
    Renew_3_DevInstncWorkbookSave
    wbSource.Activate
          
    '~~ Attempt to turn Addin to "IsAddin=False", uninstall and close it
    If AddInInstncWrkbkIsOpen Then
        Renew_4_Set_IsAddin_ToFalse wbTarget
        Renew_5_CloseAddinInstncWorkbook
    End If
    
    '~~ Attempt to delete the Addin Workbook file
    Renew_6_DeleteAddInInstanceWorkbook
        
    '~~ Attempt to save the development instance as Addin
    Renew_7_SaveDevInstncWorkbookAsAddin
    
    '~~ Saving the development instance as Addin may also open the Addin.
    '~~ So if not already open it is re-opened and thus re-activated
    Renew_8_OpenAddinInstncWorkbook
        
    '~~ Re-instate references to the Addin which had been removed
    Renew_9_RestoreReferencesToAddIn
    
    bSucceeded = True
    
xt: mMe.RenewLogAction = RenewFinalResult
    Application.EnableEvents = True
'    RenewLogDisplay
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub CfgConfirm()
' ------------------------------------------------------
'
' ------------------------------------------------------
    Const PROC = "CfgConfirm"
    Const BTTN_CFG_CONFIRMED = "Confirmed"
    
    On Error GoTo eh
    Dim sMsg            As tMsg
    Dim sReply          As String
    
    With sMsg
        .Section(1).sLabel = "Dedicated folder for the CompMan Add-in:"
        .Section(1).sText = "Please confirm the currently configured folder for the CompMan AddIn or change it if not. " & _
                            "Make sure that the CompMan Add-in resides in its own dedicated location, " & _
                            "i.e. one different from Excel's default user-specific! Roaming folder."
        .Section(2).sText = CompManAddinPath
        .Section(2).bMonspaced = True
        .Section(3).sLabel = "CompMan serviced root:"
        .Section(3).sText = "Please confirm the root folder serviced by CompMan. Only Workbooks/VB-Projects " & _
                            "in any subfolder of the configured root will be serviced when opened or closed."
        .Section(4).sText = RootServicedByCompMan
        .Section(4).bMonspaced = True
    End With
    
    While sReply <> BTTN_CFG_CONFIRMED
        sReply = mMsg.Dsply(msg_title:="Confirm or change the current basic configuration for the CompMan Addin" _
                          , msg:=sMsg _
                          , msg_buttons:=mMsg.Buttons(BTTN_CFG_CONFIRMED, vbLf, BTTN_CHANGE_ADDIN_LCTN, BTTN_CHANGE_DEVELOPMENT_ROOT) _
                           )
        Select Case sReply
            Case BTTN_CHANGE_ADDIN_LCTN
                Do
                    CompManAddinPath = mBasic.SelectFolder("Select the ""obligatory!"" folder for the AddIn instance of the CompManDev Workbook")
                    If CompManAddinPath <> vbNullString Then Exit Do
                Loop
                TrustThisFolder FolderPath:=CompManAddinPath
            
            Case BTTN_CHANGE_DEVELOPMENT_ROOT
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

Public Sub Renew_1_ConfirmConfig()
    mMe.RenewLogAction = "Assert current basic configuration"
    CfgConfirm
    mMe.RenewLogResult = "Asserted"
End Sub

Private Sub Renew_2_SaveAndRemoveAddInReferences()
' ----------------------------------------------------------------
' - Allows the user to close any open Workbook which refers to the
'   Addin, which definitly hinders the Addin from being re-newed.
' - Returns TRUE when the user closed all open Workbboks.
' ----------------------------------------------------------------
    Const PROC = "Renew_2_SaveAndRemoveAddInReferences"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim wb          As Workbook
    Dim ref         As Reference
    Dim sWbs        As String
    Dim bOneRemoved As Boolean
    Dim fso         As New FileSystemObject

    mMe.RenewLogAction = "Save and remove references to the Addin from open Workbooks"
    
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
        mMe.RenewLogResult() = "Passed"
    Else
        mMe.RenewLogResult(sRenewAction & vbLf & "None of the open Workbooks referred to the Addin '" & AddInInstanceName & "'" _
                          ) = "Passed"
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Renew_3_DevInstncWorkbookSave()
    Const PROC = "Renew_3_DevInstncWorkbookSave"
    
    On Error GoTo eh
    mMe.RenewLogAction = "Save the Development instance Workbook '" & DevInstncName & "'"
    
    Set wbSource = Application.Workbooks(DevInstncName)
    wbSource.Save
    wbSource.Activate
    mMe.RenewLogResult() = "Passed"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Renew_4_Set_IsAddin_ToFalse(ByVal wb As Workbook)
    Const PROC = "Renew_4_Set_IsAddin_ToFalse"
    
    On Error GoTo eh

    mMe.RenewLogAction = "Set the 'IsAddin' property of the Addin Workbook to FALSE"
    
    If wb.IsAddin = True Then
        wb.IsAddin = False
        mMe.RenewLogResult() = "Passed"
    Else
        mMe.RenewLogResult("The 'IsAddin' property of the Addin Workbook was already set to FALSE" _
                          ) = "Passed"
    End If
    
xt:     Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Renew_5_CloseAddinInstncWorkbook()
    Const PROC = "Renew_5_CloseAddinInstncWorkbook"
    
    mMe.RenewLogAction = "Close Addin Workbook instance"
    On Error Resume Next
    wbSource.Activate
    wbTarget.Close False
    If Err.Number <> 0 _
    Then mMe.RenewLogResult("Closing the Addin instance Workbook '" & AddInInstanceName & "' failed with:" & vbLf & _
                            "(" & Err.Description & ")" _
                           ) = "Failed" _
    Else mMe.RenewLogResult = "Passed"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Renew_6_DeleteAddInInstanceWorkbook()
    Const PROC = "Renew_6_DeleteAddInInstanceWorkbook"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject

    On Error Resume Next
    mMe.RenewLogAction = "Remove the Addin instance Workbook '" & AddInInstanceName & "'"
    With fso
        If .FileExists(AddInInstanceFullName) Then
            .DeleteFile AddInInstanceFullName
            If Err.Number = 0 _
            Then mMe.RenewLogResult = "Passed" _
            Else mMe.RenewLogResult("Deleting the Addin instance Workbook '" & AddInInstanceName & "' failed with:" & vbLf & _
                                    "(" & Err.Description & ")" _
                                   ) = "Failed"
        Else
            mMe.RenewLogResult = "Passed"
        End If
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Renew_7_SaveDevInstncWorkbookAsAddin()
    Const PROC = "Renew_7_SaveDevInstncWorkbookAsAddin"
    
    On Error GoTo eh
    mMe.RenewLogAction = "Save the Development instance (version " & AddInVersion & ") as Addin instance Workbook '" & AddInInstanceName & "'"
    
    With Application
        If Not AddInInstncWrkbkExists Then
            '~~ At this point the Addin must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbSource.SaveAs AddInInstanceFullName, FileFormat:=xlAddInFormat
            If Err.Number <> 0 _
            Then mMe.RenewLogResult("Save Development instance (version " & AddInVersion & ") as Addin failed!" _
                                   ) = "Failed" _
            Else mMe.RenewLogResult() = "Passed"
            .EnableEvents = True
'            mCompMan.ExportChangedComponents wbDevlp
        Else ' file still exists
            mMe.RenewLogResult("Setup/renew of the Addin with version " & AddInVersion & " of the development instance failed" _
                              ) = "Failed"
        End If
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Renew_8_OpenAddinInstncWorkbook()
    Const PROC = "Renew_8_OpenAddinInstncWorkbook"
    
    On Error GoTo eh
    Dim wb              As Workbook
    Dim sBaseAddinName  As String
    Dim sBaseDevName    As String
    Dim fso             As New FileSystemObject
    
    If Not AddInInstncWrkbkIsOpen Then
        If AddInInstncWrkbkExists Then
            mMe.RenewLogAction = "Re-open the Addin Workbook instance '" & AddInInstanceName & "'"
            On Error Resume Next
            Set wb = Application.Workbooks.Open(AddInInstanceFullName)
            If Err.Number = 0 Then
                With fso
                    sBaseAddinName = .GetBaseName(wb.name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.name)
                    wb.VBProject.name = sBaseAddinName
                End With
                mMe.RenewLogResult() = "Passed"
            Else
                mMe.RenewLogResult("(Re)opening the Addin Workbook '" & AddInInstanceName & "' failed with:" & vbLf & _
                                   "(" & Err.Description & ")" _
                                  ) = "Failed"
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

Private Sub Renew_9_RestoreReferencesToAddIn()
    Const PROC = "Renew_9_RestoreReferencesToAddIn"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim wb              As Workbook
    Dim sWbs            As String
    Dim bOneRestored    As Boolean
    
    mMe.RenewLogAction = "Restore references to the Addin in open Workbooks"
    
    For Each v In dctAddInRefs
        Set wb = v
        wb.VBProject.References.AddFromFile AddInInstanceFullName
        sWbs = wb.name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        mMe.RenewLogResult() = "Passed"
    Else
        mMe.RenewLogResult(sRenewAction & vbLf & "Restoring references did not find any saved to restore" _
                          ) = "Passed"
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
    Dim fso As New FileSystemObject
    
    With Application
        If Not DevInstncWorkbookExists Then
            '~~ At this point the Development instance Workbook must no longer exist at its location
            .EnableEvents = False
            mMe.RenewLogAction = "Save the Addin instance Workbook '" & AddInInstanceName & "' (version " & AddInVersion & " as Development instance Workbook '" & DevInstncName & "'"
            
            On Error Resume Next
            wbAddIn.SaveAs DevInstncFullName, FileFormat:=xlDevlpFormat, ReadOnlyRecommended:=False
            
            If Not mCompMan.WbkIsOpen(io_name:=DevInstncName) _
            Then Stop _
            Else wbDevlp.VBProject.name = fso.GetBaseName(DevInstncName)
            
            If Err.Number <> 0 Then
                mMe.RenewLogResult("Saving the Addin instance Workbook '" & AddInInstanceName & "' (version " & AddInVersion & " as Development instance Workbook '" & DevInstncName & "' failed!" _
                                  ) = "Failed"
            Else
                mMe.RenewLogResult() = "Passed"
            End If
            .EnableEvents = True
        Else ' file still exists
            mMe.RenewLogResult("Saving the Addin instance Workbook '" & AddInInstanceName & "' (version " & AddInVersion & " as Development instance Workbook '" & DevInstncName & "' failed!" _
                              ) = "Failed"
        End If
    End With

xt: Exit Sub

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
            Application.Run AddInInstanceName & "!mCompMan.UpdateRawClones", ThisWorkbook, wbAddIn.HOSTED_RAWS
        End If
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

