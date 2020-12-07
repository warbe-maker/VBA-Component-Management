Attribute VB_Name = "mAddIn"
Option Explicit
Option Private Module
' ---------------------------------------------------------------------------
' Standard Module mAddIn
'          Methods/procedures specifically for the Addin Workbook.
'
' Methods (public):
' - ControlItemRenewAdd    Used when the development and test instance
'                          Workbook is opened to add a Renew control
'                          item to the "Add-Ins" poupup menu.
' - ControlItemRenewRemove Used when the development and test instance
'                          Workbook is closed.
' - Renew                  Called via the Renew control item in the
'                          "Add-Ins" popup menu or executed via the
'                          VBE when the code of the Addin had been modified
'                          in its Development instance Workbook.
'                          + -------------------------------------------------- +
' - SaveAsDev              || Saves the Addin Workbook as Development instance ||
'                          || Workbook. Used exclusively through the immediate ||
'                          || Window when the code in the Addin has been       ||
'                          || modified directly - instead of through the       ||
'                          || Development instance Workbook                    ||
'                          + -------------------------------------------------- +
' - UpdateUsedCommComps    Exclusively used by the Development instance
'                          Workbook to update its own used Common Components
'                          of which the origin code had changed - via the
'                          Addin - provided it is open - by means of the
'                          the corresponding method.
'
' W. Rauschenberger, Berlin Nov 2020
' ---------------------------------------------------------------------------
Const CONTROL_CAPTION = "Renew Addin"

Public cllLog           As Collection                   ' log records of the Renew and/or SaveAsDev process
Private wbDevlp         As Workbook
Private wbSource        As Workbook                     ' This development instance as the renew source
Private wbTarget        As Workbook                     ' The Addin instance as renew target
Private bSucceeded      As Boolean
Private bAllRemoved     As Boolean
Private dctAddInRefs    As Dictionary
Private lStep           As Long

Private Enum enSaveRestore
    enSaveAndOff
    enRestore
End Enum

Private Property Get Step() As Long
    If lStep = 0 Then lStep = 1 Else lStep = lStep + 1
    Step = lStep
End Property

Private Sub AddinInstanceWorkbookClose()
    
    Dim lStep As Long

    On Error Resume Next
    wbSource.Activate
    wbTarget.Close False
    lStep = Step
    Application.StatusBar = lStep & ". Close the Addin instance Workbook"

    If Err.Number <> 0 _
    Then cllLog.Add lStep & ". Close the Addin Workbook failed: unable to close!" & vbLf & _
                 "  (" & Err.Description & ")" _
    Else cllLog.Add lStep & ". Close Addin Workbook passed"

End Sub

Private Sub AddInInstanceWorkbookDelete()

    Dim lStep   As Long
    Dim fso     As New FileSystemObject

    On Error Resume Next
    lStep = Step
    Application.StatusBar = lStep & ". Delete the Addin instance Workbook"

    With fso
        If .FileExists(wbAddIn.AddInInstanceFullName) Then
            .DeleteFile wbAddIn.AddInInstanceFullName
            If Err.Number = 0 _
            Then cllLog.Add lStep & ". Delete Addin Workbook passed." _
            Else cllLog.Add lStep & ". Delete Addin Workbook file failed: " & vbLf & _
                         "  (" & Err.Description & ")"
        Else
            cllLog.Add lStep & ". Delete Addin Workbook passed: already deleted"
        End If
    End With
    
End Sub

Private Function AddInInstanceWorkbookExists() As Boolean
Dim fso As New FileSystemObject
    AddInInstanceWorkbookExists = fso.FileExists(wbAddIn.AddInInstanceFullName)
End Function

Private Sub AddinInstanceWorkbookOpen()

    Dim lStep           As Long
    Dim wb              As Workbook
    Dim sBaseAddinName  As String
    Dim sBaseDevName    As String
    Dim fso             As New FileSystemObject

    lStep = Step
    Application.StatusBar = lStep & ". Open the Addin instance Workbook"
    
    If Not mAddIn.AddInInstanceWorkbookIsOpen Then
        If AddInInstanceWorkbookExists Then
            On Error Resume Next
            Set wb = Application.Workbooks.Open(wbAddIn.AddInInstanceFullName)
            If Err.Number = 0 Then
                With fso
                    sBaseAddinName = .GetBaseName(wb.name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.name)
                    wb.VBProject.name = sBaseAddinName
                End With
                cllLog.Add lStep & ". (Re)open Addin Workbook passed: successfully (re)opened/(re)activated and renamed from " & sBaseDevName & " to " & sBaseAddinName
 _
            Else
                cllLog.Add lStep & ". (Re)open the Addin Workbook failed: could't be (re)opened." & _
                         vbLf & "  (" & Err.Description & ")"
            End If
        End If
    End If

End Sub

Private Sub AddinInstanceWorkbookSaveAsDevlp()

    Dim lStep   As Long
    Dim fso     As New FileSystemObject

    lStep = Step
    Application.StatusBar = lStep & ". Save the Addin instance Workbook as Development instance"
    
    With Application
        If Not DevlpInstanceWorkbookExists Then
            '~~ At this point the Development instance Workbook must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbAddIn.SaveAs wbAddIn.DevlpInstanceFullName, FileFormat:=wbAddIn.xlDevlpFormat, ReadOnlyRecommended:=False
            
            If Not mWrkbk.IsOpen(wbAddIn.DevlpInstanceName, wbDevlp) _
            Then Stop _
            Else wbDevlp.VBProject.name = fso.GetBaseName(wbAddIn.DevlpInstanceName)
            
            If Err.Number <> 0 Then
                cllLog.Add lStep & ". Save Addin instance (version " & wbAddIn.AddInVersion & " as Development instance Workbook failed!"
            Else
                cllLog.Add lStep & ". Save Addin instance Workbook (version " & wbAddIn.AddInVersion & " saved as Development instance Workbook passed."
            End If
            .EnableEvents = True
        Else ' file still exists
            cllLog.Add lStep & ". Saving Addin with version " & wbAddIn.AddInVersion & " as Development instance Workbook failed"
        End If
    End With

End Sub

Private Function AssertUpToDateVersion() As Boolean
' ---------------------------------------------------------------------------------------------------------------------
' Returns TRUE when the version of the Addin instance Workbook is identical with the delopment instance Workbook.
' The version of the Addin instance is obtained via the call of "mAddIn.Version" by Application.Run. Because
' ByRef parameter is not supported by this method the value is returned via the Version class object.
' See: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1022_ByRef_Argument_with_the_Application_Run_method.shtml
' ---------------------------------------------------------------------------------------------------------------------
    Const PROC = "AssertUpToDateVersion"
    
    On Error GoTo eh
    Dim lStep       As Long
    Dim cVersion    As clsAddinVersion

    If cllLog Is Nothing Then Set cllLog = New Collection
    Set cVersion = New clsAddinVersion
    lStep = Step
    Application.StatusBar = lStep & ". Assert the up-to-date version"
    
    Application.StatusBar = lStep & ". Assert the up-to-date version"
    If mAddIn.AddInInstanceWorkbookIsOpen Then
        Application.Run wbTarget.name & "!mCompMan.Version", cVersion
        If cVersion.Version = wbAddIn.AddInVersion Then
            cllLog.Add lStep & ". Renew of the Addin (version " & cVersion.Version & ") successfull! Addin is open again and active :-)"
            bSucceeded = True
        Else
            cllLog.Add lStep & ". Renew of the Addin failed (is still version " & cVersion.Version & ":-("
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

Public Sub ControlItemRenewAdd()
'------------------------------------
' Add control to "Add_Ins" popup menu
'------------------------------------
    Const PROC = "ControlItemRenewAdd"
    
    On Error GoTo eh
    Dim cmb   As CommandBar
    Dim cmbb  As CommandBarButton
    
    mAddIn.ControlItemRenewRemove
    Set cmb = Application.CommandBars("Worksheet Menu Bar")
    Set cmbb = cmb.Controls.Add(Type:=msoControlButton, id:=2950)
    With cmbb
        .caption = CONTROL_CAPTION
        .Style = msoButtonCaption
        .TooltipText = "Saves the development instance as Addin"
        .OnAction = "wbAddIn.Renew"
        .Visible = True
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ControlItemRenewRemove()
'--------------------------------------------------------------
' Remove the "Renew" control item from the "Add_Ins" popup menu
'--------------------------------------------------------------
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(CONTROL_CAPTION).Delete
End Sub

Private Sub DevlpInstanceWorkbookClose()
    
    Dim lStep As Long

    On Error Resume Next
    lStep = Step
    Application.StatusBar = lStep & ". Close the Development instance Workbook"

    wbDevlp.Activate
    wbDevlp.Close False
    If Err.Number <> 0 _
    Then cllLog.Add lStep & ". Close Development instance Workbook failed: Unable to close" & vbLf & _
                 "  (" & Err.Description & ")" _
    Else cllLog.Add lStep & ". Close Development instance Workbook passed"
End Sub

Private Sub DevlpInstanceWorkbookDelete()
    
    Dim lStep   As Long
    Dim fso     As New FileSystemObject
    
    On Error Resume Next
    lStep = Step
    Application.StatusBar = lStep & ". Delete the Development instance Workbook"
    
    With fso
        If .FileExists(wbAddIn.DevlpInstanceFullName) Then
            .DeleteFile wbAddIn.DevlpInstanceFullName
            If Err.Number = 0 _
            Then cllLog.Add lStep & ". Delete Development instance Workbook passed" _
            Else cllLog.Add lStep & ". Delete Development instance Workbook failed!" & vbLf & _
                         "  (" & Err.Description & ")"
        Else
            cllLog.Add lStep & ". Development instance Workbook already deleted"
        End If
    End With
End Sub

Private Function DevlpInstanceWorkbookExists() As Boolean
Dim fso As New FileSystemObject
    DevlpInstanceWorkbookExists = fso.FileExists(wbAddIn.DevlpInstanceFullName)
End Function

Private Function DevlpInstanceWorkbookIsOpen() As Boolean
    On Error Resume Next
    Set wbDevlp = Application.Workbooks(wbAddIn.DevlpInstanceName)
    DevlpInstanceWorkbookIsOpen = Err.Number = 0
End Function

Private Sub DevlpInstanceWorkbookSave()
    
    Dim lStep As Long

    lStep = Step
    Application.StatusBar = lStep & ". Save the Development instance Workbook"
    
    Set wbSource = Application.Workbooks(wbAddIn.DevlpInstanceName)
    Application.EnableEvents = False
    wbSource.Save
    Application.EnableEvents = True
    wbSource.Activate
    cllLog.Add lStep & ". Save Development instance Workbook passed"

End Sub

Private Sub DevlpInstanceWorkbookSaveAsAddin()

    Dim lStep   As Long

    lStep = Step
    Application.StatusBar = lStep & ". Save the Development instance Workbook as Addin"
    
    With Application
        If Not AddInInstanceWorkbookExists Then
            '~~ At this point the Addin must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbSource.SaveAs wbAddIn.AddInInstanceFullName, FileFormat:=wbAddIn.xlAddInFormat
            If Err.Number <> 0 Then
                cllLog.Add lStep & ". Save Development instance (version " & wbAddIn.AddInVersion & ") as Addin failed!"
            Else
                cllLog.Add lStep & ". Save Development instance (version " & wbAddIn.AddInVersion & ") as Addin passed"
            End If
            .EnableEvents = True
            mCompMan.ExportChangedComponents wbDevlp
        Else ' file still exists
            cllLog.Add lStep & ". Setup/renew of the Addin with version " & wbAddIn.AddInVersion & " of the development instance failed"
        End If
    End With
    
End Sub

Private Sub DisplayRenewResult()

    Dim sMsg        As String
    Dim i           As Long

    If cllLog.Count > 0 Then
        sMsg = cllLog.Item(1)
        For i = 2 To cllLog.Count
            sMsg = sMsg & vbLf & cllLog.Item(i)
        Next i
        If bSucceeded _
        Then mMsg.Box dsply_title:="Successful! The Addin '" & wbAddIn.AddInInstanceName & "' has been renewed by the development instance '" & wbAddIn.DevlpInstanceName & "' (see details below)" _
                    , dsply_msg:=sMsg _
                    , dsply_msg_monospaced:=True _
        Else mMsg.Box dsply_title:="Failed! Renewing the Addin " & wbAddIn.AddInInstanceName & " by the development instance failed (see details below)" _
                    , dsply_msg:=sMsg _
                    , dsply_msg_monospaced:=True
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
        Then mMsg.Box dsply_title:="Successful! The Addin '" & wbAddIn.AddInInstanceName & "' has been saved as Development instance Workbook '" & wbAddIn.DevlpInstanceName & "' (see details below)" _
                    , dsply_msg:=sMsg _
                    , dsply_msg_monospaced:=True _
        Else mMsg.Box dsply_title:="Failed! Saving the Addin " & wbAddIn.AddInInstanceName & " as Development instance Workbook failed (see details below)" _
                    , dsply_msg:=sMsg _
                    , dsply_msg_monospaced:=True
    End If
    Application.StatusBar = vbNullString
    lStep = 0
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mAddIn" & "." & sProc
End Function

Private Sub ReferencesToAddInRestore()
    Const PROC = "ReferencesToAddInRestore"
    
    Dim v               As Variant
    Dim wb              As Workbook
    Dim sWbs            As String
    Dim bOneRestored    As Boolean
    Dim lStep           As Long
    
    On Error GoTo eh
    lStep = Step
    Application.StatusBar = lStep & ". Restore references to the Addin for open Workbooks"
    
    For Each v In dctAddInRefs
        Set wb = v
        wb.VBProject.References.AddFromFile wbAddIn.AddInInstanceFullName
        sWbs = wb.name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        cllLog.Add lStep & ". Reference to the Addin in open Workbooks restored"
        cllLog.Add "   (" & sWbs & ")"
    Else
        cllLog.Add lStep & ". Reference to the Addin in none of the open Workbooks restored (none originally referred to the Addin)"
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function AddInInstanceWorkbookIsOpen() As Boolean
    Const PROC = "AddInInstanceWorkbookIsOpen"
    
    On Error GoTo eh
    Dim i As Long
    
    AddInInstanceWorkbookIsOpen = False
    For i = 1 To Application.AddIns2.Count
        If Application.AddIns2(i).name = wbAddIn.AddInInstanceName Then
            On Error Resume Next
            Set wbTarget = Application.Workbooks(wbAddIn.AddInInstanceName)
            AddInInstanceWorkbookIsOpen = Err.Number = 0
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

Private Function ReferencesToAddInSaveAndRemove() As Boolean
' ----------------------------------------------------------------
' - Allows the user to close any open Workbook which refers to the
'   Addin, which definitly hinders the Addin from being re-newed.
' - Returns TRUE when the user closed all open Workbboks.
' ----------------------------------------------------------------
    Const PROC = "Renew1_ReferencesToAddInSaveAndRemove"
    
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
    Application.StatusBar = lStep & ". Save and remove references to the Addin from open Workbooks"
    
    With Application
        Set dct = mWrkbk.Opened ' Returns a Dictionary with all open Workbooks in any application insance
        Set dctAddInRefs = New Dictionary
        For Each v In dct
            Set wb = dct.Item(v)
            
            For Each ref In wb.VBProject.References
                If InStr(ref.name, fso.GetBaseName(wbAddIn.AddInInstanceName)) <> 0 Then
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
        cllLog.Add lStep & ". Reference to the Addin from open Workbooks saved and removed"
        cllLog.Add "   (" & sWbs & ")"
    Else
        cllLog.Add lStep & ". Removing reference to the Addin from open Workbooks not required (none referred to the Addin)."
    End If

    ReferencesToAddInSaveAndRemove = bAllRemoved
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub Renew()
' -----------------------------------------------------------
' Renews the code of the Addin instance of this Workbook with
' this Workbook's code by displaying a detailed result of the
' whole Renew process.
' Note: It cannot be avoided that this procedure is available
'       also in the Addin instance. However, its execution is
'       limited to this Workbook's development instance.
' -----------------------------------------------------------
    Const PROC = "Renew"
    
    On Error GoTo eh
    Dim lStep   As Long

    mErH.BoP ErrSrc(PROC)
        
    Set cllLog = New Collection
    
    '~~ Assert ThisWorkbook is the development instance of the CompMan Addin
    lStep = Step
    If Not wbAddIn.IsDevlpInstance() Then
        cllLog.Add lStep & ". ""Renew"" not asserted: Execution of the ""Renew"" procedure failed because it had not been executed from within the Development instance Workbook (" & wbAddIn.DevlpInstanceName & ")!"
        GoTo xt
    Else
        cllLog.Add lStep & ". ""Renew asserted: Procedure is executed from within the Development instance Workbook (" & wbAddIn.DevlpInstanceName & ") asserted"
    End If
                     
    '~~ Get the current CompMan's base configuration confirmed or changed
    mCompManCfg.Confirm
                     
    '~~ Assert no Workbooks are open referring to the Addin
    ReferencesToAddInSaveAndRemove
    If Not bAllRemoved Then GoTo xt

    '~~ Assure the current code version has been saved
    '~~ Note: Unconditionally saving the Workbook does an incredible trick:
    '~~       The un-unstalled and IsAddin=False Workbook is released from the Application
    '~~       and no longer considered "used"
    DevlpInstanceWorkbookSave
    wbSource.Activate
          
    '~~ Attempt to turn Addin to "IsAddin=False", uninstall and close it
    If AddInInstanceWorkbookIsOpen Then
        Set_IsAddin_ToFalse wbTarget
        AddinInstanceWorkbookClose
    End If
    
    '~~ Attempt to delete the Addin Workbook file
    AddInInstanceWorkbookDelete
    
    '~~ Attempt to save the development instance as Addin
    DevlpInstanceWorkbookSaveAsAddin
    
    '~~ Saving the development instance as Addin may also open the Addin.
    '~~ So if not already open it is re-opened and thus re-activated
    AddinInstanceWorkbookOpen
    
    '~~ Assert the correct version has been renewed/re-opened
    AssertUpToDateVersion
    
    '~~ Re-instate references to the Addin which had been removed
    ReferencesToAddInRestore

xt: DisplayRenewResult
    mErH.EoP ErrSrc(PROC)
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
    
    Dim lStep   As Long

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    Set cllLog = New Collection
        
    '~~ Assert ThisWorkbook is the development instance of the CompMan Addin
    lStep = Step
    If Not wbAddIn.IsAddinInstance() Then
        cllLog.Add lStep & ". Save as Development instance failed: It had not been executed from within the Addin instance Workbook!"
        GoTo xt
    Else
        cllLog.Add lStep & ". Save as Development instance passed"
    End If
    Set wbSource = Application.Workbooks(wbAddIn.AddInInstanceName)
    
    '~~ Save and remove references to the Addin from any open Workbook referring to it
    ReferencesToAddInSaveAndRemove
    If Not bAllRemoved Then GoTo xt

    '~~ Attempt to close the Development instance Workbook when open
    If DevlpInstanceWorkbookIsOpen Then
        If MsgBox("The Development instance Workbook of the Addin is still open. Saving the Addin as Development instance will cause any code modifications get lost. Reply with Ok or Cancel", _
                  vbOKCancel, _
                  "Development instance Workbook still open") = vbOK Then
            DevlpInstanceWorkbookClose
        Else
            GoTo xt
        End If
    End If
    
    '~~ Attempt to delete the Addin Workbook file
    DevlpInstanceWorkbookDelete
    
    '~~ Turn the Workbook property "IsAddin" to "False"
    Set_IsAddin_ToFalse wb:=wbSource
    
    '~~ Attempt to save the development instance as Addin
    AddinInstanceWorkbookSaveAsDevlp
    
    '~~ Re-instate references to the Addin which had been removed
    ReferencesToAddInRestore
    bSucceeded = True

xt: DisplaySaveAsDevResult
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Set_IsAddin_ToFalse(ByVal wb As Workbook)
    
    Dim lStep As Long

    lStep = Step
    Application.StatusBar = lStep & ". Set the IsAddin property to False"
    If wb.IsAddin = True Then
        wb.IsAddin = False
        cllLog.Add lStep & ". " & DQUOTE & "IsAddin" & DQUOTE & " property of the Addin Workbook set to FALSE"
    Else
        cllLog.Add lStep & ". " & DQUOTE & "IsAddin" & DQUOTE & " property of the Addin Workbook set to FALSE (were already done)"
    End If
    
End Sub

Public Sub UpdateUsedCommComps()
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
    If wbAddIn.IsDevlpInstance Then
        If mAddIn.AddInInstanceWorkbookIsOpen Then
            Application.Run wbAddIn.AddInInstanceName & "!mCompMan.UpdateUsedCommCompsTheRawHasChanged", ThisWorkbook
        End If
    End If
End Sub

