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
' - SaveAsDev              Saves the Addin Workbook as Development
'                          instance Workbook. Used exclusively through the
'                          immediate Window when the code in the Addin has
'                          been modified directly - instead through the
'                          Development instance Workbook
' - UpdateUsedCommComps    Exclusively used by the Development instance
'                          Workbook to update its own used Common Components
'                          of which the origin code had changed - via the
'                          Addin - provided it is open - by means of the
'                          the corresponding method.
'
' W. Rauschenberger, Berlin April 2020
' ---------------------------------------------------------------------------
Const CONTROL_CAPTION   As String = "Renew Addin"
Public cllLog           As Collection                   ' log records of the Renew and/or SaveAsDev process
Private wbDevlp         As Workbook
Private wbSource        As Workbook                     ' This development instance as the renew source
Private wbTarget        As Workbook                     ' The Addin instance as renew target
Private bSucceeded      As Boolean
Private bAllRemoved     As Boolean
Private bOneClosed      As Boolean
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

Private Function AddInInstanceWorkbookExists() As Boolean
Dim fso As New FileSystemObject
    AddInInstanceWorkbookExists = fso.FileExists(wbAddIn.AddInInstanceFullName)
End Function

Private Function AddInInstanceWorkbookIsOpen() As Boolean
    On Error Resume Next
    Set wbTarget = Application.Workbooks(wbAddIn.AddInInstanceName)
    AddInInstanceWorkbookIsOpen = Err.Number = 0
End Function

Private Function AssertUpToDateVersion() As Boolean
' ---------------------------------------------------------------------------------------------------------------------
' Returns TRUE when the version of the Addin instance Workbook is identical with the delopment instance Workbook.
' The version of the Addin instance is obtained via the call of "mAddIn.Version" by Application.Run. Because
' ByRef parameter is not supported by this method the value is returned via the Version class object.
' See: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1022_ByRef_Argument_with_the_Application_Run_method.shtml
' ---------------------------------------------------------------------------------------------------------------------
Dim lStep       As Long
Dim sVersion    As String
Dim cVersion    As clsAddinVersion

    On Error GoTo on_error
    Set cVersion = New clsAddinVersion
    lStep = Step
    Application.StatusBar = lStep & ". Assert the up-to-date version"
    
    Application.StatusBar = lStep & ". Assert the up-to-date version"
    If mAddIn.AddInInstanceWorkbookIsOpen Then
        Application.Run wbTarget.Name & "!mCompMan.Version", cVersion
        If cVersion.Version = wbAddIn.AddInVersion Then
            cllLog.Add lStep & ". Renew of the Addin (version " & cVersion.Version & ") successfull! Addin is open again and active :-)"
            bSucceeded = True
        Else
            cllLog.Add lStep & ". Renew of the Addin failed (is still version " & cVersion.Version & ":-("
            bSucceeded = False
        End If
    End If
exit_proc:
    Exit Function
    
on_error:
    Debug.Print Err.Description: Stop: Resume
End Function

Private Sub AddinInstanceWorkbookClose()
Dim lStep As Long

    On Error Resume Next
    wbSource.Activate
    wbTarget.Close False
    lStep = Step
    Application.StatusBar = lStep & ". Close the Addin instance Workbook"

    If Err.Number <> 0 _
    Then cllLog.Add lStep & ". The Addin Workbook could't be closed" & vbLf & _
                 "  (" & Err.Description & ")" _
    Else cllLog.Add lStep & ". Addin Workbook closed"
End Sub

Private Sub DevlpInstanceWorkbookClose()
Dim lStep As Long

    On Error Resume Next
    lStep = Step
    Application.StatusBar = lStep & ". Close the Development instance Workbook"

    wbDevlp.Activate
    wbDevlp.Close False
    If Err.Number <> 0 _
    Then cllLog.Add lStep & ". The Development instance Workbook of the Addin could't be closed" & vbLf & _
                 "  (" & Err.Description & ")" _
    Else cllLog.Add lStep & ". Development instance Workbook closed"
End Sub

Public Sub ControlItemRenewAdd()
'------------------------------------
' Add control to "Add_Ins" popup menu
'------------------------------------
Dim cmb   As CommandBar
Dim cmbb  As CommandBarButton

    On Error GoTo on_error
    mAddIn.ControlItemRenewRemove
    
    Set cmb = Application.CommandBars("Worksheet Menu Bar")
    Set cmbb = cmb.Controls.Add(Type:=msoControlButton, ID:=2950)
    With cmbb
        .Caption = CONTROL_CAPTION
        .Style = msoButtonCaption
        .TooltipText = "Saves the development instance as Addin"
        .OnAction = "wbAddIn.Renew"
        .Visible = True
    End With
    
exit_proc:
    Exit Sub

on_error:
    Debug.Print Err.Description: Stop: Resume
End Sub

Public Sub ControlItemRenewRemove()
'--------------------------------------------------------------
' Remove the "Renew" control item from the "Add_Ins" popup menu
'--------------------------------------------------------------
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(CONTROL_CAPTION).Delete
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
            Then cllLog.Add lStep & ". Addin Workbook file deleted" _
            Else cllLog.Add lStep & ". Deleting the Addin Workbook file failed" & vbLf & _
                         "  (" & Err.Description & ")"
        Else
            cllLog.Add lStep & ". Addin Workbook file were already deleted"
        End If
    End With
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
            Then cllLog.Add lStep & ". Development instance Workbook file deleted" _
            Else cllLog.Add lStep & ". Deleting the Development instance Workbook file failed" & vbLf & _
                         "  (" & Err.Description & ")"
        Else
            cllLog.Add lStep & ". Development instance Workbook file were already deleted"
        End If
    End With
End Sub

Private Function DevlpInstanceWorkbookIsOpen() As Boolean
    On Error Resume Next
    Set wbDevlp = Application.Workbooks(wbAddIn.DevlpInstanceName)
    DevlpInstanceWorkbookIsOpen = Err.Number = 0
End Function

Private Function DevlpInstanceWorkbookExists() As Boolean
Dim fso As New FileSystemObject
    DevlpInstanceWorkbookExists = fso.FileExists(wbAddIn.DevlpInstanceFullName)
End Function

Private Sub DisplayRenewResult()
Dim sMsg    As String
Dim i       As Long

    If cllLog.Count > 0 Then
        sMsg = cllLog.Item(1)
        For i = 2 To cllLog.Count
            sMsg = sMsg & vbLf & cllLog.Item(i)
        Next i
        If bSucceeded Then
            mBasic.Msg sTitle:="Successful! The Addin '" & wbAddIn.AddInInstanceName & "' has been renewed by the development instance '" & wbAddIn.DevlpInstanceName & "' (see details below)", _
                       sMsgText:=sMsg, bFixed:=True, vReplies:=vbOKOnly
        Else
            mBasic.Msg sTitle:="Failed! Renewing the Addin " & wbAddIn.AddInInstanceName & " by the development instance failed (see details below)", _
                   sMsgText:=sMsg, bFixed:=True, vReplies:=vbOKOnly
        End If
    End If
    Application.StatusBar = vbNullString
    lStep = 0
End Sub

Private Sub DisplaySaveAsDevResult()
Dim sMsg    As String
Dim i       As Long

    If cllLog.Count > 0 Then
        sMsg = cllLog.Item(1)
        For i = 2 To cllLog.Count
            sMsg = sMsg & vbLf & cllLog.Item(i)
        Next i
        If bSucceeded Then
            mBasic.Msg sTitle:="Successful! The Addin '" & wbAddIn.AddInInstanceName & "' has been saved as Development instance Workbook '" & wbAddIn.DevlpInstanceName & "' (see details below)", _
                       sMsgText:=sMsg, bFixed:=True, vReplies:=vbOKOnly
        Else
            mBasic.Msg sTitle:="Failed! Saving the Addin " & wbAddIn.AddInInstanceName & " as Development instance Workbook failed (see details below)", _
                   sMsgText:=sMsg, bFixed:=True, vReplies:=vbOKOnly
        End If
    End If
    Application.StatusBar = vbNullString
    lStep = 0
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mAddIn" & "." & sProc
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
                    sBaseAddinName = .GetBaseName(wb.Name)
                    sBaseDevName = .GetBaseName(ThisWorkbook.Name)
                    wb.VBProject.Name = sBaseAddinName
                End With
                cllLog.Add lStep & ". Addin Workbook successfully (re)opened/(re)activated and renamed from " & sBaseDevName & " to " & sBaseAddinName
 _
            Else
                cllLog.Add lStep & ". Addin Workbook could't be (re)opened." & _
                         vbLf & "  (" & Err.Description & ")"
            End If
        End If
    End If

End Sub

Private Function ReferencesToAddInSaveAndRemove() As Boolean
' ----------------------------------------------------------------
' - Allows the user to close any open Workbook which refers to the
'   Addin, which definitly hinders the Addin from being re-newed.
' - Returns TRUE when the user closed all open Workbboks.
' ----------------------------------------------------------------
Const PROC      As String = "Renew1_ReferencesToAddInSaveAndRemove"
Dim lStep       As Long
Dim dct         As Dictionary
Dim v           As Variant
Dim wb          As Workbook
Dim ref         As Reference
Dim sWbs        As String
Dim bOneRemoved As Boolean
Dim fso         As New FileSystemObject

    On Error GoTo on_error
    lStep = Step
    Application.StatusBar = lStep & ". Save and remove references to the Addin from open Workbooks"
    
    With Application
        Set dct = mWrkbk.Opened ' Returns a Dictionary with all open Workbooks in any application insance
        Set dctAddInRefs = New Dictionary
        For Each v In dct
            Set wb = dct.Item(v)
            
            For Each ref In wb.VBProject.References
                If InStr(ref.Name, fso.GetBaseName(wbAddIn.AddInInstanceName)) <> 0 Then
                    dctAddInRefs.Add wb, ref
                    sWbs = wb.Name & ", " & sWbs
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
    
exit_proc:
    Exit Function
    
on_error:
#If Debugging = 1 Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Public Sub Renew()
' ---------------------------------------------------------
' Renews the code of Addin instance Workbook with the code
' of the development instance Workbook.
' Even though the procedure is - by nature - available in
' the Addin performance of it is limited to the development
' instance.
' The result of the whole Renew process is displayed in
' detail.
' ---------------------------------------------------------
Const PROC  As String = "Renew"
Dim lStep   As Long

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    Set cllLog = New Collection
    
    '~~ Assert ThisWorkbook is the development instance of the CompMan Addin
    lStep = Step
    If Not wbAddIn.IsDevlpInstance() Then
        cllLog.Add lStep & ". Execution of the ""Renew"" procedure failed because it had not been executed from within the Development instance Workbook (" & wbAddIn.DevlpInstanceName & ")!"
        GoTo exit_proc
    Else
        cllLog.Add lStep & ". Execution of the ""Renew"" procedure from within the Development instance Workbook (" & wbAddIn.DevlpInstanceName & ") asserted"
    End If
                     
    '~~ Assert no Workbooks are open referring to the Addin
    ReferencesToAddInSaveAndRemove
    If Not bAllRemoved Then GoTo exit_proc

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

exit_proc:
    DisplayRenewResult
    
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub ReferencesToAddInRestore()
Const PROC          As String = "ReferencesToAddInRestore"
Dim v               As Variant
Dim wb              As Workbook
Dim ref             As Reference
Dim sWbs            As String
Dim bOneRestored    As Boolean
Dim lStep           As Long

    On Error GoTo on_error
    lStep = Step
    Application.StatusBar = lStep & ". Restore references to the Addin for open Workbooks"
    
    For Each v In dctAddInRefs
        Set wb = v
        wb.VBProject.References.AddFromFile wbAddIn.AddInInstanceFullName
        sWbs = wb.Name & ", " & sWbs
        bOneRestored = True
    Next v
    
    If bOneRestored Then
        sWbs = Left(sWbs, Len(sWbs) - 2)
        cllLog.Add lStep & ". Reference to the Addin in open Workbooks restored"
        cllLog.Add "   (" & sWbs & ")"
    Else
        cllLog.Add lStep & ". Reference to the Addin in none of the open Workbooks restored (none originally referred to the Addin)"
    End If
    
exit_proc:
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl

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
Const PROC  As String = "SaveAsDev"
Dim lStep   As Long

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    Set cllLog = New Collection
        
    '~~ Assert ThisWorkbook is the development instance of the CompMan Addin
    lStep = Step
    If Not wbAddIn.IsAddinInstance() Then
        cllLog.Add lStep & ". Execution of the ""SaveAsDev"" procedure failed because it had not been executed from within the Addin instance Workbook!"
        GoTo exit_proc
    Else
        cllLog.Add lStep & ". Execution of the ""SaveAsDev"" procedure from within the Addin instance Workbook asserted"
    End If
    Set wbSource = Application.Workbooks(wbAddIn.AddInInstanceName)
    
    '~~ Save and remove references to the Addin from any open Workbook referring to it
    ReferencesToAddInSaveAndRemove
    If Not bAllRemoved Then GoTo exit_proc

    '~~ Attempt to close the Development instance Workbook when open
    If DevlpInstanceWorkbookIsOpen Then
        If MsgBox("The Development instance Workbook of the Addin is still open. Saving the Addin as Development instance will cause any code modifications get lost. Reply with Ok or Cancel", _
                  vbOKCancel, _
                  "Development instance Workbook still open") = vbOK Then
            DevlpInstanceWorkbookClose
        Else
            GoTo exit_proc
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

exit_proc:
    DisplaySaveAsDevResult
    
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub AddinInstanceWorkbookSaveAsDevlp()
Dim lStep   As Long
Dim sFolder As String
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
            Else wbDevlp.VBProject.Name = fso.GetBaseName(wbAddIn.DevlpInstanceName)
            
            If Err.Number <> 0 Then
                cllLog.Add lStep & ". Addin instance (version " & wbAddIn.AddInVersion & " could not be saved as Development instance Workbook"
            Else
                cllLog.Add lStep & ". Addin instance Workbook (version " & wbAddIn.AddInVersion & " saved as Development instance Workbook"
            End If
            .EnableEvents = True
        Else ' file still exists
            cllLog.Add lStep & ". Saving Addin with version " & wbAddIn.AddInVersion & " as Development instance Workbook failed"
        End If
    End With
End Sub

Private Sub DevlpInstanceWorkbookSave()
Dim lStep As Long

    lStep = Step
    Application.StatusBar = lStep & ". Save the Development instance Workbook"
    
    Set wbSource = Application.Workbooks(wbAddIn.DevlpInstanceName)
    Application.EnableEvents = False
    wbSource.Save
    Application.EnableEvents = True
    wbSource.Activate
    cllLog.Add lStep & ". Development instance Workbook of the Addin saved"

End Sub

Private Sub DevlpInstanceWorkbookSaveAsAddin()
Dim lStep   As Long
Dim sFolder As String

    lStep = Step
    Application.StatusBar = lStep & ". Save the Development instance Workbook as Addin"
    
    With Application
        If Not AddInInstanceWorkbookExists Then
            '~~ At this point the Addin must no longer exist at its location
            .EnableEvents = False
            On Error Resume Next
            wbSource.SaveAs wbAddIn.AddInInstanceFullName, FileFormat:=wbAddIn.xlAddInFormat
            If Err.Number <> 0 Then
                cllLog.Add lStep & ". Development instance of the Addin (version " & wbAddIn.AddInVersion & " could not be saved as Addin"
            Else
                cllLog.Add lStep & ". Development instance of the Addin (version " & wbAddIn.AddInVersion & " saved as Addin"
            End If
            .EnableEvents = True
            mCompMan.ExportChangedComponents wbDevlp
        Else ' file still exists
            cllLog.Add lStep & ". Renewing the Addin with version " & wbAddIn.AddInVersion & " of the development instance failed"
        End If
    End With
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
            Application.Run wbAddIn.AddInInstanceName & "!mCompMan.UpdateUsedCommCompsTheOriginHasChanged", ThisWorkbook
        End If
    End If
End Sub
