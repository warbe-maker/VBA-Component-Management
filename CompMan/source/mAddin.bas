Attribute VB_Name = "mAddin"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mAddin    Sercices abut CompMan's Add-in instance
'
' - IsOpen      Returns True when the 'AddIn Instance' Workbook is open
' - Folder      Get/Let the configured path for CompMan's 'AddIn Instance'
' - Exists      Returns TRUE when the 'Add-in Instace' Workbook exists
'
' ----------------------------------------------------------------------------
Private Const ADDIN_WORKBOOK_EXTENSION  As String = "xlam"  ' Extension may depend on Excel version
Private Const PAUSED_REG_KEY            As String = "HKCU\SOFTWARE\CompMan\Addin\"
Private Const PAUSED_REG_VALUE_NAME     As String = "Paused"

Public Property Get AutoOpenShortCut()
    AutoOpenShortCut = Environ$("APPDATA") & "\Microsoft\Excel\XLSTART\CompManAddin.lnk"
End Property

Public Property Get Paused() As Boolean
    If mReg.Exists(PAUSED_REG_KEY, PAUSED_REG_VALUE_NAME) _
    Then Paused = CBool(mReg.Value(PAUSED_REG_KEY, PAUSED_REG_VALUE_NAME)) _
    Else Paused = False
End Property

Public Property Let Paused(ByVal b As Boolean): mReg.Value(PAUSED_REG_KEY, PAUSED_REG_VALUE_NAME) = b:  End Property

Public Property Get WbkFullName() As String:    WbkFullName = mEnvironment.AddInFolderPath & DBSLASH & WbkName: End Property

Public Property Get WbkName() As String
    WbkName = fso.GetBaseName(ThisWorkbook.FullName) & "." & ADDIN_WORKBOOK_EXTENSION
End Property

Private Sub AutoOpenShortCutRemove()
    Dim s   As String
    
    s = AutoOpenShortCut
    With fso
        If .FileExists(s) Then .DeleteFile s
    End With
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mAddin" & "." & sProc
End Function

Public Sub GiveUp()
' ----------------------------------------------------------------------------
' Remove CompMan Addin
' ----------------------------------------------------------------------------
    Const PROC = "GiveUp"
    
    On Error GoTo eh
    mAddin.Set_IsAddin_ToFalse
    mAddin.WbkClose
    mAddin.WbkRemove WbkFullName
    AutoOpenShortCutRemove

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function IsOpen(Optional ByRef i_wbk As Workbook) As Boolean
    Const PROC = "IsOpen"
    
    On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    s = mAddin.WbkName
    For i = 1 To Application.AddIns2.Count
        If Application.AddIns2(i).Name = s Then
            On Error Resume Next
            Set i_wbk = Application.Workbooks(s)
            IsOpen = Err.Number = 0
            On Error GoTo eh
            GoTo xt
        End If
    Next i
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub PausedFlipFlop()

    Application.ScreenUpdating = False
    If Paused _
    Then Paused = False _
    Else Paused = True
    wsConfig.CurrentStatus
    
End Sub

Public Sub ReferencesRemove(Optional ByRef rr_dct As Dictionary, _
                            Optional ByRef rr_wbks As String, _
                            Optional ByRef rr_removed_one As Boolean, _
                            Optional ByRef rr_removed_all As Boolean)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As Dictionary
    Dim v   As Variant
    Dim ref As Reference
    Dim wbk As Workbook
    
    With Application
        Set dct = mWbk.Opened ' Returns a Dictionary with all open Workbooks in any application instance
        Set rr_dct = New Dictionary
        For Each v In dct
            Set wbk = dct.Item(v)
            For Each ref In wbk.VBProject.References
                If InStr(ref.Name, fso.GetBaseName(mAddin.WbkName)) <> 0 Then
                    rr_dct.Add wbk, ref
                    rr_wbks = wbk.Name & ", " & rr_wbks
                End If
            Next ref
        Next v
        
        For Each v In rr_dct
            Set wbk = v
            Set ref = rr_dct(v)
            wbk.VBProject.References.Remove ref
            rr_removed_one = True
        Next v
        rr_removed_all = True
    End With
    
End Sub

Public Sub SetupRenew()
' ----------------------------------------------------------------------------
' Sets up CompMan as Add-in or, when already set up, renews it.
' ----------------------------------------------------------------------------
    Const PROC = "SetupRenew"
    mCompMan.ServicedWrkbk = ThisWorkbook
    mEnvironment.Provide True, ErrSrc(PROC)
    mMe.Renew___AddIn
End Sub

Public Function Set_IsAddin_ToFalse() As Boolean
    Const PROC = "Set_IsAddin_ToFalse"
    
    On Error GoTo eh
    Dim wbk As Workbook
    If mAddin.IsOpen(wbk) Then
        If wbk.IsAddin = True Then
            Set_IsAddin_ToFalse = True
            wbk.IsAddin = False
        End If
    Else
        Set_IsAddin_ToFalse = True
    End If

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function WbkClose(Optional ByRef wc_err_desc As String) As Boolean
    Const PROC = "WbkClose"
    
    Dim wbk As Workbook
    
    mCompManClient.Events ErrSrc(PROC), False
    If mAddin.IsOpen(wbk) Then
        On Error Resume Next
        wbk.Close False
        If Err.Number <> 0 Then
            wc_err_desc = Err.Description
            WbkClose = False
            GoTo xt
        End If
    End If
    WbkClose = True ' not open, already closed

xt: mCompManClient.Events ErrSrc(PROC), True
    Exit Function

End Function
    
Public Function WbkRemove(ByVal wr_wbk_full_name As String) As Boolean
    If fso.FileExists(wr_wbk_full_name) Then fso.DeleteFile wr_wbk_full_name
End Function

Private Sub SaveAsNormalWorkbook()
    
    Dim wbkAddIn    As Workbook
    Dim wbkOrigin   As Workbook
    Dim sFullName   As String

    Application.EnableEvents = False
    Set wbkAddIn = ThisWorkbook                     ' or specify the name of the add-in workbook
    sFullName = DevInstncFullName                   ' Specify the path where you want to save the normal workbook
    Set wbkOrigin = wbkAddIn
    
    wbkAddIn.SaveAs FileName:=sFullName _
                  , FileFormat:=xlExcel12   ' Save the add-in workbook as a normal workbook
    
    wbkOrigin.Activate
    MsgBox "Workbook saved as a normal workbook at: " & sFullName
    Application.EnableEvents = True

End Sub

