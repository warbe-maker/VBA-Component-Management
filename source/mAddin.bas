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

Private Sub AutoOpenShortCutRemove()
    Dim fso As New FileSystemObject
    Dim s   As String
    
    s = AutoOpenShortCut
    If fso.FileExists(s) Then fso.DeleteFile s
    Set fso = Nothing
    
End Sub

Public Property Get Folder() As String:                        Folder = mWsh.Value(wsConfig, "FolderAddin"):                    End Property

Public Property Let Folder(ByVal s As String):                 mWsh.Value(wsConfig, "FolderAddin") = s:                         End Property

Public Property Get FolderOld() As String:                     FolderOld = mWsh.Value(wsConfig, "FolderAddinOld"):              End Property

Public Property Let FolderOld(ByVal s As String):               mWsh.Value(wsConfig, "FolderAddinOld") = s:                     End Property

Public Property Get IsSetup() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the CompMan-AddIn is configured and exists in the
' configured folder.
' ----------------------------------------------------------------------------
    With New FileSystemObject
        If mAddin.Folder <> vbNullString _
        Then IsSetup = .FileExists(mAddin.Folder & "\" & mAddin.WbkName)
    End With

End Property

Public Property Get Paused() As Boolean
    If mReg.Exists(PAUSED_REG_KEY, PAUSED_REG_VALUE_NAME) _
    Then Paused = CBool(mReg.Value(PAUSED_REG_KEY, PAUSED_REG_VALUE_NAME)) _
    Else Paused = False
End Property

Public Property Let Paused(ByVal b As Boolean):                mReg.Value(PAUSED_REG_KEY, PAUSED_REG_VALUE_NAME) = b:  End Property

Public Property Get WbkFullName() As String
    WbkFullName = mAddin.Folder & DBSLASH & WbkName
End Property

Public Property Get WbkName() As String
    With New FileSystemObject
        WbkName = .GetBaseName(ThisWorkbook.FullName) & "." & ADDIN_WORKBOOK_EXTENSION
    End With
End Property

Public Sub AddinFlipStatus()
    If mMe.IsDevInstnc Then
        If mAddin.Paused Then
            mAddin.Paused = False
        Else
            mAddin.Paused = True
        End If
        wsConfig.CurrentStatus
    End If
End Sub

Public Sub Clear(ByVal c_addin_folder_full_name As String)
' ----------------------------------------------------------------------------
' Clears all items concerning the Add-in despite the Add-in folder.
' ----------------------------------------------------------------------------

    mAddin.ReferencesRemove
    mAddin.SetIsAddinToFalse
    mAddin.WbkClose
    mAddin.WbkRemove c_addin_folder_full_name & "\" & mAddin.WbkName
    mAddin.AutoOpenShortCutRemove
    
End Sub

Public Function WbkClose(Optional ByRef wc_err_desc As String) As Boolean
    Dim wbk As Workbook
    
    If mAddin.IsOpen(wbk) Then
        Application.EnableEvents = False
        On Error Resume Next
        wbk.Close False
        If Err.Number <> 0 Then
            wc_err_desc = Err.Description
            WbkClose = False
            GoTo xt
        End If
    End If
    WbkClose = True ' not open, already closed

xt: Application.EnableEvents = True
    Exit Function

End Function

Public Function SetIsAddinToFalse() As Boolean
    Dim wbk As Workbook
    If mAddin.IsOpen(wbk) Then
        If wbk.IsAddin = True Then
            SetIsAddinToFalse = True
            wbk.IsAddin = False
        End If
    End If
End Function

Public Sub ReferencesRemove(Optional ByRef rr_dct As Dictionary, _
                            Optional ByRef rr_wbks As String, _
                            Optional ByRef rr_removed_one As Boolean, _
                            Optional ByRef rr_removed_all As Boolean)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
        
    Dim fso As New FileSystemObject
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
                If InStr(ref.name, fso.GetBaseName(mAddin.WbkName)) <> 0 Then
                    rr_dct.Add wbk, ref
                    rr_wbks = wbk.name & ", " & rr_wbks
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
    Set fso = Nothing
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mAddin" & "." & sProc
End Function

Public Function Exists() As Boolean
    Dim fso As New FileSystemObject
    Exists = fso.FileExists(mAddin.WbkFullName)
    Set fso = Nothing
End Function

Public Function IsOpen(Optional ByRef io_wbk As Workbook) As Boolean
    Const PROC = "IsOpen"
    
    On Error GoTo eh
    Dim i As Long
    
    For i = 1 To Application.AddIns2.Count
        If Application.AddIns2(i).name = mAddin.WbkName Then
            On Error Resume Next
            Set io_wbk = Application.Workbooks(mAddin.WbkName)
            IsOpen = Err.Number = 0
            GoTo xt
        End If
    Next i
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkExists() As Boolean
    Dim fso As New FileSystemObject
    WbkExists = fso.FileExists(WbkFullName)
    Set fso = Nothing
End Function
    
Public Function WbkRemove(ByVal wr_wbk_full_name As String) As Boolean
    Dim fso As New FileSystemObject
    If fso.FileExists(wr_wbk_full_name) Then fso.DeleteFile wr_wbk_full_name
    Set fso = Nothing
End Function

