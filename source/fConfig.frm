VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fConfig 
   Caption         =   "CompMan's basic configuration (user specific in Registry )"
   ClientHeight    =   7152
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   11655
   OleObjectBlob   =   "fConfig.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fso                         As New FileSystemObject
Private bFolderServicedIsValid      As Boolean
Private bFolderAddinIsValid         As Boolean
Private bFolderExportIsValid        As Boolean
Private bFolderSyncedIsValid        As Boolean
Private bAddinConfigObligatory      As Boolean
Private bSyncConfigObligatory       As Boolean
Private bCanceled                   As Boolean

Private Sub UserForm_Initialize()
    With Me
        .FolderAddin = mConfig.FolderAddin
        .FolderServiced = mConfig.ServicedDevAndTestFolder
        .FolderExport = Replace(Split(mConfig.FolderExport, ",")(UBound(Split(mConfig.FolderExport, ","))), "\", vbNullString)
        .FolderSynced = mConfig.ServicedSyncTargetFolder
        .Caption = "CompMan's basic configuration"
    End With
    VerifyConfig
End Sub

Private Sub VerifyConfig()
    FolderServicedVerify
    FolderAddinVerify
    FolderExportVerify
    FolderSyncedVerify
    If MayBeConfirmed Then Me.cmbConfirmed.Enabled = True
End Sub

Private Property Get MayBeConfirmed() As Boolean
    With Me
        Select Case True
            Case (bAddinConfigObligatory And FolderAddinIsValid) Or (.FolderAddin <> vbNullString And FolderAddinIsValid) _
                 And .FolderServicedIsValid _
                 And .FolderExportIsValid _
                 And (bSyncConfigObligatory And FolderSyncedIsValid) Or (.FolderSynced <> vbNullString And FolderSyncedIsValid)
                MayBeConfirmed = True
            Case Else
                MayBeConfirmed = False
        End Select
    End With
End Property

Public Property Let AddinConfigObligatory(ByVal b As Boolean):  bAddinConfigObligatory = b:                         End Property

Public Property Let SyncConfigObligatory(ByVal b As Boolean):   bSyncConfigObligatory = b:                          End Property

Public Property Get FolderAddin() As String:                    FolderAddin = Me.tbxFolderAddin.Text:               End Property

Public Property Let FolderAddin(ByVal s As String):             Me.tbxFolderAddin.Text = s:                         End Property

Private Property Let FolderAddinInfo( _
        Optional ByVal invalid As Boolean = False, _
                 ByVal s As String)
                 
    With Me.lblFolderAddinInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property

Public Property Get FolderAddinIsValid() As Boolean:            FolderAddinIsValid = bFolderAddinIsValid:           End Property

Private Property Let FolderAddinIsValid(ByVal b As Boolean):    bFolderAddinIsValid = b:                            End Property

Public Property Get FolderExport() As String:                   FolderExport = Me.tbxFolderExport.Text:             End Property

Public Property Let FolderExport(ByVal s As String):            Me.tbxFolderExport.Text = s:                        End Property

Private Property Let FolderExportInfo(Optional ByVal invalid As Boolean = False, _
                                               ByVal s As String)
                 
    With Me.lblFolderExportInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property

Public Property Get FolderExportIsValid() As Boolean:           FolderExportIsValid = bFolderExportIsValid:         End Property

Private Property Let FolderExportIsValid(ByVal b As Boolean):   bFolderExportIsValid = b:                           End Property

Public Property Get FolderServicedIsValid() As Boolean:         FolderServicedIsValid = bFolderServicedIsValid:     End Property

Private Property Let FolderServicedIsValid(ByVal b As Boolean): bFolderServicedIsValid = b:                         End Property

Public Property Get FolderSyncedIsValid() As Boolean:           FolderSyncedIsValid = bFolderSyncedIsValid:         End Property

Private Property Let FolderSyncedIsValid(ByVal b As Boolean):   bFolderSyncedIsValid = b:                           End Property

Public Property Get FolderSynced() As String:                   FolderSynced = Me.tbxFolderSynced.Text:             End Property

Public Property Let FolderSynced(ByVal s As String):             Me.tbxFolderSynced.Text = s:                       End Property

Private Property Let FolderSyncedInfo(Optional ByVal invalid As Boolean = False, _
                                               ByVal s As String)
                 
    With Me.lblFolderSyncedInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property

Public Property Get FolderServiced() As String:                 FolderServiced = Me.tbxFolderServiced.Text:         End Property

Public Property Let FolderServiced(ByVal s As String):          Me.tbxFolderServiced.Text = s:                      End Property

Private Property Let FolderServicedInfo(Optional ByVal invalid As Boolean = False, _
                                                 ByVal s As String)
    With Me.lblFolderServicedInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property

Private Sub FolderAddinVerify()
' ----------------------------------------------------------------------------
' Verifies the current configured Addin-Folder and sets
' FolderAddinIsValid = True if ok.
' ----------------------------------------------------------------------------
    
    With Me
        If .FolderAddin = vbNullString Then
            .FolderAddin = Application.AltStartupPath
            If .FolderAddin = vbNullString Then
                FolderAddinInfo(True) = "Not yet configured! " & vbLf & _
                                        "Folder for the CompMan Addin instance and application data. The folder is (or becomes) identical with the " & _
                                        "Application.AltStartupPath."
            Else
                FolderAddinInfo = "Folder for the CompMan Addin (which is the current Application.AltStartupPath in use." & vbLf & _
                                  "Attention: When another folder is selected this one will become the new Application.AltStartupPath - " & _
                                  "any items in the current folder will no longer be considered by Excel at Startup!"
                FolderAddinIsValid = True
            End If
        ElseIf Not fso.FolderExists(.FolderAddin) Then
            FolderAddinInfo(True) = "Invalid! (the folder does not exist)"
        ElseIf mConfig.FolderAddin <> vbNullString Then
            If StrComp(.FolderAddin, mConfig.FolderAddin, vbTextCompare) <> 0 Then
                FolderAddinInfo = "The folder for the CompMan Addin instance is about to change from '" & mConfig.FolderAddin & "' " & _
                                  "to '" & .FolderAddin & "'!"
            Else
                FolderAddinInfo = "Folder for the CompMan Addin instance. Defaults to the current Application.AltStartupPath " & _
                                  "if one is already specified. I.e. any selected folder becomes the Application.AltStartupPath. " & _
                                  "Take into account when altered!"

            End If
            FolderAddinIsValid = True
        Else
            FolderAddinIsValid = True
        End If
    End With

End Sub

Private Sub cmbFolderAddin_Click()
    Dim s As String
    With Me
        s = mBasic.SelectFolder("Select the folder for the 'CompMan-Addin-Instance'")
        If s <> vbNullString Then
            .FolderAddin = s
        End If
    End With
    VerifyConfig

End Sub

Private Sub cmbFolderSynced_Click()
    Dim s As String
    With Me
        s = mBasic.SelectFolder("Select the folder for CompMan's 'Synchronization' service")
        If s <> vbNullString Then
            .FolderSynced = s
        End If
    End With
    VerifyConfig

End Sub

Private Sub cmbCancel_Click()
    bCanceled = True
    Me.Hide
End Sub

Private Sub cmbConfirmed_Click()
' ----------------------------------------------------------------------------
' The confirmed and verified configuration is written to the CompMan.cfg file
' ----------------------------------------------------------------------------
    Me.Hide
End Sub

Public Property Get Canceled() As Boolean
    Canceled = bCanceled
End Property

Private Sub cmbFolderServiced_Click()
    Dim s As String
    With Me
        s = mBasic.SelectFolder("Select the obligatory 'CompMan-Serviced-Root-Folder'")
        If s <> vbNullString Then .FolderServiced = s
    End With
    VerifyConfig
End Sub

Private Sub FolderExportVerify()
' ----------------------------------------------------------------------------
' Verification the name of the Export-Folder
' ----------------------------------------------------------------------------
    With Me
        If .FolderExport <> vbNullString Then
            If mConfig.FolderExport <> vbNullString _
            And StrComp(.FolderExport, mConfig.FolderExport, vbTextCompare) <> 0 Then
                FolderExportInfo = "The name of the Export Folder will be changed from '" & mConfig.FolderExport & "' " & _
                                   "to '" & .FolderExport & "' along with the next export of changed components."
                FolderExportIsValid = True
            Else
                FolderExportInfo = "Current configured name of the folder within the Workbook folder into which CompMan exports " & _
                                   "new/modified components. The name may be changed however."
                FolderExportIsValid = True
            End If
        Else
            FolderExportInfo = "A name for the Export Folder within the Workbook folder into which CompMan exports " & _
                               "new/modified components is obligatory!"
            FolderExportIsValid = False
        End If
    End With
End Sub

Private Sub FolderServicedVerify()
' ----------------------------------------------------------------------------
' Verification of the current Serviced-Root-Folder
' ----------------------------------------------------------------------------
    With Me
        FolderServicedIsValid = False
        
        If .FolderServiced = vbNullString Then
            FolderServicedInfo(True) = "Not yet configured! Folder in which a Workbook/Workbook-Folder must be located for being serviced by CompMan."
            .cmbConfirmed.Enabled = False
        ElseIf Not fso.FolderExists(.FolderServiced) Then
            FolderServicedInfo(True) = "Invalid! (folder does not exist)"
            .cmbConfirmed.Enabled = False
        ElseIf mConfig.FolderExport <> vbNullString _
            And StrComp(.FolderServiced, mConfig.ServicedDevAndTestFolder, vbTextCompare) <> 0 Then
            .cbxMoveServicedRoot.Visible = True
            FolderServicedIsValid = True
            FolderServicedInfo = "The serviced root folder is about to chasnge from '" & mConfig.ServicedDevAndTestFolder & "' " & _
                                     "to '" & .FolderServiced & "'! Moving all content to new folder may be appropriate."
            .cmbConfirmed.Enabled = True
        Else
            .cbxMoveServicedRoot.Visible = False
            FolderServicedIsValid = True
            FolderServicedInfo = "Folder in which Workbook/Workbook folder must be located for being supported by CompMan"
            .cmbConfirmed.Enabled = True
        End If
    End With

End Sub

Private Sub FolderSyncedVerify()
' ----------------------------------------------------------------------------
' Verification of the current configured 'Synchronization-Folder'
' ----------------------------------------------------------------------------
    With Me
        FolderSyncedIsValid = False
        
        If .FolderSynced = vbNullString Then
            FolderSyncedInfo(True) = "Not yet configured! Folder in which a Workbook/Workbook-Folder " & _
                                     "must be located for being considered by CompMan's 'Synchronization' service."
            .cmbConfirmed.Enabled = True
        ElseIf Not fso.FolderExists(.FolderSynced) Then
            FolderSyncedInfo(True) = "Invalid! (the currently configured 'Synchronization' folder does not exist)"
            .cmbConfirmed.Enabled = False
        ElseIf .FolderServiced <> vbNullString Then ' a 'serviced' folder is already configured
            If InStr(.FolderSynced, .FolderServiced) <> 0 Then
                FolderSyncedIsValid = False
                FolderSyncedInfo(True) = "The folder for CompMan's 'Synchronization' service must not be a sub-folder " & _
                                         "of the 'Serviced' folder! (both may have the same root folder however)"
                .cmbConfirmed.Enabled = False
            Else
                .cmbConfirmed.Enabled = True
                FolderSyncedInfo = "Current configured folder. Any Workbook opened from within this folder will be considered " & _
                                   "by CompMan's 'Synchronization' service - provided all required preconditions are met."
                FolderSyncedIsValid = True
            End If
        Else
            .cmbConfirmed.Enabled = True
            FolderSyncedIsValid = True
        End If
    End With

End Sub

Private Sub tbxFolderExport_AfterUpdate()
    With tbxFolderExport
        If Len(.Text) = 0 _
        Then .Text = "source" _
        Else .Text = Split(.Text, vbLf)(0) ' ensure single line entry
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

