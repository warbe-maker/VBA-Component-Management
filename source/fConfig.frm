VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fConfig 
   Caption         =   "UserForm1"
   ClientHeight    =   5520
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   11295
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
Private bServicedRootFolderIsValid  As Boolean
Private bAddinFolderIsValid         As Boolean
Private bExportFolderIsValid        As Boolean

Private Sub cmbCancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub UserForm_Initialize()
    With Me
        .AddinFolder = mConfig.CompManAddinFolder
        .ServicedRootFolder = mConfig.CompManServicedRootFolder
        .ExportFolder = mConfig.CompManExportFolder
    End With
    ServicedRootFolderVerify
    AddinFolderVerify
    ExportFolderVerify
End Sub

Public Property Get MoveServicedRootFolder() As Boolean
    MoveServicedRootFolder = Me.cbxMoveServicedRoot.Value
End Property

Public Property Get MoveAddinFolder() As Boolean
    MoveAddinFolder = Me.cbxMoveAddin.Value
End Property

Public Property Get AddinFolder() As String
    AddinFolder = Me.tbxAddinFolder.Text
End Property

Public Property Let AddinFolder(ByVal s As String):             Me.tbxAddinFolder.Text = s:                 End Property

Private Property Let AddinFolderInfo( _
        Optional ByVal invalid As Boolean = False, _
                 ByVal s As String)
                 
    With Me.lblAddinFolderInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property

Public Property Get AddinFolderIsValid() As Boolean:            AddinFolderIsValid = bAddinFolderIsValid:   End Property

Private Property Let AddinFolderIsValid(ByVal b As Boolean):    bAddinFolderIsValid = b:                    End Property

Public Property Get ExportFolder() As String:                   ExportFolder = Me.tbxExportFolder.Text:     End Property

Public Property Let ExportFolder(ByVal s As String):            Me.tbxExportFolder.Text = s:                End Property

Public Property Get ExportFolderIsValid() As Boolean:           ExportFolderIsValid = bExportFolderIsValid: End Property

Private Property Let ExportFolderIsValid(ByVal b As Boolean):   bExportFolderIsValid = b:                   End Property

Public Property Get ServicedRootFolder() As String
    ServicedRootFolder = Me.tbxServicedRootFolder.Text
End Property

Public Property Let ServicedRootFolder(ByVal s As String):      Me.tbxServicedRootFolder.Text = s:          End Property

Private Property Let ServicedRootFolderInfo( _
        Optional ByVal invalid As Boolean = False, _
                 ByVal s As String)
    With Me.lblServicedRootFolderInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property

Public Property Get ServicedRootFolderIsValid() As Boolean
    ServicedRootFolderIsValid = bServicedRootFolderIsValid
End Property

Private Property Let ServicedRootFolderIsValid(ByVal b As Boolean)
    bServicedRootFolderIsValid = b
End Property

Private Sub AddinFolderVerify()
' ----------------------------------------------------------------------------
' Verification of the current configured Addin-Folder
' ----------------------------------------------------------------------------
    
    With Me
        If .AddinFolder = vbNullString Then
            AddinFolderInfo(True) = "Not yet configured! " & vbLf & _
                              "Folder for the CompMan Addin instance which becomes the future " & _
                              "Application.AltStartupPath."
            .cmbConfirmed.Enabled = False
        ElseIf Not fso.FolderExists(.AddinFolder) Then
            AddinFolderInfo(True) = "Invalid! (the folder does not exist)"
            .cmbConfirmed.Enabled = False
        ElseIf mConfig.CompManAddinFolder <> vbNullString Then
            If StrComp(.AddinFolder, mConfig.CompManAddinFolder, vbTextCompare) <> 0 Then
                .cbxMoveAddin.Visible = True
                AddinFolderInfo = "The folder for the CompMan Addin instance is about to change from '" & mConfig.CompManAddinFolder & "' " & _
                                  "to '" & .AddinFolder & "'!"
            Else
                .cbxMoveAddin.Visible = False
                AddinFolderInfo = "Current configured folder for the CompMan Addin instance = the current configured " & _
                                  "Application.AltStartupPath. Attention when about to be altered!"
            End If
            AddinFolderIsValid = True
            .cmbConfirmed.Enabled = True
        Else
            AddinFolderIsValid = True
            .cmbConfirmed.Enabled = True
        End If
    End With

End Sub

Private Sub cmbAddinFolder_Click()
    Dim s As String
    With Me
        s = mBasic.SelectFolder("Select the folder for the 'CompMan-Addin-Instance'")
        If s <> vbNullString Then .AddinFolder = s
    End With
    AddinFolderVerify

End Sub

Private Sub cmbConfirmed_Click()
' ----------------------------------------------------------------------------
' The confirmed and verified configuration is written to the CompMan.cfg file
' ----------------------------------------------------------------------------
    Me.Hide
End Sub

Private Sub cmbServicedRootFolder_Click()
    Dim s As String
    With Me
        s = mBasic.SelectFolder("Select the obligatory 'CompMan-Serviced-Root-Folder'")
        If s <> vbNullString Then .ServicedRootFolder = s
    End With
    ServicedRootFolderVerify
End Sub

Private Property Let ExportFolderInfo( _
        Optional ByVal invalid As Boolean = False, _
                 ByVal s As String)
                 
    With Me.lblExportFolderInfo
        .Caption = s
        If s <> vbNullString Then .Visible = True Else .Visible = False
        If invalid Then .ForeColor = rgbRed Else .ForeColor = rgbBlack
    End With
End Property


Private Sub ExportFolderVerify()
' ----------------------------------------------------------------------------
' Verification the name of the Export-Folder
' ----------------------------------------------------------------------------
    With Me
        If .ExportFolder <> vbNullString Then
            If mConfig.CompManExportFolder <> vbNullString _
            And StrComp(.ExportFolder, mConfig.CompManExportFolder, vbTextCompare) <> 0 Then
                ExportFolderInfo = "The name of the Export Folder will be changed from '" & mConfig.CompManExportFolder & "' " & _
                                   "to '" & .ExportFolder & "' along with the next export of changed components."
                ExportFolderIsValid = True
            Else
                ExportFolderInfo = "Current configured name of the folder within the Workbook folder into which CompMan exports " & _
                                   "new/modified components. The name may be changed however."
                ExportFolderIsValid = True
            End If
        Else
            ExportFolderInfo = "A name for the Export Folder within the Workbook folder into which CompMan exports " & _
                               "new/modified components is obligatory!"
            ExportFolderIsValid = False
        End If
    End With
End Sub

Private Sub MoveFiles(ByVal mf_source_folder As String, _
                      ByVal mf_destination_folder As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim fl As File
 
    For Each fl In fso.GetFolder(mf_source_folder).Files
        fl.Move Destination:=mf_destination_folder
    Next fl
 
End Sub

Private Sub MoveFolders(ByVal mf_source_folder As String, _
                        ByVal mf_destination_folder As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
End Sub

Private Sub ServicedRootFolderVerify()
' ----------------------------------------------------------------------------
' Verification of the current Serviced-Root-Folder
' ----------------------------------------------------------------------------
    With Me
        ServicedRootFolderIsValid = False
        
        If .ServicedRootFolder = vbNullString Then
            ServicedRootFolderInfo(True) = "Not yet configured! Folder in which a Workbook/Workbook-Folder must be located for being serviced by CompMan."
            .cmbConfirmed.Enabled = False
        ElseIf Not fso.FolderExists(.ServicedRootFolder) Then
            ServicedRootFolderInfo(True) = "Invalid! (folder does not exist)"
            .cmbConfirmed.Enabled = False
        ElseIf mConfig.CompManExportFolder <> vbNullString _
            And StrComp(.ServicedRootFolder, mConfig.CompManServicedRootFolder, vbTextCompare) <> 0 Then
            .cbxMoveServicedRoot.Visible = True
            ServicedRootFolderIsValid = True
            ServicedRootFolderInfo = "The serviced root folder is about to chasnge from '" & mConfig.CompManServicedRootFolder & "' " & _
                                     "to '" & .ServicedRootFolder & "'! Moving all content to new folder may be appropriate."
            .cmbConfirmed.Enabled = True
        Else
            .cbxMoveServicedRoot.Visible = False
            ServicedRootFolderIsValid = True
            ServicedRootFolderInfo = "Folder in which Workbook/Workbook folder must be located for being supported by CompMan"
            .cmbConfirmed.Enabled = True
        End If
    End With

End Sub

Private Sub tbxExportFolder_AfterUpdate()
    With tbxExportFolder
        If Len(.Text) = 0 _
        Then .Text = "source" _
        Else .Text = Split(.Text, vbLf)(0) ' ensure single line entry
    End With
End Sub

Private Sub TextBoxSingleLine(ByRef tbx As MsForms.TextBox)
    With tbx
         If Len(.Text) = 0 Then Exit Sub
         .Text = Split(.Text, vbLf)(0)
    End With
End Sub

