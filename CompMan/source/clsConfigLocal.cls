VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfigLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsConfig: Maintains and provides configuration items in a
' ======================= Private Profile File called CompMan.cfg.
'
' Public Properties:
' ------------------
'
' Uses the Common VBA Private Prifile services (clsPrivProf)
'
' W. Rauschenberger Berlin, Nov 2024
' ---------------------------------------------------------------------------
Private Const VALUE_NAME_EXPORT_FOLDER_NAME         As String = "ExportFolderName"
Private sSectionName                                As String

Public PPFile                                       As New clsPrivProf

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
        
    On Error GoTo eh
    Dim sCurrent As String
    
    sSectionName = "CompManLocalConfig"
    With PPFile
        .FileName = mEnvironment.CompManLocalConfigFileFullName
        .FileHeader = FileHeader
        .FileFooter = mEnvironment.PrivateProfileFileFooter
    End With
    
    If Not fso.FileExists(mEnvironment.CompManLocalConfigFileFullName) Then
        '~~ A not existing configuration file indicates a very first time setup of CompMan.
        '~~ Thus, values from the wsConfig Worksheet are provided as initial defaults.
        ExportFolderName = wsConfig.FolderExport
    Else
        '~~ When a configuration file exists, current *) values are used to update those in the wsConfig Worksheet.
        '~~ I. e. any new CompMan.xlsb version (downloaded and replacing the existing Workbook) when re-opened will
        '~~ provide the individual local configuration in the wsConfig Worksheet - where the config may be changed.
        '~~ *) The configuration file keeps all history names in a string delimited by a ">".
        HistoryItems ExportFolderName, sCurrent
        If wsConfig.FolderExport <> sCurrent Then
            Application.EnableEvents = False
            wsConfig.FolderExport = sCurrent
            Application.EnableEvents = True
        End If
    End If


xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Class_Terminate()
    Set PPFile = Nothing
End Sub

Private Function FileHeader() As String
    FileHeader = "This Private Profile file keeps and provides the used ""local"" configuration. Only when the file does not exist       " & vbLf & _
                 "(i.e. when CompMan.xlsb is opened the very first time) default values are taken from CompMan's configuration Worksheet." & vbLf & _
                 "Once the file exists, any configuration change made via this Worksheet is saved here. When the Workbook is opened,     " & vbLf & _
                 "values are written back to the configuration Worksheet in order to display the current ""local"" configuration. This   " & vbLf & _
                 "ensures that any new version of the CompMan.xlsb Workbook downloaded and opened uses the ""local"" configuration."
End Function

Friend Property Let ExportFolderName(ByVal s As String):   Value(VALUE_NAME_EXPORT_FOLDER_NAME) = s:                End Property

Friend Property Get ExportFolderName() As String:          ExportFolderName = Value(VALUE_NAME_EXPORT_FOLDER_NAME): End Property


Public Property Get NameHskpngSection() As String:         NameHskpngSection = PPFile.NameHskpngSection:    End Property

Private Property Get Value(Optional ByVal v_value_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    Value = PPFile.Value(v_value_name, sSectionName)

End Property

Private Property Let Value(Optional ByVal v_value_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    PPFile.Value(v_value_name, sSectionName) = v_value
                                    
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommonServiced" & "." & sProc
End Function

