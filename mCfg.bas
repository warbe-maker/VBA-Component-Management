Attribute VB_Name = "mCfg"
Option Explicit
Option Private Module

Private Const SECTION_BASE_CONFIG   As String = "BaseConfiguration"
Private Const VB_DEV_PROJECTS_ROOT  As String = "VBDevProjectsRoot"
Private Const COMPMAN_ADDIN_PATH    As String = "CompManAddInPath"

Private Property Get CFG_CHANGE_ADDIN_PATH() As String
    CFG_CHANGE_ADDIN_PATH = "Change CompMan" & vbLf & "AddIn Path"
End Property

Private Property Get CFG_CHANGE_DEVELOPMENT_ROOT() As String
    CFG_CHANGE_DEVELOPMENT_ROOT = "Change development" & vbLf & "root folder"
End Property

Private Property Get CFG_FILE_NAME() As String: CFG_FILE_NAME = ThisWorkbook.PATH & "\CompMan.cfg": End Property

Public Property Get CompManAddinPath() As String
    CompManAddinPath = value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMPMAN_ADDIN_PATH)
End Property

Public Property Let CompManAddinPath(ByVal s As String)
    value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMPMAN_ADDIN_PATH) = s
End Property

Public Property Get VBProjectsDevRoot() As String
    VBProjectsDevRoot = value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=VB_DEV_PROJECTS_ROOT)
End Property

Public Property Let VBProjectsDevRoot(ByVal s As String)
    value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=VB_DEV_PROJECTS_ROOT) = s
End Property

Public Function Asserted() As Boolean
' ----------------------------------------------------
' Assert that an existing Common folder is configured
' and that it contains a subfolder "CommComponents".
' Attention! This function must not run in the AddIn
' instance of this Workbook!
' ----------------------------------------------------
    Const PROC = "Assert"
    
    On Error GoTo eh
    Dim sPathCompMan As String
    
    With New FileSystemObject
        If .FolderExists(mCfg.CompManAddinPath) _
        And .FolderExists(mCfg.VBProjectsDevRoot) _
        Then
            Asserted = True
            .CopyFile Source:=CFG_FILE_NAME, Destination:=mCfg.CompManAddinPath & "\CompMan.cfg", OverWriteFiles:=True
            GoTo xt
        End If
                
        '~~ Assert the folder for the CompMan AddIn
        sPathCompMan = mCfg.CompManAddinPath
        If sPathCompMan = vbNullString Then
            sPathCompMan = mBasic.SelectFolder( _
                           sTitle:="Select the folder for the AddIn instance of the CompManDev Workbook (escape to use the Application.UserLibraryPath)")
            If sPathCompMan = vbNullString Then
                sPathCompMan = Application.UserLibraryPath ' Default because user escaped the selection
            Else
                '~~ Assure trust in this location and save it to the CompMan.cfg file
                mCfg.TrustThisFolder FolderPath:=sPathCompMan
                mCfg.CompManAddinPath = sPathCompMan
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
                    mCfg.TrustThisFolder FolderPath:=sPathCompMan
                    mCfg.CompManAddinPath = sPathCompMan
                End If
            Wend
        End If
                   
        '~~ Assert the root for VB-Development-Projects
        If mCfg.VBProjectsDevRoot = vbNullString Then
            mCfg.VBProjectsDevRoot = mBasic.SelectFolder(sTitle:="Select the root folder for any VB development/maintenance project which is to be supported by CompMan")
        End If
    
    End With

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub Confirm()
    Const PROC = "Confirm"
    Const CFG_CONFIRMED = "Confirmed"
    
    On Error GoTo eh
    Dim sMsg            As tMsg
    Dim sReply          As String
    
    With sMsg
        .section(1).sLabel = "Location (path) for the CompMan AddIn (the AddIn instance of the CompManDev Workbook):"
        .section(1).sText = mCfg.CompManAddinPath
        .section(1).bMonspaced = True
        .section(2).sLabel = "Root folder for any VB-Project in status development/maintenance about to be supported by CompMan:"
        .section(2).sText = mCfg.VBProjectsDevRoot
        .section(2).bMonspaced = True
    End With
    
    While sReply <> CFG_CONFIRMED
        sReply = mMsg.Dsply(msg_title:="Confirm or change the current Component Management's basic configuration" _
                          , msg:=sMsg _
                          , msg_buttons:=mMsg.Buttons(CFG_CONFIRMED, vbLf, CFG_CHANGE_ADDIN_PATH, CFG_CHANGE_DEVELOPMENT_ROOT) _
                           )
        Select Case sReply
            Case CFG_CHANGE_ADDIN_PATH
                Do
                    mCfg.CompManAddinPath = mBasic.SelectFolder("Select the ""obligatory!"" folder for the AddIn instance of the CompManDev Workbook")
                    If mCfg.CompManAddinPath <> vbNullString Then Exit Do
                Loop
                mCfg.TrustThisFolder FolderPath:=mCfg.CompManAddinPath
            
            Case CFG_CHANGE_DEVELOPMENT_ROOT
                Do
                    mCfg.VBProjectsDevRoot = mBasic.SelectFolder("Select the ""obligatory!"" root folder for VB development projects about to be supported by CompMan")
                    If mCfg.VBProjectsDevRoot <> vbNullString Then Exit Do
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.name & ": mCfg." & sProc
End Function

Public Sub TrustThisFolder(Optional ByVal FolderPath As String, _
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

    bSubFolders = True
    bNetworkLocation = False

    With New FileSystemObject
        If FolderPath = "" Then
            FolderPath = .GetSpecialFolder(2).PATH
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

exit_sub:
    Set oRegistry = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Property Get value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
    
    value = mFile.value(vl_file:=CFG_FILE_NAME _
                      , vl_section:=vl_section _
                      , vl_value_name:=vl_value_name _
                       )
End Property

Private Property Let value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
    mFile.value(vl_file:=CFG_FILE_NAME _
              , vl_section:=vl_section _
              , vl_value_name:=vl_value_name _
               ) = vl_value
End Property

