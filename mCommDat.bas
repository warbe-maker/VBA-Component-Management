Attribute VB_Name = "mCommDat"
Option Explicit
' ---------------------------------------------------------
' Standard Module mCommDat Maintains Common Component data.
' ---------------------------------------------------------
Private dat                                 As clsAppData
Private cfg                                 As clsAppData
                                                                                
Private Const SBJCT_FILE_NAME               As String = "\Components.dat"
Private Const SECTION_COMPONENT             As String = "Component_"
Private Const SECTION_HOST_WORKBOOK         As String = "HostWorkbook_"
                                    
Private sSection                            As String

Public Property Get CodeVersionAsOfDate( _
                    Optional ByVal sHostBaseName As String, _
                    Optional ByVal sComp As String) As Date
' ---------------------------------------------------------
' Returns a used Common Component's (sComp) code version as
' as-of-date of the origin code Export File.
' ---------------------------------------------------------
    Const PROC = "CodeVersionAsOfDate_Get"
    
    On Error GoTo eh
    Dim v   As Variant

    InitDat
    sSection = SectionComponent(sComp)
    dat.Aspect = sSection
    If dat.Exists(sName:=sHostBaseName & ValueNameAsOfUpdateDate(sSection)) Then
        v = dat.ValueGet(sHostBaseName & ValueNameAsOfUpdateDate(sSection))
        If v <> vbNullString Then
            CodeVersionAsOfDate = v
        End If
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Let CodeVersionAsOfDate( _
                    Optional ByVal sHostBaseName As String, _
                    Optional ByVal sComp As String, _
                             ByVal dt As Date)
' -----------------------------------------------------------
' Registers a used Common Component's (sComp) code version as
' as-of-date of the origin code Export File.
' -----------------------------------------------------------
    InitDat
    sSection = SectionComponent(sComp)
    dat.Aspect = sSection
    dat.ValueLet sHostBaseName & ValueNameAsOfUpdateDate(sSection), dt
End Property

Public Property Get CommCompExpFileFullName( _
                    Optional ByVal sComp As String) As String
' -----------------------------------------------------------
' Returns a Common Component's (sComp) Export File full name.
' -----------------------------------------------------------
    InitDat
    sSection = SectionComponent(sComp)
    dat.Aspect = sSection
    CommCompExpFileFullName = dat.ValueGet(ValueNameExportFile(sSection))
End Property

Public Property Let CommCompExpFileFullName( _
                    Optional ByVal sComp As String, _
                             ByVal sExportFile As String)
' -------------------------------------------------------
' Registers a Common Component's (sComp) Export File
' -------------------------------------------------------
    InitDat
    With New FileSystemObject
        If .FileExists(sExportFile) Then
            sSection = SectionComponent(sComp)
            dat.Aspect = sSection
            dat.ValueLet sName:=ValueNameExportFile(sSection), vValue:=sExportFile
        Else
            Stop ' sExportFile does not exist
        End If
    End With
End Property

Public Property Get CommCompHostWorkbookBaseName( _
                    Optional ByVal sComp As String) As String
' -----------------------------------------------------------
' Returns a Common Component's (sComp) Host BaseName.
' -----------------------------------------------------------
    InitDat
    sSection = SectionComponent(sComp)
    dat.Aspect = sSection
    CommCompHostWorkbookBaseName = dat.ValueGet(ValueNameHostBaseName(sSection))
End Property

Public Property Let CommCompHostWorkbookBaseName( _
                    Optional ByVal sComp As String, _
                             ByVal sWbBaseName As String)
' -----------------------------------------------------------------
' Registers Common Component's (sComp) Host BaseName (sWbBaseName).
' -----------------------------------------------------------------
    InitDat
    sSection = SectionComponent(sComp)
    dat.Aspect = sSection
    dat.ValueLet ValueNameHostBaseName(sSection), sWbBaseName
End Property

Public Property Get CommCompHostWorkbookFullName( _
                    Optional ByVal sComp As String) As String
' -----------------------------------------------------------
' Returns a Component's (sComp) Host Workbook FullName.
' -----------------------------------------------------------
Dim sWbHostBaseName As String

    InitDat
    With dat
        sSection = SectionComponent(sComp)
        If .Exists(sAspect:=sSection) Then
            
            .Aspect = sSection
            sWbHostBaseName = .ValueGet(ValueNameHostBaseName(sSection))
            
            sSection = SectionHostWorkbook(sWbHostBaseName)
            .Aspect = sSection
            CommCompHostWorkbookFullName = .ValueGet(ValueNameHostFullName(sSection))
        End If
    End With
    
End Property

Public Property Get CommCompsHostWorkbookFullName( _
                    Optional ByVal sHostBaseName As String) As String
' -------------------------------------------------------------------------
' Returns the FullName of a Host Workbook of one or more Common Components.
' -------------------------------------------------------------------------
    InitDat
    sSection = SectionHostWorkbook(sHostBaseName)
    dat.Aspect = sSection
    CommCompsHostWorkbookFullName = dat.ValueGet(ValueNameHostFullName(sSection))
End Property

Public Property Let CommCompsHostWorkbookFullName( _
                    Optional ByVal sHostBaseName As String, _
                             ByVal sHostFullName As String)
' ---------------------------------------------------------------------------
' Registers the FullName of a Host Workbook of one or more Common Components.
' ---------------------------------------------------------------------------
    InitDat
    sSection = SectionHostWorkbook(sHostBaseName)
    dat.Aspect = sSection
    dat.ValueLet ValueNameHostFullName(sSection), sHostFullName
End Property

Public Property Get HostWorkbookCodeBackUpExportFolder( _
           Optional ByVal sHostBaseName As String) As String
' -------------------------------------------------------------
' Get the ExportFolder for the Common Workbook (sHostBaseName).
' -------------------------------------------------------------
    
    Dim sExpFolder  As String

    InitDat
    sSection = SectionHostWorkbook(sHostBaseName)
    dat.Aspect = sSection
    sExpFolder = dat.ValueGet(ValueNameWbExpFolder(sSection))
    If sExpFolder = vbNullString Then
        sExpFolder = mConfig.CommonComponentsBasePath & "\" & sHostBaseName
        With New FileSystemObject
            If Not .FolderExists(sExpFolder) Then
                .CreateFolder sExpFolder
            End If
        End With
        dat.ValueLet ValueNameWbExpFolder(sSection), sExpFolder
    End If
    HostWorkbookCodeBackUpExportFolder = sExpFolder
    
End Property

Public Property Let HostWorkbookCodeBackUpExportFolder( _
                    Optional ByVal sHostBaseName As String, _
                             ByVal sWbExpFolder As String)
' -----------------------------------------------------------
' Write a Workbook section to the cfg-File with sHostBaseName
' as the section and its Export Folder as value.
' -----------------------------------------------------------
    InitDat
    With dat
        sSection = SectionHostWorkbook(sHostBaseName)
        .Aspect = sSection
        With New FileSystemObject
            If Not .FolderExists(sWbExpFolder) Then
                .CreateFolder sWbExpFolder
            End If
        End With
        .ValueLet ValueNameWbExpFolder(sSection), sWbExpFolder
    End With
End Property

Private Property Get SectionComponent(ByVal s As String):           SectionComponent = SECTION_COMPONENT & s:                       End Property

Private Property Get SectionHostWorkbook(ByVal s As String):        SectionHostWorkbook = SECTION_HOST_WORKBOOK & s:                End Property

Private Property Get ValueNameAsOfUpdateDate(ByVal s As String):    ValueNameAsOfUpdateDate = s & ".CurrentCodeVersionAsOfDate":    End Property

Private Property Get ValueNameExportFile(ByVal s As String):        ValueNameExportFile = s & ".ExportFile":                        End Property

Private Property Get ValueNameHostBaseName(ByVal s As String):      ValueNameHostBaseName = s & ".HostWorkbook":                    End Property

Private Property Get ValueNameHostFullName(ByVal s As String):      ValueNameHostFullName = s & ".FullName":                        End Property

Private Property Get ValueNameName(ByVal s As String):              ValueNameName = s & ".Name":                                    End Property

Private Property Get ValueNameType(ByVal s As String):              ValueNameType = s & ".Type":                                    End Property

Private Property Get ValueNameWbExpFolder(ByVal s As String):       ValueNameWbExpFolder = s & ".ExportFolder":                     End Property

Public Sub CommCompRemove(ByVal sComp As String)
    InitDat
    dat.AspectRemove sAspect:=sComp
End Sub

Public Function CommCompsMaxLenght() As Long
    Const PROC = "CommCompsMaxLenght"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lMax        As Long
    Dim dct         As Dictionary

    InitDat
    Set dct = dat.Aspects
    If dct.Count > 0 Then
        For Each v In dct
            If Left(v, Len(SECTION_COMPONENT)) = SECTION_COMPONENT Then
                lMax = Max(lMax, Len(v))
            End If
        Next v
    End If
    CommCompsMaxLenght = lMax
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub DisplayCfg()
    Const PROC = "DisplayCfg"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim v       As Variant
    Dim lMax    As Long
    Dim sMsg    As String
    Dim sName   As String

    InitDat
'    dat.ValuesDisplay
    dat.Aspect = vbNullString
    Set dct = dat.Values
    '~~ Get max Name length
    For Each v In dct
        lMax = mBasic.Max(Len(Split(v, ".")(1) & "." & Split(v, ".")(2)), lMax)
    Next v
    For Each v In dct
        sName = Split(v, ".")(1) & "." & Split(v, ".")(2)
        sMsg = sName & String(lMax - Len(sName), " ") & " = " & dct.Item(v) & vbLf & sMsg
    Next v
    mMsg.Box dsply_title:="Current content of " & dat.Subject & " (section.valuename = value)", _
             dsply_msg:=sMsg, _
             dsply_msg_monospaced:=True

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCommDat" & "." & sProc
End Function

                                   
Public Function HostWorkbooks() As Dictionary
' -------------------------------------------
' Returns a Dictionary wit all Workbooks
' which do host one or more Common Components
' -------------------------------------------
    Const PROC = "HostWorkbooks"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim sHostFullName As String

    InitDat
    Set dct = New Dictionary
    For Each v In dat.Aspects
        If mCommDat.IsHostWorkbook(v, sHostFullName) Then
            mDct.DctAdd add_dct:=dct, add_key:=v, add_item:=sHostFullName, add_seq:=seq_ascending
        End If
    Next v
    Set HostWorkbooks = dct

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Sub InitCompManCfg()
    Const PROC = "InitCompManCfg"
    
    On Error GoTo eh
    
    If cfg Is Nothing Then
        Set cfg = New clsAppData
        With cfg
            .Location = spp_File
            .Extension = ".cfg"
            ' Attention: This configuration file must exist in the Workbook's folder
            '            which for the AddIn instance of this Workbook is the configured Addin folder.
            '            Thus, when the AddIn instance is setup/renewed this cfg file must be copied
            '            to the AddIn folder
            .Subject = ThisWorkbook.Path & "\CompMan.cfg"
        End With
    End If

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub InitDat()
    Const PROC = "InitDat"
    
    On Error GoTo eh
    
    If dat Is Nothing Then
        Set dat = New clsAppData
        With dat
            .Location = spp_File
            .Extension = ".dat"
            .Subject = CommonComponentsBasePath & SBJCT_FILE_NAME
        End With
    End If

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function IsCommonComponent( _
                ByVal sComp As String, _
       Optional ByRef flExpFile As File, _
       Optional ByRef sWbHostFullName As String) As Boolean
' ---------------------------------------------------------
' - Returns TRUE when the Common Component (sComp) is known
'   in the Components.dat File
' - Returns the code-baxkup Export File as object and the
'   "hosting" Workbook's full name.
' --------------------------------------------------------------------
    Const PROC = "IsCommonComponent"
    
    On Error GoTo eh
    Dim sSection    As String
    Dim sExportFile As String
    Dim sWbHost     As String

    InitDat
    With dat
        sSection = SectionComponent(sComp)
        If Not .Aspects.Exists(sSection) Then Exit Function
        .Aspect = sSection
        sExportFile = .ValueGet(ValueNameExportFile(sSection))
        sWbHost = .ValueGet(ValueNameHostBaseName(sSection))
         With New FileSystemObject
             If .FileExists(sExportFile) Then
                 Set flExpFile = .GetFile(sExportFile)
                 IsCommonComponent = True
             End If
         End With
        sSection = SectionHostWorkbook(sWbHost)
        .Aspect = sSection
        sWbHostFullName = .ValueGet(ValueNameHostFullName(sSection))
    End With

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function
                                  
Public Function IsHostWorkbook(ByVal sHostBaseName As String, _
                        Optional ByRef sHostFullName As String) As Boolean
' -------------------------------------------------------------------
'
' -------------------------------------------------------------------
    Const PROC = "IsHostWorkbook"
    
    On Error GoTo eh
    Dim dctAspects  As Dictionary
    Dim sBaseName   As String
    Dim sSection    As String
    
    InitDat
    With dat
        Set dctAspects = .Aspects
        sBaseName = Replace$(sHostBaseName, SECTION_HOST_WORKBOOK, vbNullString)
        If dctAspects.Exists(SectionHostWorkbook(sBaseName)) Then
            sSection = SectionHostWorkbook(sBaseName)
            .Aspect = sSection
            sHostFullName = .ValueGet(ValueNameHostFullName(sSection))
            IsHostWorkbook = True
        End If
    End With
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Function KnownCommonComponents() As Dictionary
' ---------------------------------------------------
' - Returns the known Common Components, i.e. those
'   registered in the Components.dat File.
' ----------------------------------------------------
    Const PROC = "KnownCommonComponents"
    
    On Error GoTo eh
    Dim dct As New Dictionary
    Dim v   As Variant
    
    InitDat
    With dat
        For Each v In .Aspects
            If InStr(v, SECTION_COMPONENT) <> 0 Then
                dct.Add Replace(v, SECTION_COMPONENT, vbNullString), v
            End If
        Next v
    End With
    Set KnownCommonComponents = dct

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Function SourceUsesCommComp(ByVal sWbSourceFullName As String, _
                                   ByVal sComp As String) As Boolean
' ---------------------------------------------------------------------
' Returns TRUE when the Common Component (sComp) is installed in (is
' used by) the Workbook (sWbSourceFullName).
' ---------------------------------------------------------------------
    Const PROC = "SourceUsesCommComp"
    
    On Error GoTo eh
    Dim sHostBaseName   As String
    Dim sSection        As String
    
    sHostBaseName = mBasic.BaseName(sWbSourceFullName)
    InitDat
    With dat
        sSection = SectionComponent(sComp)
        If .Exists(sAspect:=sSection) Then
            .Aspect = sSection
            If .Exists(sName:=sHostBaseName & ValueNameAsOfUpdateDate(sSection)) Then
                SourceUsesCommComp = True
            End If
        End If
    End With
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub TrustThisFolder(Optional ByVal FolderPath As String, _
                           Optional ByVal TrustSubfolders As Boolean = True, _
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

exit_sub:
    Set oRegistry = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

