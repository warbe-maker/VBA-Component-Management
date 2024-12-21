Attribute VB_Name = "mConfig"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mConfig: Services for the initial setup of the default
' ======================== environment and the maintenance of the current
'                          setup.
' Public Properties:
' ------------------
' - CompManParentFolderNameDefault
' - ServicedRootFolderNameCurrent
'
' Public services:
' ----------------
' - Adjust              Adjusts the configuration in the "Config" Worksheet
'                       when the serviced root folder has been moved to
'                       another location and/or has been renamed.
' - DefaultEnvDisplay   Displays the to-be-setup default environment.
' - EnvIsMissing        Returns TRUE when the current ThisWorkbook parent
'                       folder has no Addin and ThisWorkbooks parent.parent
'                       folder has no Common-Components folder.
' - DefaultEnvConfirmed Confirms the setup of the default environment.
' - SelfSetupDefaultEnvironment     Sets up CompMan's default environment of files and
'                       folders
' ----------------------------------------------------------------------------
Public Const DEFAULT_FOLDER_COMPMAN_PARENT     As String = "CompMan"
Public Const DEFAULT_FOLDER_COMMON_COMPONENTS  As String = "Common-Components"
Public Const DEFAULT_FOLDER_COMPMAN_ROOT       As String = "CompManServiced"
Public Const DEFAULT_FOLDER_EXPORT             As String = "source"

Private Property Get CommonCompsFolderNameCurrent() As String:      CommonCompsFolderNameCurrent = ServicedRootFolderNameCurrent & "\" & DEFAULT_FOLDER_COMMON_COMPONENTS:                  End Property

Private Property Get CommonCompsFolderNameDefault() As String:      CommonCompsFolderNameDefault = ServicedRootFolderNameDefault & "\" & DEFAULT_FOLDER_COMMON_COMPONENTS:                  End Property

Private Property Get CompManParentFolderNameCurrent() As String:    CompManParentFolderNameCurrent = FSo.GetFile(ThisWorkbook.FullName).ParentFolder:                                       End Property

Public Property Get CompManParentFolderNameDefault() As String:     CompManParentFolderNameDefault = ServicedRootFolderNameDefault & "\" & DEFAULT_FOLDER_COMPMAN_PARENT:                   End Property

Private Property Get ExportFolderNameCurrent() As String:           ExportFolderNameCurrent = CompManParentFolderNameCurrent & "\" & wsConfig.FolderExport:                                 End Property

Private Property Get ExportFolderNameDefault() As String:           ExportFolderNameDefault = CompManParentFolderNameDefault & "\" & DEFAULT_FOLDER_EXPORT:                                 End Property

Public Property Get ServicedRootFolderNameCurrent() As String:      ServicedRootFolderNameCurrent = FSo.GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder:                           End Property

Private Property Get ServicedRootFolderNameDefault() As String:     ServicedRootFolderNameDefault = FSo.GetFile(ThisWorkbook.FullName).ParentFolder & "\" & DEFAULT_FOLDER_COMPMAN_ROOT:    End Property

Public Sub Adjust()
' ----------------------------------------------------------------------------
' Adjusts the configuration in the "Config" Worksheet when the serviced root
' folder has been moved to another location and/or has been renamed.
' ----------------------------------------------------------------------------
    With wsConfig
        .FolderCompManServicedRoot = ServicedRootFolderNameCurrent          ' adjust the root path and
        .FolderCommonComponentsPath = CommonCompsFolderNameCurrent  ' adjust the common components folder
        If .AutoOpenAddinIsSetup Then .AutoOpenAddinSetup           ' re-setup a setup Addin auto-open
        If .AutoOpenCompManIsSetup Then .AutoOpenCompManSetup       ' re-setup a setup CompMan.xlsb auto-open
    End With
    
End Sub

Public Function DefaultEnvDisplay(ByVal d_bttn_goahead As String, _
                                  ByVal d_bttn_abort As String) As Variant
' ----------------------------------------------------------------------------
' Displays the to-be-setup default environment.
' ----------------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim lMax    As Long
    Dim Msg     As mMsg.udtMsg
    Dim a       As Variant
    
    mBasic.Arry(a) = FSo.GetFolder(ThisWorkbook.Path).Name
    mBasic.Arry(a) = " +--" & DEFAULT_FOLDER_COMMON_COMPONENTS
    mBasic.Arry(a) = " |  +--CompManClient.bas"
    mBasic.Arry(a) = " | "
    mBasic.Arry(a) = " +--" & FSo.GetBaseName(ThisWorkbook.Name)
    mBasic.Arry(a) = " |   +--" & ThisWorkbook.Name
    mBasic.Arry(a) = " |  +--" & wsConfig.FolderExport
    mBasic.Arry(a) = " |"
    mBasic.Arry(a) = " +--WinMerge.ini"
    lMax = Max(a) + 5
    
    With Msg
        With .Section(1).Text
            .Text = "The " & ThisWorkbook.Name & " Workbook has apparently been opened for the very first time at this location " & _
                    "which is regarded the future serviced root folder." & vbLf & _
                    "CompMans ""self-setup"" is now about to setup the below default servicing environment at the current " & _
                    "location (see """ & Replace(d_bttn_goahead, vbLf, " ") & """. When set up CompMan Workbook may be configured for " & _
                    """auto-open"" or and/as  is Addin. The serviced root folder may then - and at any time - be moved " & _
                    "to any other location and/or be renamed. 1)"
        End With
        With .Section(2)
            .Label.Text = "Future servicing environment:"
            .Label.FontColor = rgbBlue
            .Text.MonoSpaced = True
            .Text.FontSize = 8
            .Text.Text = mBasic.AlignLeft(CStr(a(0)), lMax) & "serviced root folder, defaults to folder the Workbook had been opened in 1)" & vbLf & _
                         mBasic.AlignLeft(CStr(a(1)), lMax) & "fixed name for Common Components, not re-configurable" & vbLf & _
                         mBasic.AlignLeft(CStr(a(2)), lMax) & vbLf & _
                         mBasic.AlignLeft(CStr(a(3)), lMax) & vbLf & _
                         mBasic.AlignLeft(CStr(a(4)), lMax) & "setup finally moves the Workbook into its dedicated folder 2)" & vbLf & _
                         mBasic.AlignLeft(CStr(a(5)), lMax) & vbLf & _
                         mBasic.AlignLeft(CStr(a(6)), lMax) & "default name, may be re-configured 1)" & vbLf & _
                         mBasic.AlignLeft(CStr(a(7)), lMax) & vbLf & _
                         mBasic.AlignLeft(CStr(a(8)), lMax)
        End With
        With .Section(3)
            .Label.Text = "1)"
            .Label.FontColor = rgbBlue
            With .Text
                .Text = "See ""Configuration changes"" for more information"
                .OnClickAction = "https://github.com/warbe-maker/VBA-Component-Management/blob/master/SpecsAndUse.md#configuration-changes-compmans-config-worksheet"
                
            End With
        End With
        With .Section(4)
            .Label.Text = "2)"
            .Label.FontColor = rgbBlue
            With .Text
                .Text = "In order to allow modified Common Components are updated in the " & ThisWorkbook.Name & " itself " & _
                        "it is enabled for the update service - which can only be provided via the Addin instance when configured."
                
            End With
        End With
        With .Section(5)
            .Label.FontBold = True
            .Label.FontColor = rgbBlue
            .Label.Text = Replace(d_bttn_goahead, vbLf, " ")
            .Text.Text = "The above files and folders structure will be created and the opened Workbook will be saved into the" & vbLf & _
                         CompManParentFolderNameDefault & vbLf & _
                         "folder and closed. Re-opening it from within its new folder structure will finalize " & _
                         "CompMan's self setup process."
        End With
        With .Section(6)
            .Label.FontBold = True
            .Label.FontColor = rgbBlue
            .Label.Text = Replace(d_bttn_abort, vbLf, " ")
            .Text.Text = "The by default assumed ""serviced root folder"" - the folder the Workbook had been moved into after " & _
                         "download - is not the one intended. The Workbook needs to be moved to a different location (which " & _
                         "then may become ""CompMan's serviced root folder"") and the Wotkbook will be re-opened there for its self-setup."
        End With
    End With
    
    DefaultEnvDisplay = mMsg.Dsply(d_title:="CompMan's self setup (when opened for the very first time after download)" _
                                 , d_msg:=Msg _
                                 , d_buttons:=mMsg.Buttons(d_bttn_goahead, vbLf, d_bttn_abort) _
                                 , d_label_spec:="R90")

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mConfig" & "." & es_proc
End Function

Public Sub SetupConfirmed()
' ----------------------------------------------------------------------------
' Confirms the setup of the default environment.
' ----------------------------------------------------------------------------
    Dim Msg             As mMsg.udtMsg
    Dim sSetupLocation  As String
    Dim a               As Variant
    Dim aComment        As Variant
    Dim lMax            As Long
    
    mBasic.Arry(a) = mEnvironment.CompManServicedRootFolder:        mBasic.Arry(aComment) = "current ""CompMan serviced"" root folder 1)"
    mBasic.Arry(a) = " +--" & mEnvironment.CommCompsPath:           mBasic.Arry(aComment) = "Common Components folder, not re-configurable"
    mBasic.Arry(a) = " |  +--CommComps.dat":                        mBasic.Arry(aComment) = "Private Profile file documenting the Common Component's origin"
    mBasic.Arry(a) = " |  +--mCodingGuidLines":                     mBasic.Arry(aComment) = "Common Component hosted in CompMan"
    mBasic.Arry(a) = " |  +--mCodingRules":                         mBasic.Arry(aComment) = "Common Component hosted in CompMan"
    mBasic.Arry(a) = " |  +--mCompManClient.bas":                   mBasic.Arry(aComment) = "Common Component hosted in CompMan"
    mBasic.Arry(a) = " | ":                                         mBasic.Arry(aComment) = ""
    mBasic.Arry(a) = " +--" & FSo.GetBaseName(ThisWorkbook.Name):   mBasic.Arry(aComment) = "CompMan's own dedicated folder (setup and finally moved into it"
    mBasic.Arry(a) = " |  +--" & mEnvironment.CompManServiceFolder: mBasic.Arry(aComment) = "CompMan's service folder (as for all serviced Workbooks"
    mBasic.Arry(a) = " |  |  +--" & ConfigLocal.ExportFolderName:   mBasic.Arry(aComment) = "default Export-Folder name, may be re-configured 1)"
    mBasic.Arry(a) = " |  |  +--CommComps.dat                  ":   mBasic.Arry(aComment) = "Private Profile file documenting used/hosted Common Components"
    mBasic.Arry(a) = " |  |  +--CompMan.cfg":                       mBasic.Arry(aComment) = "CompMan's local configuration (used by new CompMan.xlsb versions)"
    mBasic.Arry(a) = " |  |  +--ExecTrace.log":                     mBasic.Arry(aComment) = "Execution trace log file"
    mBasic.Arry(a) = " |  +--" & ThisWorkbook.Name:                 mBasic.Arry(aComment) = ""
    mBasic.Arry(a) = " |":                                          mBasic.Arry(aComment) = ""
    mBasic.Arry(a) = " +--CompManAddin":                            mBasic.Arry(aComment) = "Folder for CompMan's Addin instance (when configured) 1)"
    mBasic.Arry(a) = " +--WinMerge.ini":                            mBasic.Arry(aComment) = "WimMerge configuration used to display code changef"
    lMax = Max(a) + 5

    sSetupLocation = FSo.GetFolder(ServicedRootFolderNameCurrent).ParentFolder.ParentFolder
    
    With Msg
        .Section(1).Label.Text = "CompMan's self-setup environment." & vbLf & _
                                 "CompMan is now ready for servicing any Workbook with an " & _
                                 "enabled service when the Workbook islocated in a dedicated " & _
                                 "folder in the serviced root folder."
        With .Section(2).Text
            .MonoSpaced = True
            .Text = mBasic.AlignLeft(CStr(a(0)), lMax) & aComment(0) & vbLf & _
                    mBasic.AlignLeft(CStr(a(1)), lMax) & aComment(1) & vbLf & _
                    mBasic.AlignLeft(CStr(a(2)), lMax) & aComment(2) & vbLf & _
                    mBasic.AlignLeft(CStr(a(3)), lMax) & aComment(3) & vbLf & _
                    mBasic.AlignLeft(CStr(a(4)), lMax) & aComment(4) & vbLf & _
                    mBasic.AlignLeft(CStr(a(5)), lMax) & aComment(5) & vbLf & _
                    mBasic.AlignLeft(CStr(a(6)), lMax) & aComment(6) & vbLf & _
                    mBasic.AlignLeft(CStr(a(7)), lMax) & aComment(7) & vbLf & _
                    mBasic.AlignLeft(CStr(a(8)), lMax) & aComment(8) & vbLf & _
                    mBasic.AlignLeft(CStr(a(9)), lMax) & aComment(9) & vbLf & _
                    mBasic.AlignLeft(CStr(a(10)), lMax) & aComment(10) & vbLf & _
                    mBasic.AlignLeft(CStr(a(11)), lMax) & aComment(11) & vbLf & _
                    mBasic.AlignLeft(CStr(a(12)), lMax) & aComment(12) & vbLf & _
                    mBasic.AlignLeft(CStr(a(13)), lMax) & aComment(13) & vbLf & _
                    mBasic.AlignLeft(CStr(a(14)), lMax) & aComment(14) & vbLf & _
                    mBasic.AlignLeft(CStr(a(15)), lMax) & aComment(15) & vbLf & _
                    mBasic.AlignLeft(CStr(a(16)), lMax) & aComment(16)
            .FontSize = 9
        End With
    End With
    mMsg.Dsply d_title:="Setup of CompMan's default environment completed!" _
             , d_msg:=Msg _
             , d_buttons:=vbOKOnly _
             , d_label_spec:="R90"
             
    
End Sub

Public Sub SelfSetupPublishHostedCommonComponents(ByVal s_hosted As String)
' ----------------------------------------------------------------------------
' Publishes all hosted Common Components by copying their Export-File to the
' Common-Components folder.
' ----------------------------------------------------------------------------

    Dim Comp    As clsComp
    Dim sComp   As String
    Dim sTarget As String
    Dim sName   As String
    Dim v       As Variant
    
    mEnvironment.Provide True
    Set CommonServiced = New clsCommonServiced
    Set CommonPublic = New clsCommonPublic
    
    For Each v In StringAsArray(s_hosted)
        sComp = v
        Set Comp = New clsComp
        With Comp
            .CompName = sComp
            .KindOfComp = enCompCommonHosted
            .Export
            .SetServicedProperties
            sName = FSo.GetFileName(.ExpFileFullName)
            sTarget = mEnvironment.CommCompsPath & "\" & sName
            FSo.CopyFile .ExpFileFullName, sTarget
            .SetPublicEqualServiced
        End With
        Set Comp = Nothing
    Next v
        
End Sub

Public Sub SelfSetupDefaultEnvironment(ByRef s_compman_fldr As String)
' ----------------------------------------------------------------------------
' Sets up CompMan's default environment of files and folders. This setup is
' based on the assumption that CompMan.xlsb had been opened from within the
' future serviced root folder.
' ----------------------------------------------------------------------------
    Const PROC = "SelfSetupDefaultEnvironment"
        
    On Error GoTo eh
    Dim sRoot       As String
    Dim sCommComps  As String
    
    sRoot = ThisWorkbook.Path
    sCommComps = sRoot & "\" & mEnvironment.FLDR_NAME_COMMON_COMPONENTS
    If Not FSo.FolderExists(sCommComps) Then FSo.CreateFolder sCommComps
    
    s_compman_fldr = sRoot & "\CompMan"
    If Not FSo.FolderExists(s_compman_fldr) Then FSo.CreateFolder s_compman_fldr
    
    With wsConfig
        .FolderCompManServicedRoot = sRoot
        .FolderCommonComponentsPath = sCommComps
        .AutoOpenCompManRemove
        .AutoOpenAddinRemove
        If Not .Verified Then .Activate
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

