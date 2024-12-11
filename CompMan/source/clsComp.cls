VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------------
' Class Module clsComp: Provides specific CompMan services for a VB-Project's
' ===================== component - by default of the serviced Workbook. The module
' provides instances of the class modult clsCode for: the current component'a code,
' an existing public code (when a public Common Component exists), a pending release
' code (when the component has a modification pending release), and the exported code
' (when the component's export file exists).
'
' Public (Friend) Properties:
' ---------------------------
' - Changed             True when the current code differs from the last exported code
'                       or when the component has yet not/never been exported.
' - CodeName
' - IsCompCommonPending True when the component is represented as an Export-File in the
'                       CommonPending folder
' - IsCommCompPublic  True when the component is represented as an an Export-File in the
'                       Common-Components folder.
' - CompName
' - Exists
' - ExpFile
' - ExpFileExt
' - ExpFileTemp
' - ExpFileFullName
' - ExpFileTempFullName
' - ExpFilePath
' - IsWorkbk
' - KindOfComp
' - MaxLenComp
' - MaxLenType
' -
' - Service
' - Sheet
' - SheetName
' - TypeString
' - VBComp
' - Wrkbk
' - WrkbkBaseName
' - WrlbkFullName
'
' Services (Methods):
' -------------------
' - CompType
' - Export
' - RemoteSheetExists
' - ReplaceRawWithClone
' - ReplaceRawWithCloneWhenConfirmed
'
' Note: Even when the component is not one of the current serviced Workbook it may
'       still be one which is a public Common Component or a pending release Common
'       Component and as such the public and the pending code would be provided.
'
' -----------------------------------------------------------------------------------
Private flExpFile           As File
Private lKindOfComp         As enKindOfComp
Private sCompName           As String       ' The class module's global component name
Private sExpFileExt         As String
Private sExpFileFullName    As String       ' Full filespec for the Export-File of CompName
Private sExpFilePath        As String       ' Folder for the Export-File of CompName
Private sTmpFolder          As String       ' Folder for a temporary Export File
Private sWrkbkBaseName      As String       ' Common Component host Workbook's base name
Private sWrkbkFullName      As String       ' Common Component host Workbook's full name
Private vbc                 As VBComponent  ' VBComponent managed by this Class Module
Private wbk                 As Workbook     ' The class' Workbook
Private sService            As String       ' Name of the 'Service' folder, i.e. the folder of the serviced Workbook
Private dctType             As Dictionary
Private lMaxLenComp         As Long
Private wsh                 As Worksheet

Public CodeExprtd           As clsCode
Public CodeCrrent           As clsCode
Public CodePublic           As clsCode
Public CodePnding           As clsCode

Private Sub Class_Initialize()
    If dctType Is Nothing Then Set dctType = New Dictionary Else dctType.RemoveAll
    dctType.Add vbext_ct_ActiveXDesigner, "ActiveX Designer"
    dctType.Add vbext_ct_ClassModule, "Class Module"
    dctType.Add vbext_ct_Document, "Document Module"
    dctType.Add vbext_ct_MSForm, "UserForm"
    dctType.Add vbext_ct_StdModule, "Standard Module"
    Wrkbk = ActiveWorkbook
    sWrkbkFullName = ActiveWorkbook.FullName
End Sub

Private Sub Class_Terminate()
    
    Set wbk = Nothing
    Set vbc = Nothing
    Set flExpFile = Nothing
    CleanUpTemps
    Set CodeCrrent = Nothing
    Set CodeExprtd = Nothing
    Set CodePublic = Nothing
End Sub

Friend Property Get CodeLines() As Dictionary
' ---------------------------------------------------------------------------
' Returns the content this VBComponent's Export-File (Me.ExpFileFullName) as
' Dictionary with the extra Module Information at the beginning and leading/
' trailing empty code lines removed.
' ---------------------------------------------------------------------------
    Const PROC = "CodeLines"

    On Error GoTo eh
    Dim dct As New Dictionary
    
    With New FileSystemObject
        Set dct = mFso.FileDict(Me.ExpFile)
    End With
    RemoveModuleInfo dct
    RemoveLeadingTrailingEmptyLines dct

xt: Set CodeLines = dct
    Set dct = Nothing
    Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get CompName() As String:           CompName = sCompName:       End Property

Friend Property Let CompName(ByVal s_comp As String)
' ------------------------------------------------------------------------------
' Provides: The properties:      CompName        (sCompName)
'                                ExpFileFullName (sExpFileFullName)
'                                VBComp          (vbc)
'                                ExpFileExt      (sExpFileExt)
'           The class instances: CodeCrrent
'                                CodeExprtd
'                                CodePnding
'                                CodePublic
' ------------------------------------------------------------------------------
    Const PROC = "CompName-Let"
    
    Dim sTempExport         As String
    Dim dctPublicExpFiles   As Dictionary
    
    Set CodeCrrent = Nothing:   Set CodeCrrent = New clsCode
    Set CodeExprtd = Nothing:   Set CodeExprtd = New clsCode
    Set CodePnding = Nothing:   Set CodePnding = New clsCode
    Set CodePublic = Nothing:   Set CodePublic = New clsCode
    
    If wbk Is Nothing _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Component Name assigned for a yet unknown Workbook!"
    
    Set dctPublicExpFiles = CommonPublic.ExportFiles
    
    sCompName = s_comp
    If Serviced.CompExists(sCompName) Then
        Set vbc = wbk.VBProject.VBComponents(s_comp)
        Select Case vbc.Type
            Case vbext_ct_StdModule:    sExpFileExt = ".bas"
            Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
            Case vbext_ct_MSForm:       sExpFileExt = ".frm"
            Case vbext_ct_Document:     sExpFileExt = ".cls"
        End Select
        sExpFileFullName = mEnvironment.ExportServiceFolderPath & "\" & sCompName & sExpFileExt
    
        '~~ Provide all local info for the serviced component
        If FSo.FileExists(sExpFileFullName) Then
            With CodeExprtd
                .CompName = sCompName
                .Source = sExpFileFullName
                .SourceDesc = mDiff.SourceServicedExport
            End With
        End If
            
        With CodeCrrent
            .CompName = sCompName
            .Source = Me.VBComp
            .SourceDesc = mDiff.SourceServicedCode
            If .IsNone Then Stop
        End With
        CleanUpTemps
    
        If dctPublicExpFiles.Exists(sCompName) Then
            '~~ The component is known/existing public or - in case of a new hosted component -
            '~~ has already been registered as a pending release component
            With CodePublic
                .CompName = sCompName
                .Source = dctPublicExpFiles(sCompName)
                .SourceDesc = mDiff.SourcePublic
            End With
        End If
            
        If FSo.FileExists(CommonPending.ExportFiles(sCompName)) Then
            With CodePnding
                .CompName = sCompName
                .Source = CommonPending.ExportFiles(sCompName)
                .SourceDesc = mDiff.SourcePending
            End With
        End If
    End If

End Property

Friend Property Get ExpFile() As File
       
    If flExpFile Is Nothing Then
        If FSo.FileExists(ExpFileFullName) Then
            Set flExpFile = FSo.GetFile(ExpFileFullName)
        Else
            '~~ When the component has yet no never been exported
            '~~ it will be along with the Export changed service
        End If
    End If
    
    Set ExpFile = flExpFile
    
End Property

Friend Property Let ExpFile(ByVal fl As File):              Set flExpFile = fl:                                                 End Property

Friend Property Get ExpFileExt() As String:                 ExpFileExt = sExpFileExt:                                           End Property

Friend Property Let KindOfComp(ByVal en As enKindOfComp):   lKindOfComp = en:                                                   End Property

Friend Property Get MaxLenComp() As Long:                   MaxLenComp = lMaxLenComp:                                           End Property

Friend Property Get PendingLastModAt() As String:           PendingLastModAt = CommonPending.LastModAt(sCompName):              End Property

Friend Property Get PendingLastModBy() As String:           PendingLastModBy = CommonPending.LastModBy(sCompName):              End Property

Friend Property Get PendingLastModIn() As String:           PendingLastModIn = CommonPending.LastModInWrkbkFullName(sCompName): End Property

Friend Property Get PendingLastModOn() As String:           PendingLastModOn = CommonPending.LastModOn(sCompName):              End Property

Friend Property Let PendingLastModOn(ByVal s As String):    CommonPending.LastModOn(sCompName) = s:                             End Property

Friend Property Get PendingKindOfComp() As String:          PendingKindOfComp = CommonPending.LastModKindOfComp(sCompName):     End Property

Friend Property Get PublicLastModAt() As String:            PublicLastModAt = CommonPublic.LastModAt(sCompName):                End Property

Friend Property Get PublicLastModBy() As String:            PublicLastModBy = CommonPublic.LastModBy(sCompName):                End Property

Friend Property Get PublicLastModIn() As String:            PublicLastModIn = CommonPublic.LastModInWrkbkFullName(sCompName):   End Property

Friend Property Get PublicLastModOn() As String:            PublicLastModOn = CommonPublic.LastModOn(sCompName):                End Property

Friend Property Get Service() As String
    If sService = vbNullString Then
        sService = FSo.GetFileName(Me.Wrkbk.Path)
    End If
    Service = sService
End Property

                         
Friend Property Get ServicedLastModAt(Optional ByVal l_comp As String = vbNullString) As String
    If l_comp = vbNullString Then l_comp = sCompName
    ServicedLastModAt = CommonServiced.LastModAt(l_comp)
End Property

Friend Property Let ServicedLastModAt(Optional ByVal l_comp As String = vbNullString, _
                                                  ByVal s As String)
    If l_comp = vbNullString Then l_comp = sCompName
    CommonServiced.LastModAt(l_comp) = s
End Property

Friend Property Get ServicedLastModBy() As String:              ServicedLastModBy = CommonServiced.LastModBy(sCompName):                    End Property

Friend Property Let ServicedLastModBy(ByVal s As String):       CommonServiced.LastModBy(sCompName) = s:                                    End Property

Friend Property Get ServicedLastModExpFile() As String:         ServicedLastModExpFile = CommonServiced.LastModExpFileOrigin(sCompName):    End Property

Friend Property Let ServicedLastModExpFile(ByVal s As String):  CommonServiced.LastModExpFileOrigin(sCompName) = s:                         End Property

Friend Property Get ServicedLastModKindOfComp() As String:      ServicedLastModKindOfComp = CommonServiced.KindOfComponent(sCompName):      End Property

Friend Property Get ServicedLastModIn() As String:              ServicedLastModIn = CommonServiced.LastModInWrkbkFullName(sCompName):       End Property

Friend Property Let ServicedLastModIn(ByVal s As String):       CommonServiced.LastModInWrkbkFullName(sCompName) = s:                       End Property

Friend Property Get ServicedLastModOn() As String:              ServicedLastModOn = CommonServiced.LastModOn(sCompName):                    End Property

Friend Property Let ServicedLastModOn(ByVal s As String):    CommonServiced.LastModOn(sCompName) = s:                               End Property

Friend Property Get Sheet() As Worksheet:                    Set Sheet = wsh:                                                   End Property

Friend Property Let Sheet(ByRef s_wsh As Worksheet):         wsh = s_wsh:                                                       End Property

Friend Property Get VBComp() As VBComponent:                 Set VBComp = wbk.VBProject.VBComponents(sCompName):                End Property

Friend Property Get Wrkbk() As Workbook:                     Set Wrkbk = wbk:                                                   End Property

Friend Property Let Wrkbk(ByVal w_wbk As Workbook)
' ------------------------------------------------------------------------------
' By default a class instance is initialized with the Serviced.W
' ------------------------------------------------------------------------------
    If Not w_wbk Is wbk Then
        '~~ The Workbook is not the one the class instance has been initialized with
        If Not sCompName = vbNullString Then
            '~~ A component's name has already been provided
            CompName = sCompName ' re-instantiates all class clsCode instances
        End If
    End If
    Set wbk = w_wbk
    
End Property

Friend Property Get WrkbkFullName() As String:                  WrkbkFullName = sWrkbkFullName:                                     End Property

Public Function Changed() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Component's Export-File does not exist or the
' Export-File's code differs from the current code of the component.
' ------------------------------------------------------------------------------
    Const PROC = "Changed"
    
    On Error GoTo eh
    Select Case True
        Case Not FSo.FileExists(Me.ExpFileFullName)     ' first time export
            Changed = True
        Case Else                                       ' export file exists
            If mDiff.ServicedExportVersusServicedCode(Me) Then
                Changed = True
            End If
    End Select
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub CleanUpTemps()
       
    With New FileSystemObject
        If .FolderExists(sTmpFolder) Then .DeleteFolder sTmpFolder
    End With
    
End Sub

Private Sub CommCompFacts(ByVal c_sect As Long, _
                          ByVal c_rev_no_Used As String, _
                          ByVal c_rev_no_raw As String)

    With Msg
        c_sect = c_sect + 1
        With .Section(c_sect)
            .Label.Text = c_rev_no_Used & " ="
            .Label.FontColor = rgbBlack
            .Label.FontBold = True
            .Label.MonoSpaced = True
            .Text.Text = "Revision Number of the used Common Component."
        End With
        c_sect = c_sect + 1
        With .Section(c_sect)
            .Label.Text = c_rev_no_raw & " ="
            .Label.FontColor = rgbBlack
            .Label.FontBold = True
            .Label.MonoSpaced = True
            .Text.Text = "Revision Number of the Common Component in the Common-Components folder."
        End With
        c_sect = c_sect + 1
        With .Section(c_sect)
            .Label.Text = CommonPublic.LastModInWrkbkFullName(sCompName) & " ="
            .Label.FontColor = rgbBlack
            .Label.FontBold = True
            .Label.MonoSpaced = True
            .Text.Text = "Workbook which last modified the Common Component - and saved the Export-file to the Common components folder."
        End With
    End With
                          
End Sub

Public Function CommCompHasBeenModifiedinTheServicedWorkbook(Optional ByRef c_mod_at_datetime_utc As String, _
                                                             Optional ByRef c_mod_in_wbk_fullname As String, _
                                                             Optional ByRef c_mod_on_machine As String) As Boolean
        
    If IsCommCompPublic Then
        If FSo.FileExists(ExpFileFullName) Then
            CommCompHasBeenModifiedinTheServicedWorkbook = Not CodeExprtd.Meets(CodePublic)
            c_mod_at_datetime_utc = ServicedLastModAt
            c_mod_in_wbk_fullname = Me.ServicedLastModIn
        End If
    End If

End Function

Public Function CommCompHasModificationPendingRelease(Optional ByRef p_last_mod_at_datetime_utc As String, _
                                                      Optional ByRef p_last_mod_export_filename As String, _
                                                      Optional ByRef p_last_mod_in_wbk_fullname As String, _
                                                      Optional ByRef p_last_mod_in_wbk_name As String, _
                                                      Optional ByRef p_last_mod_on_machine As String) As Boolean
    
    If IsCommComp Then
        CommCompHasModificationPendingRelease = CommonPending.Exists(sCompName _
                                                                   , p_last_mod_at_datetime_utc _
                                                                   , p_last_mod_export_filename _
                                                                   , p_last_mod_in_wbk_fullname _
                                                                   , p_last_mod_in_wbk_name _
                                                                   , p_last_mod_on_machine)
    End If
    
End Function

Public Function CommCompIsPendingByServicedWorkbook()
    
    Dim sLastModInWbkFullName As String
    
    If CommCompHasModificationPendingRelease(, , sLastModInWbkFullName) Then
        CommCompIsPendingByServicedWorkbook = sLastModInWbkFullName = WrkbkFullName _
                                              And CodeExprtd.Meets(CodePnding)
    End If
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String:   ErrSrc = "clsComp" & "." & sProc:   End Function

Private Function Exists(ByVal x_wbk As Workbook, _
                        ByVal x_comp_name As String, _
               Optional ByRef x_vbc As VBComponent) As Boolean
' ------------------------------------------------------------------------------
' The function returns TRUE and the corresponding VBComponent (x_vbc) when a
' named component (x_comp_name) exists in the Workbook (x_wbk).
' ------------------------------------------------------------------------------
    Exists = Not IsError(x_wbk.VBProject.VBComponents(x_comp_name))
    If Exists Then Set x_vbc = x_wbk.VBProject.VBComponents(x_comp_name)
End Function

Public Function ExpFileFullName() As String
    If sExpFileFullName = vbNullString Then
        sExpFileFullName = mEnvironment.ExportServiceFolderPath & "\" & sCompName & Me.ExpFileExt
    End If
    ExpFileFullName = sExpFileFullName
End Function

Private Function ExpFileTempFullName() As String
' ------------------------------------------------------------------------------
' Returns the name for a temporary Export File with the corresponding temporary
' folder already created thereby.
' ------------------------------------------------------------------------------
    Const PROC = "ExpFileTempFullName-Get"
    
    On Error GoTo eh
    If Not wbk Is Nothing Then
        With FSo
            If sTmpFolder = vbNullString Then
                sTmpFolder = Services.TempExportFolder
            End If
            If Not .FolderExists(sTmpFolder) Then
                mBasic.TimedDoEvents ErrSrc(PROC)
                .CreateFolder sTmpFolder
                mBasic.TimedDoEvents ErrSrc(PROC)
            End If
            ExpFileTempFullName = sTmpFolder & "\" & sCompName & sExpFileExt
        End With
    Else
        Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Workbook object yet not available for component '" & sCompName & "'!"
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsCommComp() As Boolean
    IsCommComp = Serviced.CompsCommon.Exists(sCompName)
End Function

Public Function IsCommCompPublic() As Boolean
    IsCommCompPublic = CommonPublic.Exists(sCompName)
End Function

Public Function IsCommCompPending() As Boolean
    IsCommCompPending = CommonPending.Exists(sCompName)
End Function

Public Sub Export()
' ------------------------------------------------------------------------------
' ! This clsComp service is essential. It ensures the update of the CodeExprtd !
' ! which needs to be updated to represent the current Export File's code.     !
' ------------------------------------------------------------------------------

    VBComp.Export ExpFileFullName
    Set CodeExprtd = Nothing
    Set CodeExprtd = New clsCode
    CodeExprtd.Source = ExpFileFullName
    CodeExprtd.CompName = sCompName
    
End Sub

Public Function IsCommCompUpToDate() As Boolean

    If Not CodePublic.IsNone Then
        If Not CodeExprtd.IsNone Then
            IsCommCompUpToDate = CodeExprtd.Meets(CodePublic)
        Else
            IsCommCompUpToDate = CodeCrrent.Meets(CodePublic)
        End If
    End If
    
End Function

Public Function IsThePendingSource() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when a Common Component (sCompName) is considered being the source
' of the current "pending release" by the fact that the user, the computer and
' the Workbook are identical with the registered "pending release".
' ------------------------------------------------------------------------------
    
    With CommonPending
        Select Case True
            Case .LastModBy(sCompName) <> mEnvironment.ThisComputersUser
            Case .LastModInWrkbkFullName(sCompName) <> Serviced.Wrkbk.FullName
            Case .LastModOn(sCompName) <> mEnvironment.ThisComputersName
            Case Else
                IsThePendingSource = True
        End Select
    End With
    
End Function

Private Function IsWrkbk() As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbk = Err.Number = 0
    
End Function

Public Function ServicedMeetPublicProperties() As Boolean
    Const PROC = "ServicedMeetPublicProperties"
    
    Dim sAt As String
    Dim sBy As String
    Dim sIn As String
    Dim sOn As String
    Dim sXf As String
    
    With CommonPublic
        sAt = .LastModAt(sCompName)
        sBy = .LastModBy(sCompName)
        sIn = .LastModInWrkbkFullName(sCompName)
        sOn = .LastModOn(sCompName)
        sXf = .LastModExpFileOrigin(sCompName)
    End With
    
    With CommonServiced
        Select Case True
            Case sAt <> .LastModAt(sCompName):              ' Debug.Print ErrSrc(PROC) & " " & sCompName & " LastModAt differs"
            Case sBy <> .LastModBy(sCompName):              ' Debug.Print ErrSrc(PROC) & " " & sCompName & " LastModBy differs"
            Case sXf <> .LastModExpFileOrigin(sCompName):   ' Debug.Print ErrSrc(PROC) & " " & sCompName & " LastModXf differs"
            Case sIn <> .LastModInWrkbkFullName(sCompName): ' Debug.Print ErrSrc(PROC) & " " & sCompName & " LastModIn differs"
            Case sOn <> .LastModOn(sCompName):              ' Debug.Print ErrSrc(PROC) & " " & sCompName & " LastModOn differs"
            Case Else:  ServicedMeetPublicProperties = True
        End Select
    End With
    
End Function

Private Sub RemoveLeadingTrailingEmptyLines(ByRef dct As Dictionary)
' ------------------------------------------------------------------------------
' Removes any empty lines (length = 0) from the beginning and the end of a
' Dictionary (dct).
' ------------------------------------------------------------------------------
    Const PROC = "RemoveLeadingTrailingEmptyLines"
    
    On Error GoTo eh
    If dct.Count > 0 Then
        While Len(Trim(dct.Items()(0))) = 0                 ' Remove leading empty items
            dct.Remove dct.Keys()(0)
        Wend
    End If
    
    If dct.Count > 0 Then
        While Len(Trim(dct.Items()(dct.Count - 1))) = 0     ' Remove trailing empty items
            dct.Remove dct.Keys()(dct.Count - 1)
        Wend
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub RemoveModuleInfo(ByRef dct As Dictionary)
' ------------------------------------------------------------------------------
' Remove the extra module information lines from a Dictionary which resulted
' from the contents of a VBComponent's Export-File.
' ------------------------------------------------------------------------------
    Const ATTRIB_VB = "Attribute VB_*"
    
    Dim i   As Long
    Dim j   As Long
    
    For i = Min(15, dct.Count - 1) To 0 Step -1
        If dct.Items()(i) Like ATTRIB_VB Then Exit For
    Next i
        
    For j = 0 To i
        dct.Remove dct.Keys()(0)
    Next j
    
End Sub

Public Sub SetPendingEqualServiced()
' ------------------------------------------------------------------------------
' Set the pending release Common Component's properties in the Pending.dat file
' equal to those in the serviced Common Components CommComps.dat file - when the
' modified Common Component is registered pending release.
' ------------------------------------------------------------------------------
    Dim sAt As String
    Dim sBy As String
    Dim sIn As String
    Dim sKo As String
    Dim sOn As String
    Dim sXf As String
    
    With CommonServiced
        sAt = .LastModAt(sCompName)
        sBy = .LastModBy(sCompName)
        sIn = .LastModInWrkbkFullName(sCompName)
        sKo = .KindOfComponent(sCompName)
        sOn = .LastModOn(sCompName)
        sXf = .LastModExpFileOrigin(sCompName)
    End With
    With CommonPending
        .LastModAt(sCompName) = sAt
        .LastModBy(sCompName) = sBy
        .LastModExpFileOrigin(sCompName) = sXf
        .LastModInWrkbkFullName(sCompName) = sIn
        .LastModKindOfComp(sCompName) = sKo
        .LastModOn(sCompName) = sOn
    End With
   
End Sub

Public Sub SetPublicEqualPending()
' ------------------------------------------------------------------------------
' Set the public Common Component's properties in the CommComps.date file in
' the Common-Components folder equal to those in the pending release Common
' Components PendingReleases.dat file - in the Common-Components folder - when
' the pending release component is released to public.
' ------------------------------------------------------------------------------
    Dim sAt As String
    Dim sBy As String
    Dim sIn As String
    Dim sOn As String
    Dim sKo As String
    Dim sXf As String
    
    With CommonPending
        sAt = .LastModAt(sCompName)
        sBy = .LastModBy(sCompName)
        sIn = .LastModInWrkbkFullName(sCompName)
        sOn = .LastModOn(sCompName)
        sKo = .LastModKindOfComp(sCompName)
        sXf = .LastModExpFileOrigin(sCompName)
    End With
    With CommonPublic
        .LastModAt(sCompName) = sAt
        .LastModBy(sCompName) = sBy
        .LastModExpFileOrigin(sCompName) = sXf
        .LastModInWrkbkFullName(sCompName) = sIn
        .LastModKindOfOriginComp(sCompName) = sKo
        .LastModOn(sCompName) = sOn
    End With
    
End Sub

Public Sub SetPublicEqualServiced()
' ------------------------------------------------------------------------------
' Attention: This is an exception for the case a component's Export-File has
'            manually been copied into the Common-Components folder indicating
'            it as a new Common Component.
' ------------------------------------------------------------------------------
    Dim sAt As String
    Dim sBy As String
    Dim sIn As String
    Dim sOn As String
    Dim sKo As String
    Dim sXf As String
    
    sAt = ServicedLastModAt
    sBy = ServicedLastModBy
    sIn = ServicedLastModIn
    sOn = ServicedLastModOn
    sKo = ServicedLastModKindOfComp
    sXf = ServicedLastModExpFile
    
    With CommonPublic
        .LastModAt(sCompName) = sAt
        .LastModBy(sCompName) = sBy
        .LastModExpFileOrigin(sCompName) = sXf
        .LastModInWrkbkFullName(sCompName) = sIn
        .LastModKindOfOriginComp(sCompName) = sKo
        .LastModOn(sCompName) = sOn
    End With
    
End Sub

Public Sub SetServicedEqualPublic()
' ------------------------------------------------------------------------------
' Set the serviced Common Component's properties in the CommComps.dat file
' equal to those of the public Common Component's properties in the CommComp.dat
' file in the Common-Components folder - when the outdated Common Component had
' been updated.
' ------------------------------------------------------------------------------
    Dim sAt As String
    Dim sBy As String
    Dim sIn As String
    Dim sOn As String
    Dim sXf As String
    
    With CommonPublic
        sAt = .LastModAt(sCompName)
        sBy = .LastModBy(sCompName)
        sIn = .LastModInWrkbkFullName(sCompName)
        sOn = .LastModOn(sCompName)
        sXf = .LastModExpFileOrigin(sCompName)
    End With
    With CommonServiced
        .LastModAt(sCompName) = sAt
        .LastModBy(sCompName) = sBy
        .LastModExpFileOrigin(sCompName) = sXf
        .LastModInWrkbkFullName(sCompName) = sIn
        .LastModOn(sCompName) = sOn
    End With
    
End Sub

Public Sub SetServicedProperties()
' ------------------------------------------------------------------------------
' Set the serviced Common Component's properties in the CommComps.dat file when
' the modified Common Component has been exported.
' ------------------------------------------------------------------------------
    Dim sAt As String
    Dim sBy As String
    Dim sIn As String
    Dim sOn As String
    Dim sXf As String
    
    sAt = mCompMan.UTC(FSo.GetFile(Me.ExpFileFullName).DateLastAccessed)
    sBy = mEnvironment.ThisComputersUser
    sIn = Serviced.Wrkbk.FullName
    sOn = mEnvironment.ThisComputersName
    sXf = ExpFileFullName
    With CommonServiced
        .LastModAt(sCompName) = sAt
        .LastModBy(sCompName) = sBy
        .LastModExpFileOrigin(sCompName) = sXf
        .LastModInWrkbkFullName(sCompName) = sIn
        .LastModOn(sCompName) = sOn
    End With
    
End Sub

Public Function TypeString() As String
' ------------------------------------------------------------------------------
' Returns the component's type as String.
' ------------------------------------------------------------------------------
    If Not vbc Is Nothing Then
        On Error Resume Next
        TypeString = dctType(vbc.Type)
        If TypeString = vbNullString Then TypeString = "unknown"
    End If
End Function

Public Function WrkbkBaseName() As String
    WrkbkBaseName = FSo.GetBaseName(sWrkbkFullName)
End Function

