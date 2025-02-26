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
' ===================== component of the serviced ActiveWorkbook. The module
' provides the following instances of the class module clsCode:
' - CodeCrrent: The current component'a code,
' - CodePublic: An existing public code, when it exists as a a public Common Component
' - CodePnding: When it exists "pending release"
' - CodeExprtd: The exported code (when the component's export file exists).
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
Private sTmpFolder          As String       ' Folder for a temporary Export File
Private sWrkbkFullName      As String       ' Common Component host Workbook's full name
Private vbc                 As VBComponent  ' VBComponent managed by this Class Module
Private wbk                 As Workbook     ' The class' Workbook
Private lMaxLenComp         As Long
Private wsh                 As Worksheet

Public CodeExprtd           As clsCode
Public CodeCrrent           As clsCode
Public CodePublic           As clsCode
Public CodePnding           As clsCode

Private Sub Class_Initialize()
    Set wbk = ActiveWorkbook
    sWrkbkFullName = wbk.FullName
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
        If fso.FileExists(sExpFileFullName) Then
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
                .SourceDesc = mDiff.SourcePublic(sCompName)
            End With
        End If
            
        If fso.FileExists(CommonPending.ExportFiles(sCompName)) Then
            With CodePnding
                .CompName = sCompName
                .Source = CommonPending.ExportFiles(sCompName)
                .SourceDesc = mDiff.SourcePending(sCompName)
            End With
        End If
    End If

End Property

Friend Property Get ExpFile() As File
       
    If flExpFile Is Nothing Then
        If fso.FileExists(ExpFileFullName) Then
            Set flExpFile = fso.GetFile(ExpFileFullName)
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

Friend Property Get KindOfComp() As enKindOfComp:           KindOfComp = lKindOfComp:                                           End Property

Friend Property Get MaxLenComp() As Long:                   MaxLenComp = lMaxLenComp:                                           End Property

Friend Property Get PendingKindOfComp() As String:          PendingKindOfComp = CommonPending.LastModKindOfComp(sCompName):     End Property

Friend Property Get PendingLastModAt() As String:           PendingLastModAt = CommonPending.LastModAt(sCompName):              End Property

Friend Property Get PendingLastModBy() As String:           PendingLastModBy = CommonPending.LastModBy(sCompName):              End Property

Friend Property Get PendingLastModIn() As String:           PendingLastModIn = CommonPending.LastModInWrkbkFullName(sCompName): End Property

Friend Property Get PendingLastModOn() As String:           PendingLastModOn = CommonPending.LastModOn(sCompName):              End Property

Friend Property Let PendingLastModOn(ByVal s As String):    CommonPending.LastModOn(sCompName) = s:                             End Property

Friend Property Get PublicLastModAt() As String:            PublicLastModAt = CommonPublic.LastModAt(sCompName):                End Property

Friend Property Get PublicLastModBy() As String:            PublicLastModBy = CommonPublic.LastModBy(sCompName):                End Property

Friend Property Get PublicLastModIn() As String:            PublicLastModIn = CommonPublic.LastModInWrkbkFullName(sCompName):   End Property

Friend Property Get PublicLastModOn() As String:            PublicLastModOn = CommonPublic.LastModOn(sCompName):                End Property

                         
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

Friend Property Get ServicedLastModIn() As String:              ServicedLastModIn = CommonServiced.LastModInWrkbkFullName(sCompName):       End Property

Friend Property Let ServicedLastModIn(ByVal s As String):       CommonServiced.LastModInWrkbkFullName(sCompName) = s:                       End Property

Friend Property Get ServicedLastModKindOfComp() As String:      ServicedLastModKindOfComp = CommonServiced.KindOfComponent(sCompName):      End Property

Friend Property Get ServicedLastModOn() As String:              ServicedLastModOn = CommonServiced.LastModOn(sCompName):                    End Property

Friend Property Let ServicedLastModOn(ByVal s As String):       CommonServiced.LastModOn(sCompName) = s:                                    End Property

Friend Property Get Sheet() As Worksheet:                       Set Sheet = wsh:                                                            End Property

Friend Property Let Sheet(ByRef s_wsh As Worksheet):            wsh = s_wsh:                                                                End Property

Friend Property Get VBComp() As VBComponent:                    Set VBComp = wbk.VBProject.VBComponents(sCompName):                         End Property

'Friend Property Get Wrkbk() As Workbook:                        Set Wrkbk = wbk:                                                            End Property

Friend Property Get WrkbkFullName() As String:                  WrkbkFullName = sWrkbkFullName:                                             End Property

Public Function Changed() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Component's Export-File does not exist or the
' Export-File's code differs from the current code of the component.
' ------------------------------------------------------------------------------
    Const PROC = "Changed"
    
    On Error GoTo eh
    Select Case True
        Case Not fso.FileExists(Me.ExpFileFullName)     ' first time export
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

Public Function CommCompHasModificationPendingRelease(Optional ByRef c_last_mod_at_datetime_utc As String, _
                                                      Optional ByRef c_last_mod_export_filename As String, _
                                                      Optional ByRef c_last_mod_in_wbk_fullname As String, _
                                                      Optional ByRef c_last_mod_on_machine As String) As Boolean
    
    If IsCommon Then
        CommCompHasModificationPendingRelease = CommonPending.Exists(sCompName _
                                                                   , c_last_mod_at_datetime_utc _
                                                                   , c_last_mod_export_filename _
                                                                   , c_last_mod_in_wbk_fullname _
                                                                   , c_last_mod_on_machine)
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
        With fso
            If sTmpFolder = vbNullString Then
                sTmpFolder = Servicing.TempExportFolder
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

Public Sub Export()
' ------------------------------------------------------------------------------
' ! This clsComp service is essential. It ensures the update of the CodeExprtd !
' ! which needs to be updated to represent the current Export File's code.     !
' ------------------------------------------------------------------------------

    VBComp.Export ExpFileFullName
    Set CodeExprtd = Nothing
    Set CodeExprtd = New clsCode
    With CodeExprtd
        .CompName = sCompName
        .Source = ExpFileFullName
    End With
    
End Sub

Public Function IsCommCompPending() As Boolean
    IsCommCompPending = CommonPending.Exists(sCompName)
End Function

Public Function IsCommCompPublic() As Boolean
    IsCommCompPublic = CommonPublic.Exists(sCompName)
End Function

Public Function IsCommCompUpToDate() As Boolean

    If Not CodePublic.IsNone Then
        If Not CodeExprtd.IsNone Then
            IsCommCompUpToDate = CodeExprtd.Meets(CodePublic)
        Else
            IsCommCompUpToDate = CodeCrrent.Meets(CodePublic)
        End If
    End If
    
End Function

Public Function IsCommon() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the component (c_comp) is either known as a public Common
' Component (in the Common-Components folder) or a hosted Common Component
' (which yet has ever been released to public).
' ----------------------------------------------------------------------------
    IsCommon = CommonPublic.Exists(sCompName) _
              Or Serviced.Hosted.Exists(sCompName)
    
End Function

Public Function IsCommonUsedOrHosted() As Boolean
' ----------------------------------------------------------------------------
' Returns True when the component (c_comp) is either known as a public Common
' Component (in the Common-Components folder) or a hosted Common Component
' (which yet has ever been released to public).
' ----------------------------------------------------------------------------
    IsCommonUsedOrHosted = CommonPublic.Exists(sCompName) _
                        And (KindOfComp = enCompCommonHosted _
                          Or KindOfComp = enCompCommonUsed)
    
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

Public Sub RegisterPending()
    CommonPending.Register Me
End Sub

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

Public Sub SetServicedProperties()
' ------------------------------------------------------------------------------
' Set the serviced Common Component's properties in the CommComps.dat file when
' the modified Common Component has been exported.
' ------------------------------------------------------------------------------
    Dim prprts As New clsProperties
    With prprts
        .LmAt = mCompMan.UTC(fso.GetFile(Me.ExpFileFullName).DateLastModified)
        .LmBy = mEnvironment.ThisComputersUser
        .LmIn = Serviced.Wrkbk.FullName
        .LmOn = mEnvironment.ThisComputersName
        .LmXf = ExpFileFullName
    End With
    CommonServiced.Properties(sCompName) = prprts
End Sub

Public Function WrkbkBaseName() As String
    WrkbkBaseName = fso.GetBaseName(sWrkbkFullName)
End Function

