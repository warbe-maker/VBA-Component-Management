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
' Class Module clsComp
'       Represents a VB-Project's Component with extended (Management) properties and methods.
'       The Component may be of either of the following kinds:
'       - a used Component which may also be a used Common Component
'       - the original/raw of a Common Component, i.e. a Component of which the original is
'         hosted in a dedicated Common Component Workbookorigin code of a Common Component
'
' Friend Properties:
' - Changed
' - CodeName
' - IsCommComp    True when the component is a Common Component, i.e. one of which an
'               Export-File exists in the Common-Components folder.
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
' Services:
' - CompType
' - Export
' - RemoteSheetExists
' - ReplaceRawWithClone
' - ReplaceRawWithCloneWhenConfirmed
'
' -----------------------------------------------------------------------------------
Private bIsCommComp         As Boolean
Private dctChanges          As Dictionary   ' Dictionary of code changes
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
Private sService            As String       ' Name of the 'Service' folder, i.e. the folder of the Workbook
Private dctType             As Dictionary
Private lMaxLenComp         As Long
Private wsh                 As Worksheet

Public Raw                  As clsRaw       ' The raw component's properties for a Used Common Component
Public CodeExported         As New clsCode
Public CodeCurrent          As New clsCode
Public CodePublic           As New clsCode

Private Sub Class_Initialize()
    If dctType Is Nothing Then Set dctType = New Dictionary Else dctType.RemoveAll
    dctType.Add vbext_ct_ActiveXDesigner, "ActiveX Designer"
    dctType.Add vbext_ct_ClassModule, "Class Module"
    dctType.Add vbext_ct_Document, "Document Module"
    dctType.Add vbext_ct_MSForm, "UserForm"
    dctType.Add vbext_ct_StdModule, "Standard Module"
End Sub

Private Sub Class_Terminate()
    
    Set dctChanges = Nothing
    Set wbk = Nothing
    Set vbc = Nothing
    Set flExpFile = Nothing
    CleanUpTemps
    Set Raw = Nothing
    Set CodeCurrent = Nothing
    Set CodeExported = Nothing
    Set CodePublic = Nothing
End Sub

Public Function Changed() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Component's Export-File does not exist or the
' Export-File's code differs from the current code of the component.
' ------------------------------------------------------------------------------
    Const PROC = "Changed"
    
    On Error GoTo eh
    Select Case True
        Case Not CodeCurrent.IsNone And CodeExported.IsNone
            Changed = True
        Case Not CodeCurrent.IsNone And Not CodeExported.IsNone
            Changed = CodeCurrent.DiffersFrom(CodeExported, False)
        Case Not CodeCurrent.IsNone And CodeExported.IsNone
            Changed = True
        Case CodeCurrent.IsNone And Not CodeExported.IsNone
            '~~ This should be covered elsewhere. Die Export-File has become obsolete eand should be removed
        Case CodeCurrent.IsNone And CodeExported.IsNone
            '~~ That is fine (will be ignored)
    End Select
    
xt: Exit Function
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

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
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get CodeName(Optional ByVal scn_name As String) As String
' ------------------------------------------------------------------------------
' Returns the sheet's CodeName through its Name.
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet
    
    For Each wsh In wbk.Worksheets
        If wsh.Name = scn_name Then
            CodeName = wsh.CodeName
            Exit For
        End If
    Next wsh
    
End Property

Friend Property Get CompName() As String:           CompName = sCompName:       End Property

Friend Property Let CompName(ByVal s As String)
' ------------------------------------------------------------------------------
' Provides the properties: CompName        (sCompName)
'                          ExpFileFullName (sExpFileFullName)
'                          VBComp          (vbc)
'                          ExpFileExt      (sExpFileExt)
' ------------------------------------------------------------------------------
    Const PROC = "CompName-Let"
    
    On Error GoTo eh
    If s = vbNullString Then Stop
    If wbk Is Nothing _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Component Name assigned for a yet unknown Workbook!"
    If Not Exists(wbk, s, vbc) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "A components named " & s & " does not exist in the Workbook " & wbk.Name & " !."
    
    sCompName = s
    CommComps.CompName = s
    If vbc Is Nothing Then Set vbc = wbk.VBProject.VBComponents(s)
    Select Case vbc.Type
        Case vbext_ct_StdModule:    sExpFileExt = ".bas"
        Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
        Case vbext_ct_MSForm:       sExpFileExt = ".frm"
        Case vbext_ct_Document:     sExpFileExt = ".cls"
    End Select
    sExpFileFullName = mExport.ExpFileFolderPath(wbk) & "\" & vbc.Name & sExpFileExt
    
    If CommComps.ExistsPublic(sCompName) _
    Or CommComps.ExistsPending(sCompName) Then
        '~~ The component is known/existing public or - in case of a new hosted component -
        '~~ has already been registered as a pending release component
        IsCommComp = True
        Set Raw = New clsRaw
        Raw.CompName = vbc.Name
        CodePublic.Source = CommComps.CurrentPublicExpFileFullName
    End If
'    If sCompName = "wbCompMan" Then Stop
    CodeCurrent.Source = vbc.CodeModule
    CodeExported.Source = sExpFileFullName
    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get ExistsBySheetName(Optional ByVal eb_name As String) As Boolean
    Dim wsh As Worksheet
    For Each wsh In Me.Wrkbk.Worksheets
        ExistsBySheetName = wsh.Name = eb_name
        If ExistsBySheetName Then Exit For
    Next wsh
End Property

Friend Property Get ExpFile() As File
        
    With Me
        If flExpFile Is Nothing Then
            If FSo.FileExists(.ExpFileFullName) Then
                Set flExpFile = FSo.GetFile(.ExpFileFullName)
            Else
                '~~ When the component has yet no never been exported
                '~~ it will be along with the Export changed service
            End If
        End If
    End With
    
    Set ExpFile = flExpFile
    
End Property

Friend Property Let ExpFile(ByVal fl As File):      Set flExpFile = fl:         End Property

Friend Property Get ExpFileExt() As String:         ExpFileExt = sExpFileExt:   End Property

Friend Property Get ExpFileFullName() As String
    If sExpFileFullName = vbNullString Then
        sExpFileFullName = mExport.ExpFileFolderPath(Me.Wrkbk) & "\" & sCompName & Me.ExpFileExt
    End If
    ExpFileFullName = sExpFileFullName
End Property

Friend Property Get ExpFileTemp() As File
    With FSo
        If Not .FileExists(Me.ExpFileTempFullName) Then
            Me.VBComp.Export Me.ExpFileTempFullName
        End If
        Set ExpFileTemp = .GetFile(Me.ExpFileTempFullName)
    End With
End Property

Friend Property Get ExpFileTempFullName() As String
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
    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get IsCommComp() As Boolean:        IsCommComp = bIsCommComp:   End Property

Friend Property Let IsCommComp(ByVal b As Boolean): bIsCommComp = b:            End Property

Friend Property Get IsWrkbk() As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbk = Err.Number = 0
    
End Property

Friend Property Get KindOfComp() As enKindOfComp
' ------------------------------------------------------------------------------
' Returns the Kind of Component (lKindOfComp) or determines and saves it when
' it is still unknown. This property ensures that, when the KindOfComp is
' requested for the very first time a Raw class object is established when the
' component is a Used Common Component.
' Attention: For a correct analysis it is mandatory that the services
'            HskpngHosted and ManageUsedCommonComponent had been
'            executed (which is done by default when the serviced Workbook is
'            opened and when it is saved.
' ------------------------------------------------------------------------------
    Const PROC = "KindOfComp-Get"
    
    On Error GoTo eh
    If Me.VBComp Is Nothing _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
                   "The kind of component cannot be returned since the VB-Component is yet unknown!"
                   
    If lKindOfComp = mCompMan.enUnknown Then
        If CommComps.ExistsPublic(sCompName) Then
            '~~ The component is known as a public Common Component.
            '~~ I.e. the component's Export-File exists in the Common-Components folder.
            If CompManDat.RegistrationState(Me.CompName) = enRegStateHosted Then
                '~~ The Common Component is currently registered as beeing hosted in the service Workbook
                lKindOfComp = mCompMan.enCommCompHosted
            Else
                If Not CompManDat.RegistrationState(Me.CompName) = enRegStatePrivate Then
                    lKindOfComp = mCompMan.enCommCompUsed
                End If
            End If
        ElseIf lKindOfComp = mCompMan.enUnknown Then
            lKindOfComp = mCompMan.enInternal
        End If
    End If
    
xt: KindOfComp = lKindOfComp
    Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Let KindOfComp(ByVal en As enKindOfComp):   lKindOfComp = en:           End Property

Friend Property Get MaxLenComp() As Long:                   MaxLenComp = lMaxLenComp:   End Property

Friend Property Get Outdated() As Boolean
' ------------------------------------------------------------------------------
' Returns (exclusively for a used or hosted Common Component) TRUE when the
' component's public Common Component (in the Common-Components folder) differs
' from the current code and the component's Last-Modified-Date is less then the
' Last-Modified-Date of the public Common Component in the Common-Components
' folder.
' Note: In case the component is not outdated but just the opposite, i.e. the
'       component differs but has a more recent Last-Modified-Date the component
'       is a Release candidate and collected as such
' ------------------------------------------------------------------------------
    Const PROC                  As String = "Outdated"
    
    On Error GoTo eh
    Dim sCommCompPublicLastModDateTime  As String
    Dim bCodeDiffers                    As Boolean
    Dim i                               As Long
    
    mBasic.BoP ErrSrc(PROC)
    CommComps.CompName = sCompName
    sCommCompPublicLastModDateTime = CommComps.LastModAtDateTime(sCompName)
    CodePublic.Source = CommComps.CurrentPublicExpFileFullName
    CodeCurrent.Source = Me.ExpFileFullName
    
    '~~ Note: Outdated is only considered when a public code exists
    '~~       and this component's code differs from the current public code
    If Not CodePublic.IsNone Then
        bCodeDiffers = CodeCurrent.DiffersFrom(CodePublic, True)
        If bCodeDiffers Then Me.CodeCurrent.DiffersFromDebug CodePublic, True
    End If
    
    mCompMan.MsgInit
    ' Cases: --------------------------------1-2-3-4-5-6--------------
    ' C1 This code differs from the public   y y y . . .
    ' C1 This LastModInf = LastModinf public y - - . . . .
    ' C1 This LastModInf < LastModinf public - y - . . . .
    ' C1 This LastModInf > LastModinf public - - y . . . .
    ' --------------------------------------
    '    Inconsistency message               x
    '    Outdated = TRUE                       x
    '    Manage Pending Release                  x
    ' ---------------------------------------------
    Select Case True
        Case bCodeDiffers And LastModAtDateTime = sCommCompPublicLastModDateTime
            '~~ The Common Component in the Common-Components folder and the used/hosted Common Component
            '~~ do have equal last modified information although the code differs.
            '~~ This is regarded a severe logic error!
            i = i + 1
            Msg.Section(i).Text.Text = _
                "While the Export-File of the of the " & CompManDat.RegStateString(Me.CompName) & _
                "Common Component   " & mBasic.Spaced(Me.CompName) & "   differs from the " & _
                "Export File in the Common-Components folder, the Revision numbers are equal." & vbLf & _
                "This is a severe logic error which requires further investigation regarding " & _
                "the cause (code inspection and testing)." & vbLf & _
                "See facts below."
            CommCompFacts i, LastModAtDateTime, sCommCompPublicLastModDateTime
            Do
                Select Case mMsg.Dsply(dsply_title:="Severe inconsistent Common-Component facts!" _
                                     , dsply_msg:=Msg _
                                     , dsply_Label_spec:="R130" _
                                     , dsply_width_min:=30 _
                                     , dsply_buttons:=mMsg.Buttons(BTTN_DSPLY_DIFF, vbLf, vbOKOnly))
                    Case vbOK: GoTo xt
                    Case BTTN_DSPLY_DIFF
                         CodeCurrent.DsplyDiffs d_this_file_name:="CurrentCode" _
                                              , d_this_file_title:="The Common Component's current code" _
                                              , d_from_code:=CodePublic _
                                              , d_from_file_name:="CurrentPublic" _
                                              , d_from_file_title:="The Common Component's current public code"
                End Select
            Loop
        Case bCodeDiffers And LastModAtDateTime < sCommCompPublicLastModDateTime
            Outdated = True
        Case LastModAtDateTime > sCommCompPublicLastModDateTime
            '~~ When the used/hosted Common Component's last modified date-time is after the last modified date-time
            '~~ of the public/raw Common Component this is regarded a pending Release of a code modification,
            '~~ provided there isn't one registered with a modification in another Workbook.
            CommComps.PendingReleaseManagement Me
            GoTo xt
    End Select
    
    If Not Outdated Then
        If LastModAtDateTime <> sCommCompPublicLastModDateTime Then
            LastModAtDateTime = sCommCompPublicLastModDateTime
        End If
        With Services
            .ServicedItem = Me.VBComp
            .ServicedItemLogEntry "Component is up-to-date! Revision-Number has been set to to " & sCommCompPublicLastModDateTime & " (the raw's in the Common Components Folder)"
        End With
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Property
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

                          
Friend Property Get LastModAtDateTime() As String:             LastModAtDateTime = CompManDat.LastModAtDateTime(sCompName): End Property

Friend Property Let LastModAtDateTime(ByVal s As String):      CompManDat.LastModAtDateTime(sCompName) = s:              End Property

Friend Property Get Service() As String
    If sService = vbNullString Then
        sService = FSo.GetFileName(Me.Wrkbk.Path)
    End If
    Service = sService
End Property

Friend Property Get Sheet() As Worksheet:                   Set Sheet = wsh:            End Property

Friend Property Let Sheet(ByRef s_wsh As Worksheet):        Set wsh = s_wsh:            End Property

Friend Property Get TypeString() As String
' ------------------------------------------------------------------------------
' Returns the component's type as String.
' ------------------------------------------------------------------------------
    If Not vbc Is Nothing Then
        On Error Resume Next
        TypeString = dctType(vbc.Type)
        If TypeString = vbNullString Then TypeString = "unknown"
    End If
End Property

Friend Property Get VBComp() As VBComponent:                Set VBComp = vbc:           End Property

Friend Property Let VBComp(ByRef vb_comp As VBComponent)
    Set vbc = vb_comp
    If sCompName = vbNullString Then
        '~~ When the component's name has yet not been set
        CompName = vb_comp.Name
    End If
End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    
    If wbk Is Nothing Then
        If Not FSo.FileExists(sWrkbkFullName) _
        Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
        Set wbk = mCompMan.WbkGetOpen(sWrkbkFullName)
    Else
        Set Wrkbk = wbk
    End If
    
xt: Exit Property
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Let Wrkbk(ByRef w_wbk As Workbook)
' ------------------------------------------------------------------------------
' Provides the properties: - Wrkbk         (wbk)
'                          - WrkbFullName  (sWrkbkFullName)
'                          - WrkbkBaseName (sWrkbkBaseName)
'                          - ExpFilePath   (sExpFilePath)
'                          - Service       (sService)
'                          - MaxLenComp    (lMaxLenComp)
' ------------------------------------------------------------------------------
    Dim vbc As VBComponent
    
    Set wbk = w_wbk
    sWrkbkFullName = wbk.FullName
    With FSo
        sWrkbkBaseName = .GetBaseName(sWrkbkFullName)
        sExpFilePath = .GetParentFolderName(sWrkbkFullName)
        sService = .GetFileName(wbk.Path)
    End With
    For Each vbc In wbk.VBProject.VBComponents
        lMaxLenComp = mBasic.Max(lMaxLenComp, Len(vbc.Name))
    Next vbc
    If Not Me.VBComp Is Nothing And Me.CompName = vbNullString Then Me.CompName = Me.VBComp.Name
    
End Property

Friend Property Get WrkbkBaseName() As String:                 WrkbkBaseName = BaseName(sWrkbkFullName):        End Property

Friend Property Get WrkbkFullName() As String:                 WrkbkFullName = sWrkbkFullName:                  End Property

Friend Property Let WrkbkFullName(ByVal s As String):          sWrkbkFullName = s:                              End Property

Private Function BaseName(ByVal s As String) As String
    With New FileSystemObject:  BaseName = .GetBaseName(s): End With
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
            .Label.Text = CommComps.LastModInWbkName(sCompName) & " ="
            .Label.FontColor = rgbBlack
            .Label.FontBold = True
            .Label.MonoSpaced = True
            .Text.Text = "Workbook which last modified the Common Component - and saved the Export-file to the Common components folder."
        End With
    End With
                          
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
End Function

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

Public Sub Export()
' ------------------------------------------------------------------------------
' Export of a modified component (or one which had yet not ever been exported)
' to its dedicated export folder in the serviced Workbook's root folder.
'
' When the modified component is one one as a 'Common Component' (either used
' or hosted the Export-File is copied ti the Common-Components folder and the
' Revision-Number is increased.
' ------------------------------------------------------------------------------
    Const PROC = "Export"
    
    On Error GoTo eh
    
    VBComp.Export ExpFileFullName   ' 1. The 'system' export
    
    If KindOfComp = enCommCompHosted _
    Or KindOfComp = enCommCompUsed Then
        LastModAtDateTime = mCompMan.UTC
        CommComps.PendingReleaseManagement Me
    End If

xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
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

