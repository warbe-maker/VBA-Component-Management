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
Private bIsCommComp           As Boolean
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
Private fso                 As FileSystemObject
Public Raw                  As clsRaw       ' The raw component's properties for a Used Common Component

Private Sub Class_Initialize()
    If dctType Is Nothing Then Set dctType = New Dictionary Else dctType.RemoveAll
    dctType.Add vbext_ct_ActiveXDesigner, "ActiveX Designer"
    dctType.Add vbext_ct_ClassModule, "Class Module"
    dctType.Add vbext_ct_Document, "Document Module"
    dctType.Add vbext_ct_MSForm, "UserForm"
    dctType.Add vbext_ct_StdModule, "Standard Module"
    Set fso = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set dctChanges = Nothing
    Set wbk = Nothing
    Set vbc = Nothing
    Set flExpFile = Nothing
    
    CleanUpTemps
    
    Set fso = Nothing
    Set Raw = Nothing
    
End Sub

Friend Property Get Changed() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Component's Export-File dose not exist or a temporary
' Export-File differs from the existing one.
' ------------------------------------------------------------------------------
    Const PROC = "Changed"
    
    On Error GoTo eh
    '~~ Any component is regarded changed when a temporary Export-File differs from the regular one
    Me.VBComp.Export Me.ExpFileTempFullName
    mBasic.TimedDoEvents ErrSrc(PROC)
    
    '~~ Any component its Export-File not exists is regarded changed
    '~~ In fact it is either new or the Export-Folder location had changed
    If Not fso.FileExists(Me.ExpFileFullName) Then
        Changed = True
        GoTo xt
    End If
    
    '~~ Any component is regarded changed when the collection of changed lines (dctChanges)
    '~~ not is nothing and the number of collected lines is <> 0
    If dctChanges Is Nothing Then
        '~~ The difference between the regular and the temporary export file
        '~~ is only provided once
        Set dctChanges = mCompMan.Services.FilesDifference(f_file_1:=fso.GetFile(Me.ExpFileTempFullName) _
                                                         , f_file_2:=fso.GetFile(Me.ExpFileFullName))
    End If
    Changed = dctChanges.Count <> 0
    
xt: Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

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

Friend Property Get IsCommComp() As Boolean:        IsCommComp = bIsCommComp:   End Property

Friend Property Let IsCommComp(ByVal b As Boolean): bIsCommComp = b:            End Property

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
    Dim sModByWbkName       As String
    Dim sRawExpFileFullName As String
    Dim fRawExpFile         As File
    Dim sLastModRevNo       As String
    
    If s = vbNullString Then Stop
    sCompName = s
     If Not wbk Is Nothing Then
        If Exists(wbk, vbc) Then
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
                Case vbext_ct_Document:     sExpFileExt = ".cls"
            End Select
            sExpFileFullName = mExport.ExpFileFolderPath(wbk) & "\" & vbc.Name & sExpFileExt
        End If
        If Not CommComps Is Nothing Then
            If CommComps.Exists(vbc, sExpFileExt, sModByWbkName, sRawExpFileFullName, fRawExpFile, sLastModRevNo) Then
                Me.IsCommComp = True
                Set Raw = New clsRaw
                With Raw
                    .CompName = vbc.Name
                    .LastModExpFile = fRawExpFile
                    .LastModByWbkName = sModByWbkName
                    .LastModRevisionNumber = sLastModRevNo
                End With
            End If
        End If
    Else
        Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Component Name assigned for a yet unknown Workbook!"
    End If
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function Exists(ByVal x_wbk As Workbook, _
              Optional ByRef x_vbc As VBComponent) As Boolean
    Dim v As VBComponent
    If Not x_wbk Is Nothing Then
        On Error Resume Next
        Set v = x_wbk.VBProject.VBComponents(sCompName)
        Exists = Not v Is Nothing
    ElseIf vbc Is Nothing Then
        On Error Resume Next
        Set vbc = wbk.VBProject.VBComponents(sCompName)
        Exists = Not vbc Is Nothing
    Else
        Exists = Not vbc Is Nothing
    End If
    If Exists Then Set x_vbc = wbk.VBProject.VBComponents(sCompName)
    
End Function

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
            If fso.FileExists(.ExpFileFullName) Then
                Set flExpFile = fso.GetFile(.ExpFileFullName)
            Else
                '~~ The component apparently has yet no never been exported
                If Not Me.VBComp Is Nothing Then
                    Me.VBComp.Export .ExpFileFullName
                    Set flExpFile = fso.GetFile(.ExpFileFullName)
                End If
            End If
        End If
    End With
    
    Set ExpFile = flExpFile
    
End Property

Friend Property Let ExpFile(ByVal fl As File):      Set flExpFile = fl:         End Property

Friend Property Get ExpFileExt() As String:         ExpFileExt = sExpFileExt:   End Property

Friend Property Get ExpFileFullName() As String
' ------------------------------------------------------------------------------
' Will already have been prepared along with 'Property Let CompName'. However,
' when the file does not exist it is created by the Export service.
' ------------------------------------------------------------------------------
    If sExpFileFullName = vbNullString Then
        sExpFileFullName = mExport.ExpFileFolderPath(Me.Wrkbk) & "\" & sCompName & Me.ExpFileExt
    End If
    ExpFileFullName = sExpFileFullName
    With New FileSystemObject
        If Not .FileExists(sExpFileFullName) Then
            Me.VBComp.Export sExpFileFullName
        End If
    End With
End Property

Friend Property Get ExpFilesDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Export File of this Used Common Component differs from
' the Export File of the corresponding Raw Common Component.
' ------------------------------------------------------------------------------
    ExpFilesDiffer = mCompMan.Services.FilesDiffer(f_file_1:=Raw.LastModExpFile _
                                                 , f_file_2:=Me.ExpFile)
End Property

Friend Property Get ExpFileTemp() As File
    With fso
        If Not .FileExists(Me.ExpFileTempFullName) Then
            Me.VBComp.Export Me.ExpFileTempFullName
        End If
        Set ExpFileTemp = .GetFile(Me.ExpFileTempFullName)
    End With
End Property

Friend Property Get ExpFileTempAndRawExpFileDiffers() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the temporary Export File of this Used Common Component
' differs from the Export File of the corresponding Raw Common Component.
' ------------------------------------------------------------------------------
    ExpFileTempAndRawExpFileDiffers = mCompMan.Services.FilesDiffer(f_file_1:=Raw.LastModExpFile _
                                                                  , f_file_2:=Me.ExpFileTemp)
End Property

Friend Property Get ExpFileTempFullName() As String
' ------------------------------------------------------------------------------
' Returns the name for a temporary Export File with the corresponding temporary
' folder already created thereby.
' ------------------------------------------------------------------------------
    Const PROC = "ExpFileTempFullName-Get"
    
    On Error GoTo eh
    If Not wbk Is Nothing Then
        With fso
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

'Friend Property Get ExpFileTempsDiffer() As Boolean
'' ------------------------------------------------------------------------------
'' Returns TRUE when the temporary Export File of this component differs from
'' the temporary Export File of the Raw. This comparison is only possible when
'' the Raw's host Workbook is accidentially open.
'' ------------------------------------------------------------------------------
'    ExpFileTempsDiffer = mCompMan.Services.FilesDiffer(f_file_1:=Raw.ExpFileTemp _
'                                            , f_file_2:=Me.ExpFileTemp)
'End Property

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
        If CommComps.Exists(Me.VBComp, Me.ExpFileExt) Then
            '~~ The component is registered in the ComComps.dat and thus is known as a Common Component,
            '~~ i.e. one of which an Export-File exists in the Common-Components folder.
            If CompManDat.RegistrationState(sCompName) = enRegStateHosted Then
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Let KindOfComp(ByVal en As enKindOfComp)
' ------------------------------------------------------------------------------
' For the "self-set-up" only
' ------------------------------------------------------------------------------
    lKindOfComp = en
End Property

Friend Property Get MaxLenComp() As Long:                   MaxLenComp = lMaxLenComp:   End Property

Friend Property Get Outdated(ByRef o_discontinue As Boolean) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the current component is a Used Common Component of which
' the corresponding Raw Common Component's code has changed.
' ------------------------------------------------------------------------------
    Const PROC = "Outdated"
    
    On Error GoTo eh
    Dim sRevNoCommCompFolder    As String
    Dim bExpFilesDiffer         As Boolean
    Dim Msg                     As TypeMsg
    Dim sRegStat                As String
    mBasic.BoP ErrSrc(PROC)
    With Me
        sRevNoCommCompFolder = CommComps.RevisionNumber(.CompName)
        bExpFilesDiffer = Services.FilesDiffer(f_file_1:=.ExpFile _
                                             , f_file_2:=CommComps.LastModExpFile(.CompName) _
                                             , f_ignore_export_header:=True)
        Select Case True
            Case .RevisionNumber < sRevNoCommCompFolder
                If bExpFilesDiffer Then
                    Outdated = True
                Else
                    Err.Raise AppErr(1), ErrSrc(PROC), _
                              "While the Export-File of the used Common Component is identical with the " & _
                              "Export File in the Common-Components folder the Revision numbers differ! " & _
                              "This is a severe logic error which requires further code inspection and testing."
                End If
            Case .RevisionNumber = sRevNoCommCompFolder
                If bExpFilesDiffer Then
                    '~~ Servere logic error!
                    Err.Raise AppErr(2), ErrSrc(PROC), _
                              "While the Export-File of the used Common Component differs from the " & _
                              "Export File in the Common-Components folder the Revision numbers are equal! " & _
                              "This is a severe logic error which requires further code inspection and testing."
                Else
                End If
            Case .RevisionNumber > sRevNoCommCompFolder
                '~~ The Common-Components folder appears to be outdated
                With Msg.Section(1)
                    With .Text
                        .Text = "Attention!" & vbLf & _
                                "The ""Revision Number"" of the " & CompManDat.RegistrationState(Me.CompName) & _
                                " Common Component   " & mBasic.Spaced(Me.CompName) & "   is greater than " & _
                                "the ""Revision Number"" of the Export-File in the Common-Components folder. " & _
                                "This means that the content of the Common-Components folder is outdated!!!"
                        .FontColor = rgbRed
                    End With
                End With
                Msg.Section(2).Text.Text = "Any check for outedated Common Components will be discontinued!"
                o_discontinue = True
                mMsg.Dsply dsply_title:="The Common-Components folder appears to be outdated!" _
                         , dsply_msg:=Msg _
                         , dsply_buttons:=vbOKOnly
                GoTo xt
        End Select
        If Not Outdated Then
            If .RevisionNumber <> sRevNoCommCompFolder Then
                .RevisionNumber = sRevNoCommCompFolder
            End If
            With Services
                .ServicedItem = Me.VBComp
                .LogServicedEntry "Component is up-to-date! Revision-Number has been set to to " & sRevNoCommCompFolder & " (the raw's in the Common Components Folder)"
            End With
        End If
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

'Private Property Get RawDiffers() As Boolean
'' ----------------------------------------------------------------------------
'' Returns TRUE when the Used Common Component's (ME's) corresponding Raw had
'' changed which is indicated by Raw Revision Number greater than the current
'' Revision Number of the Used Common Component.
'' Note: When the Used Common Component had already been updated manually via
''       the open Raw host Workbook - indicated by the fact that the Export
''       Files are equal - only the Revison Number of the Used Common Component
''       is updated and FALSE is returned.
'' ------------------------------------------------------------------------------
'    Const PROC = "RawDiffers"
'
'    On Error GoTo eh
'    Dim bIgnoreExpFileHeader    As Boolean
'    Dim sRevUsed                As String
'    Dim sRevNoCommCompFolder                 As String
'
'    mBasic.BoP ErrSrc(PROC), Me.CompName
'    sRevUsed = RevisionNumber
'    sRevNoCommCompFolder = Raw.LastModRevisionNumber
'    bIgnoreExpFileHeader = VBComp.Type = vbext_ct_MSForm
'
'    If Services.FilesDiffer(f_file_1:=Me.ExpFile _
'                          , f_file_2:=Raw.LastModExpFile _
'                          , f_ignore_export_header:=bIgnoreExpFileHeader) Then
'        RawDiffers = True
'        '~~ This should enforce the update choice dialog which will consider the
'        '~~ Revision-Number difference.
'    Else
'        '~~ Because the 'used' Common Component is identical with its 'raw' version
'        '~~ in the 'Common Component Folder' the 'used' must have been instantly updated
'        '~~ after the change e.g. by dragging the component from the VB-Project it had
'        '~~ been changed into the other Workbook. Thus, just the Revision Number
'        '~~ needs to be set equal.
'        If sRevUsed <> sRevNoCommCompFolder Then
'            RevisionNumber = sRevNoCommCompFolder
'            With Services
'                .ServicedItem = Me.VBComp
'                .LogServicedEntry "Component is up-to-date! Revision-Number has been set to to " & sRevNoCommCompFolder & " (the raw's in the Common Components Folder)"
'            End With
'        End If
'    End If
'
'xt: mBasic.EoP ErrSrc(PROC), Me.CompName
'    Exit Property
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Property

Friend Property Get RevisionNumber() As String:             RevisionNumber = CompManDat.RevisionNumber(sCompName): End Property

Friend Property Let RevisionNumber(ByVal s As String):      CompManDat.RevisionNumber(sCompName) = s:              End Property

Friend Property Get RevisionNumbersDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the RevisionNumber of the Raw Common Component differs from
' the current RevisionNumber of this Used Common Component.
' ------------------------------------------------------------------------------
    If KindOfComp = enCommCompUsed Then
        RevisionNumbersDiffer = RevisionNumber <> Raw.LastModRevisionNumber
    End If
End Property

Friend Property Get Service() As String
    If sService = vbNullString Then
        sService = fso.GetFileName(Me.Wrkbk.Path)
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
        Me.CompName = vb_comp.Name
    End If
End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    
    If wbk Is Nothing Then
        If Not fso.FileExists(sWrkbkFullName) _
        Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
        Set wbk = mCompMan.WbkGetOpen(sWrkbkFullName)
    Else
        Set Wrkbk = wbk
    End If
    
xt: Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    With fso
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
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
    Dim sRevNo As String
    
    With Me
        .VBComp.Export .ExpFileFullName   ' 1. The 'system' export
        Select Case .KindOfComp
            Case enCommCompHosted, enCommCompUsed
                '~~ The last Workbook in which a Common Component is modified becomes
                '~~ the current "Raw host" regardless of which Workbook claims hosting it.
                CommComps.SaveToCommonComponentsFolder .CompName, .ExpFile
                RevisionNumberIncrease sRevNo
                CommComps.LastModExpFileFullNameOrigin(.CompName) = .ExpFileFullName
                CommComps.LastModWbk(.CompName) = .Wrkbk
                CommComps.RevisionNumber(.CompName) = sRevNo
                CompManDat.RevisionNumber(.CompName) = sRevNo
        End Select
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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

Public Sub RevisionNumberIncrease(ByRef r_rev_no As String)
' ----------------------------------------------------------------------------
' Increases the revision number by one starting with 1 for a new day.
' ----------------------------------------------------------------------------
    Dim lRevNo As Long
    Dim sRevNo As String
    
    sRevNo = CompManDat.RevisionNumber(Me.CompName)
    If sRevNo = vbNullString Then
        sRevNo = Format(Now(), FORMAT_REV_DATE) & "." & Format(1, FORMAT_REV_NO)
    Else
        lRevNo = Split(sRevNo, ".")(1)
        If Split(sRevNo, ".")(0) <> Format(Now(), FORMAT_REV_DATE) Then ' not/nolonger the same day
            lRevNo = 1
        Else
            lRevNo = lRevNo + 1
        End If
        sRevNo = Format(Now(), FORMAT_REV_DATE) & "." & Format(lRevNo, FORMAT_REV_NO)
    End If
    
    r_rev_no = sRevNo
    
End Sub

