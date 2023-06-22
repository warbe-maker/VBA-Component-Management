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

Public Property Get DueModificationWarning() As Boolean
    DueModificationWarning = mCompManDat.DueModificationWarning(sCompName)
End Property

Public Property Let DueModificationWarning(ByVal b As Boolean)
    mCompManDat.DueModificationWarning(sCompName) = b
End Property

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
        Set dctChanges = mCompMan.Services.FilesDifference(fd_exp_file_1:=fso.GetFile(Me.ExpFileTempFullName) _
                                                , fd_exp_file_2:=fso.GetFile(Me.ExpFileFullName))
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
    sCompName = s
    If Not wbk Is Nothing Then
        If Me.Exists Then
            Set vbc = wbk.VBProject.VBComponents(sCompName)
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
                Case vbext_ct_Document:     sExpFileExt = ".cls"
            End Select
            sExpFileFullName = mExport.ExpFileFolderPath(wbk) & "\" & vbc.Name & sExpFileExt
        End If
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "Component Name assigned for a yet unknown Workbook!"
    End If

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get Exists(Optional ByRef ex_wbk As Workbook = Nothing) As Boolean
    Dim v As VBComponent
    If Not ex_wbk Is Nothing Then
        On Error Resume Next
        Set v = ex_wbk.VBProject.VBComponents(sCompName)
        Exists = Not v Is Nothing
    ElseIf vbc Is Nothing Then
        On Error Resume Next
        Set vbc = wbk.VBProject.VBComponents(sCompName)
        Exists = Not vbc Is Nothing
    Else
        Exists = Not vbc Is Nothing
    End If
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

Friend Property Set ExpFile(ByVal fl As File):      Set flExpFile = fl:     End Property

Friend Property Get ExpFileExt() As String: ExpFileExt = sExpFileExt:   End Property

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
    ExpFilesDiffer = mCompMan.Services.FilesDiffer(fd_exp_file_1:=Raw.SavedExpFile _
                                                , fd_exp_file_2:=Me.ExpFile)
End Property

Friend Property Get ExpFileTempAndRawExpFileDiffers() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the temporary Export File of this Used Common Component
' differs from the Export File of the corresponding Raw Common Component.
' ------------------------------------------------------------------------------
    ExpFileTempAndRawExpFileDiffers = mCompMan.Services.FilesDiffer(fd_exp_file_1:=Raw.SavedExpFile _
                                                                 , fd_exp_file_2:=Me.ExpFileTemp)
End Property

Friend Property Get ExpFileTempsDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the temporary Export File of this component differs from
' the temporary Export File of the Raw. This comparison is only possible when
' the Raw's host Workbook is accidentially open.
' ------------------------------------------------------------------------------
    ExpFileTempsDiffer = mCompMan.Services.FilesDiffer(fd_exp_file_1:=Raw.ExpFileTemp _
                                            , fd_exp_file_2:=Me.ExpFileTemp)
End Property

Friend Property Get RevisionNumbersDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the RevisionNumber of the Raw Common Component differs from
' the current RevisionNumber of this Used Common Component.
' ------------------------------------------------------------------------------
    RevisionNumbersDiffer = mCommComps.RevisionNumber(sCompName) <> mCompManDat.RevisionNumber(sCompName)
End Property

Friend Property Get ExpFileTemp() As File
    With fso
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
        Err.Raise AppErr(1), ErrSrc(PROC), "Workbook object yet not available for component '" & sCompName & "'!"
    End If
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

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
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    If lKindOfComp = mCompMan.enUnknown Then
        If mCommComps.ExistsRegistered(sCompName) Then
            '~~ The component is registered in the ComComps-Saved.dat as a known Common Component
            If mCommComps.RawHostWbName(comp_name:=sCompName) = Me.Wrkbk.Name Then
                '~~ When the registered host is this Component's Workbook the component is regarded
                '~~ the hosted one else it is a Used Common Component
                lKindOfComp = mCompMan.enCommCompHosted
            Else
                If Not mCompManDat.RegistrationState(Me.CompName) = enRegStatePrivate Then
                    lKindOfComp = mCompMan.enCommCompUsed
                    '~~ For a Used Common Component a class object for the corresponding Raw is established
                    '~~ which provides all required properties
                    If Raw Is Nothing Then Set Raw = New clsRaw
                    Raw.CompName = sCompName
                    '~~ all the Raw's other properties are derived based on its name through the data
                    '~~ stored in the ComCompsSaved.dat file
                End If
            End If
        ElseIf lKindOfComp = mCompMan.enUnknown Then
            lKindOfComp = mCompMan.enInternal
        End If
'        Srvc.LogEntry = Me.TypeString & " (" & KoCStrng(lKindOfComp) & ")"
    End If
    
xt: KindOfComp = lKindOfComp
    Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get MaxLenComp() As Long:                   MaxLenComp = lMaxLenComp:   End Property

Friend Property Get Outdated() As Boolean
' ------------------------------------------------------------------------------
' Returns True when the current component is a Used Common Component of which
' the corresponding Raw Common Component's code has changed.
' ------------------------------------------------------------------------------
    Const PROC = "Outdated"
    
    mBasic.BoP ErrSrc(PROC)
    With Me
        If .KindOfComp = mCompMan.enCommCompUsed Then Outdated = RawChanged
    End With
    mBasic.EoP ErrSrc(PROC)
    
End Property

Private Property Get RawChanged() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Used Common Component's (ME's) corresponding Raw had
' changed which is indicated by Raw Revision Number greater than the current
' Revision Number of the Used Common Component.
' Note: When the Used Common Component had already been updated manually via
'       the open Raw host Workbook - indicated by the fact that the Export
'       Files are equal - only the Revison Number of the Used Common Component
'       is updated and FALSE is returned.
' ------------------------------------------------------------------------------
    Const PROC = "RawChanged"
    
    On Error GoTo eh
    Dim bExcludeExportHeader As Boolean
    
    mBasic.BoP ErrSrc(PROC), Me.CompName
    bExcludeExportHeader = Me.TypeString = "UserForm"
    If Me.RevisionNumber < mCommComps.RevisionNumber(sCompName) Then
        '~~ The Revision Number of the Raw Common Component is greater which means that
        '~~ a modification had been exported - and the Revision Number thereby increased
        If Not mCompMan.Services.FilesDiffer(fd_exp_file_1:=Me.ExpFile _
                                          , fd_exp_file_2:=Raw.SavedExpFile _
                                          , fd_ignore_export_header:=bExcludeExportHeader) Then
            '~~ Because the Export File of the Used Common Component is identical
            '~~ with the Export File of the Raw Common Component the Used Common
            '~~ Component must have been updated manually already - e.g. by dragging
            '~~ the component from the hosting Workbook. So just the Revision Number
            '~~ needs to be set equal.
            mCompManDat.RevisionNumber(sCompName) = mCommComps.RevisionNumber(sCompName)
        Else
            RawChanged = True
        End If
    ElseIf Me.RevisionNumber = mCommComps.RevisionNumber(sCompName) Then
        If mCompMan.Services.FilesDiffer(fd_exp_file_1:=Me.ExpFile _
                                      , fd_exp_file_2:=Raw.SavedExpFile _
                                      , fd_ignore_export_header:=bExcludeExportHeader) Then
            '~~ it is not the Raw which had changed but the Used Common Component in its Workbook.
            '~~ In order to prevent Used Common Components differ between Workbooks using them
            '~~ and remaining different unrecognized the Raw is indicated changed to enforce
            '~~ an update - and by the way revert the modification in the Used Common Component.
            RawChanged = True
        End If
    End If

xt: mBasic.EoP ErrSrc(PROC), Me.CompName
    Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get RevisionNumber() As String:         RevisionNumber = mCompManDat.RevisionNumber(sCompName): End Property

Friend Property Let RevisionNumber(ByVal s As String):  mCompManDat.RevisionNumber(sCompName) = s:              End Property

Friend Property Get Service() As String
    If sService = vbNullString Then
        sService = fso.GetFileName(Me.Wrkbk.Path)
    End If
    Service = sService
End Property

Friend Property Get Sheet() As Worksheet:                   Set Sheet = wsh:            End Property

Friend Property Set Sheet(ByRef s_wsh As Worksheet):        Set wsh = s_wsh:            End Property

Friend Property Get TypeString() As String
' ----------------------------------------
' Returns the component's type as String.
' ----------------------------------------
    If Not vbc Is Nothing Then
        On Error Resume Next
        TypeString = dctType(vbc.Type)
        If TypeString = vbNullString Then TypeString = "unknown"
    End If
End Property

Friend Property Get VBComp() As VBComponent:                Set VBComp = vbc:           End Property

Friend Property Set VBComp(ByRef vb_comp As VBComponent)
    Set vbc = vb_comp
    If sCompName = vbNullString Then Me.CompName = vb_comp.Name
End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    
    If wbk Is Nothing Then
        If Not fso.FileExists(sWrkbkFullName) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
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

Friend Property Set Wrkbk(ByRef w_wbk As Workbook)
' ---------------------------------------------------------
' Provides the properties: - Wrkbk         (wbk)
'                          - WrkbFullName  (sWrkbkFullName)
'                          - WrkbkBaseName (sWrkbkBaseName)
'                          - ExpFilePath   (sExpFilePath)
'                          - Service       (sService)
'                          - MaxLenComp    (lMaxLenComp)
' ---------------------------------------------------------
    
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
    
End Property

Friend Property Get WrkbkBaseName() As String:                 WrkbkBaseName = BaseName(sWrkbkFullName):        End Property

Friend Property Get WrkbkFullName() As String:                 WrkbkFullName = sWrkbkFullName:                  End Property

Friend Property Let WrkbkFullName(ByVal s As String):          sWrkbkFullName = s:                              End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

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
    Const PROC = "Export"
    On Error GoTo eh
    
    Me.VBComp.Export Me.ExpFileFullName

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub



