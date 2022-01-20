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
' - IsSheet
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
' - DsplyAllChanges
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
Private sTmpExpFileFullName As String
Private sWrkbkBaseName      As String       ' Common Component host Workbook's base name
Private sWrkbkFullName      As String       ' Common Component host Workbook's full name
Private vbc                 As VBComponent  ' VBComponent managed by this Class Module
Private wb                  As Workbook     ' The class' Workbook
Private sService            As String       ' Name of the 'Service' folder, i.e. the folder of the Workbook
Private dctType             As Dictionary
Private lMaxLenComp         As Long
Private sh                  As Worksheet
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
    Set wb = Nothing
    Set vbc = Nothing
    Set flExpFile = Nothing
    
    CleanUpTemps
    
    Set fso = Nothing
    Set Raw = Nothing
    
End Sub

Public Sub RevisionNumberIncrease()
    mComCompsRawsHosted.RawRevisionNumberIncrease (sCompName)
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
    DoEvents
    
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
        Set dctChanges = mService.FilesDifference(fd_exp_file_1:=fso.GetFile(Me.ExpFileTempFullName) _
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
' ----------------------------------------------------------
' Returns the content of an Export-File (Me.ExpFileFullName)
' as array, considering any kind of line break characters.
' ----------------------------------------------------------
    Const PROC = "CodeLines"
    Const ATTRIB_VB = "Attribute VB_"

    On Error GoTo eh
    Dim i   As Long
    Dim j   As Long
    Dim dct As Dictionary
    
    With New FileSystemObject
        Set dct = mFile.Dict(.GetFile(Me.ExpFileFullName))
    End With
    
    '~~ Remove the extra module information lines
    For i = dct.Count - 1 To 0 Step -1
        If Left(dct.Items()(i), Len(ATTRIB_VB)) = ATTRIB_VB Then Exit For
    Next i
    
    For j = 0 To i
'        Debug.Print dct.Items()(0)
        dct.Remove dct.Keys()(0)
    Next j
    '~~ Remove any leading or trailing empty items
    If dct.Count = 0 Then GoTo xt
    While dct.Items()(0) = vbNullString
        dct.Remove dct.Keys()(0)
    Wend
    
'    Debug.Print dct.Items()(0)
    
    ' Remove any trailing empty items
    While dct.Items()(dct.Count - 1) = vbNullString
        dct.Remove dct.Keys()(dct.Count - 1)
    Wend

xt: Set CodeLines = dct
    Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get CodeName( _
              Optional ByVal scn_name As String) As String
' ------------------------------------------------------------------------------
' Returns the sheet's CodeName through its Name.
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet
    For Each wsh In wb.Worksheets
        If sh.Name = scn_name Then
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
    If Not wb Is Nothing Then
        If Me.Exists Then
            Set vbc = wb.VBProject.VBComponents(sCompName)
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
                Case vbext_ct_Document:     sExpFileExt = ".cls"
            End Select
            sExpFileFullName = mExport.ExpFileFolderPath(wb) & "\" & vbc.Name & sExpFileExt
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

Friend Property Get Exists(Optional ByRef ex_wb As Workbook = Nothing) As Boolean
    Dim v As VBComponent
    If Not ex_wb Is Nothing Then
        On Error Resume Next
        Set v = ex_wb.VBProject.VBComponents(sCompName)
        Exists = Not v Is Nothing
    ElseIf vbc Is Nothing Then
        On Error Resume Next
        Set vbc = wb.VBProject.VBComponents(sCompName)
        Exists = Not vbc Is Nothing
    Else
        Exists = Not vbc Is Nothing
    End If
End Property

Friend Property Get ExistsBySheetName(Optional ByVal eb_name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Me.Wrkbk.Worksheets
        ExistsBySheetName = ws.Name = eb_name
        If ExistsBySheetName Then Exit For
    Next ws
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

'Private Property Get ExpFileExt() As String
'
'    If sExpFileExt = vbNullString Then
'        If vbc Is Nothing Then Set vbc = wb.VBProject.VBComponents(sCompName)
'        With Me
'            Select Case vbc.Type
'                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
'                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
'                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
'                Case vbext_ct_Document:     sExpFileExt = ".cls"
'            End Select
'        End With
'    End If
'    ExpFileExt = sExpFileExt
'
'End Property

Friend Property Get ExpFileFullName() As String
' ------------------------------------------------------------------
' Will already have been prepared along with 'Property Let CompName'
' ------------------------------------------------------------------
    If sExpFileFullName = vbNullString Then
        sExpFileFullName = mExport.ExpFileFolderPath(Me.Wrkbk) & "\" & sCompName & Me.ExpFileExt
    End If
    ExpFileFullName = sExpFileFullName
End Property

Friend Property Get ExpFilesDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Export File of this Used Common Component differs from
' the Export File of the corresponding Raw Common Component.
' ------------------------------------------------------------------------------
    ExpFilesDiffer = mService.FilesDiffer(fd_exp_file_1:=Raw.SavedExpFile _
                                        , fd_exp_file_2:=Me.ExpFile)
End Property

Friend Property Get ExpFileTempAndRawExpFileDiffers() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the temporary Export File of this Used Common Component
' differs from the Export File of the corresponding Raw Common Component.
' ------------------------------------------------------------------------------
    ExpFileTempAndRawExpFileDiffers = mService.FilesDiffer(fd_exp_file_1:=Raw.SavedExpFile _
                                                         , fd_exp_file_2:=Me.ExpFileTemp)
End Property

Friend Property Get ExpFileTempsDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the temporary Export File of this component differs from
' the temporary Export File of the Raw. This comparison is only possible when
' the Raw's host Workbook is accidentially open.
' ------------------------------------------------------------------------------
    ExpFileTempsDiffer = mService.FilesDiffer(fd_exp_file_1:=Raw.ExpFileTemp _
                                            , fd_exp_file_2:=Me.ExpFileTemp)
End Property

Friend Property Get RevisionNumbersDiffer() As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the RevisionNumber of the Raw Common Component differs from
' the current RevisionNumber of this Used Common Component.
' ------------------------------------------------------------------------------
    RevisionNumbersDiffer = mComCompsRawsSaved.RawSavedRevisionNumber(sCompName) <> mComCompsUsed.RevisionNumber(sCompName)
End Property



Friend Property Get ExpFileTemp() As File
    Dim fl  As File
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
    If Not wb Is Nothing Then
        With fso
            If sTmpFolder = vbNullString Then
                sTmpFolder = mExport.ExpFileFolderPath(wb) & "\Temp"
            End If
            If Not .FolderExists(sTmpFolder) Then
                DoEvents
                .CreateFolder sTmpFolder
                DoEvents
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
' ------------------------------------------------------------------------------
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    If lKindOfComp = enKindOfComp.enUnknown Then
        If mComCompsRawsSaved.Exists(sCompName) Then
            '~~ The component is registered in the ComComps-Saved.dat as a known Common Component
            If mComCompsRawsSaved.RawHostWbName(comp_name:=sCompName) = Me.Wrkbk.Name Then
                '~~ When the registered host is this Component's Workbook the component is regarded
                '~~ the hosted one else it is a Used Common Component
                lKindOfComp = enCommCompHosted
            Else
                lKindOfComp = enCommCompUsed
                '~~ For a Used Common Component a class object for the corresponding Raw is established
                '~~ which provides all required properties
                If Raw Is Nothing Then Set Raw = New clsRaw
                Raw.CompName = sCompName
                '~~ all the Raw's other properties are derived based on its name through the data
                '~~ stored in the ComCompsSaved.dat file
            End If
        ElseIf lKindOfComp = enKindOfComp.enUnknown Then
            lKindOfComp = enInternal
        End If
'        Log.Entry = Me.TypeString & " (" & KoCStrng(lKindOfComp) & ")"
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
    With Me
        If .KindOfComp = enCommCompUsed Then Outdated = .RawChanged
    End With
End Property

Public Property Get RawChanged() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the current Used Common Component's corresponding Raw had
' changed which is indicated by Raw Revision Number greater than the current
' Revision Number of the Used Common Component.
' Note: When the Used Common Component had already been updated manually via
'       the open Raw host Workbook - indicated by the fact that the Export
'       Files are equal - only the Revison Number of the Used Common Component
'       is updated and FALSE is returned.
' ------------------------------------------------------------------------------
    Const PROC = "RawChanged"
    
    On Error GoTo eh
    If Me.RevisionNumber < mComCompsRawsSaved.RawSavedRevisionNumber(sCompName) Then
        If Not mService.FilesDiffer(fd_exp_file_1:=Me.ExpFile _
                              , fd_exp_file_2:=Raw.SavedExpFile) Then
            '~~ Propably updated already manually 'antidated'
            mComCompsUsed.RevisionNumber(sCompName) = mComCompsRawsSaved.RawSavedRevisionNumber(sCompName)
        Else
            
            RawChanged = True
        End If
    ElseIf Me.RevisionNumber = mComCompsRawsSaved.RawSavedRevisionNumber(sCompName) Then
        If mService.FilesDiffer(fd_exp_file_1:=Me.ExpFile _
                              , fd_exp_file_2:=Raw.SavedExpFile) Then
            '~~ it is not the Raw which had changed but the Used Common Component in its Workbook.
            '~~ In order to prevent Used Common Components differ between Workbooks using them
            '~~ and remainin different unrecognized the Raw is indicated changed to enforce
            '~~ an update - and a revert of the modification! - with the next Workbook open
            RawChanged = True
        End If
    End If

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get RevisionNumber() As String:         RevisionNumber = mComCompsUsed.RevisionNumber(sCompName): End Property

Friend Property Let RevisionNumber(ByVal s As String):  mComCompsUsed.RevisionNumber(sCompName) = s:              End Property

Friend Property Get Service(Optional ByVal wb_full_name) As String
    If sService = vbNullString Then
        sService = fso.GetFileName(Me.Wrkbk.Path)
    End If
    Service = sService
End Property

Friend Property Get sheet() As Worksheet:               Set sheet = sh:            End Property

Friend Property Set sheet(ByRef wsh As Worksheet):      Set sh = wsh:               End Property

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

Friend Property Get VBComp() As VBComponent:    Set VBComp = vbc:   End Property

Friend Property Set VBComp(ByRef vb_comp As VBComponent):       Set vbc = vb_comp:                      End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    
    If wb Is Nothing Then
        If Not fso.FileExists(sWrkbkFullName) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
        Set wb = mCompMan.WbkGetOpen(sWrkbkFullName)
    Else
        Set Wrkbk = wb
    End If
    
xt: Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Set Wrkbk(ByRef wbk As Workbook)
' ---------------------------------------------------------
' Provides the properties: - Wrkbk         (wb)
'                          - WrkbFullName  (sWrkbkFullName)
'                          - WrkbkBaseName (sWrkbkBaseName)
'                          - ExpFilePath   (sExpFilePath)
'                          - Service       (sService)
'                          - MaxLenComp    (lMaxLenComp)
' ---------------------------------------------------------
    
    Dim vbc As VBComponent
    
    Set wb = wbk
    sWrkbkFullName = wb.FullName
    With fso
        sWrkbkBaseName = .GetBaseName(sWrkbkFullName)
        sExpFilePath = .GetParentFolderName(sWrkbkFullName)
        sService = .GetFileName(wb.Path)
    End With
    For Each vbc In wb.VBProject.VBComponents
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

Public Function CompType() As Dictionary
    Set CompType = dctType
End Function

Public Sub CopyExportFileToCommonComponentsFolder()
' ------------------------------------------------------------------------------
' Maintain a copy of the hosted raw`s export file in the Common Components
' folder as the source the update of Common Components used in VB-Projects.
' ------------------------------------------------------------------------------
    Dim frxFileFullName As String
    Dim frxFile         As File
    
    With Me
        mComCompsRawsSaved.RawExpFileFullName(sCompName) = .ExpFileFullName
        mComCompsRawsSaved.RawSavedRevisionNumber(sCompName) = mComCompsRawsHosted.RawRevisionNumber(sCompName)
        mComCompsRawsSaved.SavedExpFile(sCompName) = .ExpFile
    
        '~~ When the Export file has a .frm extension the .frx file neds to be copied too
        If fso.GetExtensionName(.ExpFileFullName) = "frm" Then
            frxFileFullName = Replace(.ExpFileFullName, "frm", "frx")
            Set frxFile = fso.GetFile(frxFileFullName)
            mComCompsRawsSaved.SavedExpFile(sCompName) = frxFile
        End If
    End With
End Sub

Private Function DebugComp() As String
    If lMaxCompLength > 0 _
    Then DebugComp = sCompName & " " & String(lMaxCompLength - Len(sCompName), ".") & ": " _
    Else DebugComp = sCompName & " : "
End Function

Public Sub DsplyAllChanges()
' ------------------------------------------------------------------------------
' Displays the code differences between a clone's and the raw's Export-File.
' ------------------------------------------------------------------------------
    Const PROC = "DsplyAllChanges"
    
    On Error GoTo eh
    Dim sMsg As TypeMsg
    
    If Me.Changed Then
        mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=sTmpExpFileFullName _
                                   , fd_exp_file_right_full_name:=Me.ExpFileFullName _
                                   , fd_exp_file_left_title:="Clone (used) Common Component: '" & sTmpExpFileFullName & "'" _
                                   , fd_exp_file_right_title:="Raw (hosted) Common Component: '" & Me.ExpFileFullName & "'"
    Else
        sMsg.Section(1).Text.Text = "The Export Files of the component '" & Me.CompName & "'"
        sMsg.Section(2).Text.Text = "Temp Export:    " & sTmpExpFileFullName & vbLf & _
                                "Current Exüort: " & Me.ExpFileFullName
        sMsg.Section(2).Text.MonoSpaced = True
        sMsg.Section(3).Text.Text = "are identical. The clone component has thus not changed."
        sMsg.Section(4).Label.Text = "Please note!"
        sMsg.Section(4).Text.Text = "Empty code lines and case differences are ignored when they are the only difference."
        
        mMsg.Dsply dsply_title:="Display of code changes failed!" _
                 , dsply_msg:=sMsg
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
End Function

Public Sub Export()
    Me.VBComp.Export Me.ExpFileFullName
End Sub

Public Sub IncreaseRevisionNumber()
' ------------------------------------------------------------------------------
' Increase of the revision number when the hosted Raw Common Component had
' been exported because it had changed.
' ------------------------------------------------------------------------------
    mComCompsRawsHosted.RawRevisionNumberIncrease Me.CompName
End Sub

Public Function KoCStrng(ByVal koc As enKindOfComp) As String
    Select Case koc
        Case enKindOfComp.enUnknown:        KoCStrng = "Unknown"
        Case enKindOfComp.enInternal:       KoCStrng = "Internal, neither Raw nor Clone"
        Case enKindOfComp.enCommCompUsed:       KoCStrng = "Clone"
        Case enKindOfComp.enCommCompHosted:      KoCStrng = "Raw (Hosted)"
    End Select
End Function

Public Function RemoteSheetExists( _
                            ByRef rs_wb As Workbook, _
                   Optional ByRef rs_ws_name As String = vbNullString, _
                   Optional ByRef rs_ws_codename As String = vbNullString, _
                   Optional ByRef rs_ws As Worksheet = Nothing, _
                   Optional ByRef rs_vbc As VBComponent = Nothing) As Boolean
' ---------------------------------------------------------------------------
' When the remote sheet exists in the remote Workbook (rs_wb) either under
' the provided remote name (remote_ws_name) or the provided remote CodeName
' (remote-ws_codename):
' - True plus the remote Worksheet object (rs_ws) and the remote VBComponent
'   (rs_vbc).
' When the remote sheet not exists in the remote Workbook (rs_wb) neither
' under the provided remote name (remote_ws_name) nor under the provided
' remote CodeName (remote-ws_codename):
' - False and neither a remote sheet object nor a remote VBComponent.
' ---------------------------------------------------------------------------
    Dim ws  As Worksheet
    Dim vbc As VBComponent
    
    For Each ws In rs_wb.Worksheets
        If rs_ws_name <> vbNullString And ws.Name = rs_ws_name Then
            Set rs_ws = ws
            If rs_ws_codename = vbNullString Then rs_ws_codename = ws.CodeName
            Exit For
        ElseIf rs_ws_codename <> vbNullString And ws.CodeName = rs_ws_codename Then
            Set rs_ws = ws
            If rs_ws_name = vbNullString Then rs_ws_name = ws.Name
        End If
    Next ws
    
    For Each vbc In rs_wb.VBProject.VBComponents
        If rs_ws_codename <> vbNullString And vbc.Name = rs_ws_codename Then
            Set rs_vbc = vbc
            Exit For
        End If
    Next vbc
   
End Function



