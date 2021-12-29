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
' - ExpFileFullName
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
Private sTmpFolder          As String
Private sTmpExpFileFullName As String
Private sWrkbkBaseName      As String       ' Common Component host Workbook's base name
Private sWrkbkFullName      As String       ' Common Component host Workbook's full name
Private vbc                 As VBComponent  ' VBComponent managed by this Class Module
Private wb                  As Workbook     ' The class' Workbook
Private sService            As String       ' Name of the 'Service' folder, i.e. the folder of the Workbook
Private dctType             As Dictionary
Private lMaxLenComp         As Long
Private sh                  As Worksheet

Private sRawHostFullName    As String       ' -------------------------------------------------------------------
Private sRawExpFileExt      As String       ' Corresponding Raw-Component facts when the Component is a Raw-Clone
Private sRawExpFileFullName As String       ' -------------------------------------------------------------------
Private fso                 As FileSystemObject

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
    With New FileSystemObject
        If .FolderExists(sTmpFolder) Then .DeleteFolder sTmpFolder
    End With
    Set flExpFile = Nothing
    
    If Me.ExpFileTempFullName <> vbNullString Then
        If fso.FileExists(Me.ExpFileTempFullName) Then
            fso.DeleteFile Me.ExpFileTempFullName
        End If
    End If
    Set fso = Nothing
    
End Sub

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

Friend Property Get Changed() As Boolean
' ------------------------------------------------------------
' Returns TRUE when the Component's Export-File dose not exist
' or a temporary Export-File differs from the existing one.
' ------------------------------------------------------------
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
        '~~ is only provided oncce
        Set dctChanges = mExport.FilesDifference(fd_exp_file_1:=fso.GetFile(Me.ExpFileTempFullName) _
                                               , fd_exp_file_2:=fso.GetFile(Me.ExpFileFullName))
    End If
    Changed = dctChanges.Count <> 0
    
xt: Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get ExpFileTempFullName() As String
    Dim sTmpFolder  As String
    
    With fso
        sTmpFolder = mCompMan.ExpFileFolderPath(Me.WrkbkFullName) & "\Temp"
        If Not .FolderExists(sTmpFolder) Then
            DoEvents
            .CreateFolder sTmpFolder
            DoEvents
        End If
        ExpFileTempFullName = sTmpFolder & "\" & Me.CompName & Me.ExpFileExt
    End With
    
End Property

Friend Property Get CodeName( _
              Optional ByVal scn_name As String) As String
' --------------------------------------------------------
' Returns the sheet's CodeName through its Name.
' --------------------------------------------------------
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
' ---------------------------------------------
' Provides the properties
' - CompName        (sCompName)
' - ExpFileFullName (sExpFileFullName)
' - VBComp          (vbc)
' ---------------------------------------------
    sCompName = s
    If Not wb Is Nothing Then
        If Me.Exists Then
            Set vbc = wb.VBProject.VBComponents(s)
            sExpFileFullName = mCompMan.ExpFileFolderPath(wb) & "\" & sCompName & Me.ExpFileExt
        End If
    End If
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
                Set flExpFile = GetFile(.ExpFileFullName)
            Else
                '~~ The component apparently has yet no never been exported
                If Not Me.VBComp Is Nothing Then
                    Me.VBComp.Export .ExpFileFullName
                    Set flExpFile = GetFile(.ExpFileFullName)
                End If
            End If
        End If
    End With
    
    Set ExpFile = flExpFile
    
End Property

Friend Property Set ExpFile(ByVal fl As File):      Set flExpFile = fl:     End Property

Friend Property Get ExpFileExt() As String
    
    If sExpFileExt = vbNullString Then
        If vbc Is Nothing Then Set vbc = wb.VBProject.VBComponents(sCompName)
        With Me
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
                Case vbext_ct_Document:     sExpFileExt = ".cls"
            End Select
        End With
    End If
    ExpFileExt = sExpFileExt

End Property

Friend Property Get ExpFileFullName() As String
' ------------------------------------------------------------------
' Will already have been prepared along with 'Property Let CompName'
' ------------------------------------------------------------------
    If sExpFileFullName = vbNullString Then
        sExpFileFullName = mCompMan.ExpFileFolderPath(Me.Wrkbk) & "\" & Me.CompName & Me.ExpFileExt
    End If
    ExpFileFullName = sExpFileFullName
End Property

Friend Property Get IsWrkbk() As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbk = Err.Number = 0
    
End Property

Friend Property Get KindOfComp() As enKindOfComp
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    If lKindOfComp = enKindOfComp.enUnknown Then
        If mComCompsSaved.Exists(sCompName) Then
            If mComCompsSaved.HostFullName(comp_name:=sCompName) = Me.Wrkbk.FullName _
            Then lKindOfComp = enCommCompHosted _
            Else lKindOfComp = enCommCompUsed
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

Public Property Let RawExpFileExt(ByVal s As String):       sRawExpFileExt = s:         End Property

Public Property Get SavedExpFileFullName() As String
    SavedExpFileFullName = mComCompsSaved.ExpFileFullName(Me.CompName)
End Property

Public Property Get SavedExpFile() As File
    With New FileSystemObject
        Set SavedExpFile = .GetFile(mComCompsSaved.ExpFileFullName(Me.CompName))
    End With
End Property

Public Property Let RawExpFileFullName(ByVal s As String):  sRawExpFileFullName = s:    End Property

Public Property Get RawMostRecentExpFile() As File
' ------------------------------------------------------------------------------
' Returns the most recent raw's export file. This may be the one saved in the
' Common Components folder or - in case it exists the export folder of the
' Workbook which hosts the Common Component. It has to be considered that even
' when the hosting Workbook exists it may be outdated compared with the saved
' one. The means to decide this is the Revision Number. When the hosted raw's
' export file exists and its revision number is equal with the saved one any
' may be returned. When the saved one's revision date is more up-to-date this
' one is returned
' ------------------------------------------------------------------------------
    With New FileSystemObject
        
        If .FileExists(mComCompsSaved.ExpFileFullName(Me.CompName)) _
        And .FileExists(mComCompsHosted.ExpFileFullName(Me.CompName)) Then
            If mComCompsSaved.RevisionNumber(Me.CompName) = mComCompsHosted.RevisionNumber(Me.CompName) Then
                Set RawMostRecentExpFile = .GetFile(mComCompsHosted.ExpFileFullName(Me.CompName))
            ElseIf mComCompsSaved.RevisionNumber(Me.CompName) > mComCompsHosted.RevisionNumber(Me.CompName) Then
                Set RawMostRecentExpFile = .GetFile(mComCompsSaved.ExpFileFullName(Me.CompName))
            End If
        
        ElseIf Not .FileExists(mComCompsSaved.ExpFileFullName(Me.CompName)) _
        And .FileExists(mComCompsHosted.ExpFileFullName(Me.CompName)) Then
            Set RawMostRecentExpFile = .GetFile(mComCompsHosted.ExpFileFullName(Me.CompName))
        
        ElseIf .FileExists(mComCompsSaved.ExpFileFullName(Me.CompName)) _
        And Not .FileExists(mComCompsHosted.ExpFileFullName(Me.CompName)) Then
            Set RawMostRecentExpFile = .GetFile(mComCompsSaved.ExpFileFullName(Me.CompName))
        End If
        
    End With
End Property

Public Property Get RawMostRecentRevisionNumber() As String
' ------------------------------------------------------------------------------
' Returns the most recent raw's Revision Number. When the most recent
' modification had yet not been exported this will be a - yet still outdated -
' Revision Number of the last saved Export File savede in the Common Components
' folder.
' ------------------------------------------------------------------------------
    With New FileSystemObject
        
        If .FileExists(mComCompsSaved.ExpFileFullName(Me.CompName)) _
        And .FileExists(mComCompsHosted.ExpFileFullName(Me.CompName)) Then
            If mComCompsSaved.RevisionNumber(Me.CompName) = mComCompsHosted.RevisionNumber(Me.CompName) _
            Or mComCompsSaved.RevisionNumber(Me.CompName) > mComCompsHosted.RevisionNumber(Me.CompName) Then
                RawMostRecentRevisionNumber = mComCompsSaved.RevisionNumber(Me.CompName)
            Else
                '~~ May be the case when the Common Components folder(/repo) on this computer is outdated.
                '~~ A pretty specific and thus propably rare case when a versioning system like Github is used:
                '~~ The repo clone of the hosting Workbook is up-to-date but the repo clone of the
                '~~ Common Components (represented via the same named folder) is yet not updated (fetched).
                RawMostRecentRevisionNumber = mComCompsHosted.RevisionNumber(Me.CompName)
            End If
        
        ElseIf Not .FileExists(mComCompsSaved.ExpFileFullName(Me.CompName)) _
        And .FileExists(mComCompsHosted.ExpFileFullName(Me.CompName)) Then
            '~~ May be the case when the Common Components folder(/repo) on this computer is outdated.
            '~~ A pretty specific and thus propably rare case when a versioning system like Github is used:
            '~~ The repo clone of the hosting Workbook is up-to-date but the repo clone of the
            '~~ Common Components (represented via the same named folder) is yet not updated (fetched).
            RawMostRecentRevisionNumber = mComCompsHosted.RevisionNumber(Me.CompName)
        
        ElseIf .FileExists(mComCompsSaved.ExpFileFullName(Me.CompName)) _
        And Not .FileExists(mComCompsHosted.ExpFileFullName(Me.CompName)) Then
            '~~ May be the case when the Common Components folder(/repo) on this computer is available
            '~~ but the Common Component's hosting Workbook is not. When Github is used it is yet not
            '~~ Forked/Cloned on this computer.
            RawMostRecentRevisionNumber = mComCompsSaved.RevisionNumber(Me.CompName)
        End If
        
    End With
End Property



Public Property Let RawHostFullName(ByVal s As String):     sRawHostFullName = s:       End Property

Private Property Get ReplaceRawWithCloneConfirmed(ByRef raw As clsRaw) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' ------------------------------------------------------------------------------
    Const PROC = "ReplaceRawWithCloneConfirmed"
    
    On Error GoTo eh
    Dim sReplyUpdateOrigin      As String
    Dim sReplyDsplyDiff         As String
    Dim sReplyIgnore            As String
    Dim sMsg                    As TypeMsg
    Dim cllButtons              As Collection
    Dim sTitle                  As String
    Dim sReply                  As String
    
    '~~ Prepare communication message and reply option
    sReplyUpdateOrigin = "Update the origin code:" & vbLf & vbLf & _
                         "Replace (remove and re-import) the component " & vbLf & _
                         "" & _
                         "in Workbook " & vbLf & _
                         "" & vbLf & _
                         " with this component's Export-File. I.e. the changes" & vbLf & _
                         "made will become ""permanent"""
    sReplyDsplyDiff = "Display the difference:" & vbLf & vbLf & _
                      "The installed ""WinMerge"" is used. Empty code" & vbLf & _
                      "lines are ignored and the compare is case-insensitive."
    sReplyIgnore = "Ignore the code change:" & vbLf & vbLf & _
                   "The change will be overwritten by the" & vbLf & _
                   "next open of this Workbook."
    mMsg.Buttons cllButtons, sReplyUpdateOrigin, sReplyDsplyDiff, sReplyIgnore
    
    '~~ Offer options
    sTitle = "A ""Remotely hosted raw's clone has been modified in this VB-Project!"
    With sMsg
        .Section(1).Label.Text = "About this incident:"
        .Section(1).Text.Text = "The module '" & Me.CompName & "' is a clone of the raw component hosted in Workbook '" & raw.WrkbkBaseName & "' " & _
                            "This clone component's code has usually has been synced with the remote raw componet's code. " & _
                            "The remote raw hasn't been changed but the cloned component's code has."
        .Section(2).Label.Text = "About the choices:"
        .Section(2).Text.Text = "When the code modification is ignored it will be regarded temporarily only and overwrtitten with the next open of this Workbook. " & _
                            "When the remote The module, updated when this Workbook has opened, is no longer identical with its origin. " & _
                            "I.e. the component's code has been changed in this Workbook using it instead of in the " & _
                            "Workbook the origin is ""hosted""."
        .Section(3).Text.Text = "Take your decission or postpone it after having displayed the difference."
    End With
        
    sReply = sReplyDsplyDiff
    Do
        sReply = mMsg.Dsply(dsply_title:=sTitle, _
                            dsply_msg:=sMsg, _
                            dsply_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sReplyUpdateOrigin
                ReplaceRawWithCloneConfirmed = True
                Exit Do
            Case sReplyDsplyDiff
                mExport.FilesDifferencesDisplay fd_exp_file_left_full_name:=Me.ExpFileFullName _
                                              , fd_exp_file_right_full_name:=raw.ExpFileFullName _
                                              , fd_exp_file_left_title:="The used Common Component's current code in '" & Me.ExpFileFullName & "'" _
                                              , fd_exp_file_right_title:="The raw Common Component's current code in '" & raw.ExpFileFullName & "'"
                ' The question is re-displayed for a decison
            Case sReplyIgnore
                ReplaceRawWithCloneConfirmed = False
                Exit Do
        End Select
    Loop
                                            
xt: Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

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

Private Function BaseName(ByVal s As String) As String
    With New FileSystemObject:  BaseName = .GetBaseName(s): End With
End Function

Public Function CompType() As Dictionary
    Set CompType = dctType
End Function

Private Function DebugComp() As String
    If lMaxCompLength > 0 _
    Then DebugComp = sCompName & " " & String(lMaxCompLength - Len(sCompName), ".") & ": " _
    Else DebugComp = sCompName & " : "
End Function

Public Sub DsplyAllChanges()
' -------------------------------------------
' Display of all differences between the
' clone's and the raw's Export-File.
' -------------------------------------------
    Const PROC = "DsplyAllChanges"
    
    On Error GoTo eh
    Dim sMsg As TypeMsg
    
    If Me.Changed Then
        mExport.FilesDifferencesDisplay fd_exp_file_left_full_name:=sTmpExpFileFullName _
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

Public Sub ReplaceRawWithClone(ByRef raw As clsRaw)
' --------------------------------------------------------------
' Replace the raw component's code with the Export-File of the
' clone component.
' --------------------------------------------------------------
    Const PROC = "ReplaceRawWithClone"

    On Error GoTo eh
    Dim wbRawHost   As Workbook
        
    Application.EnableEvents = False ' prevent the 'UpdateComCompsUsed' service when the Workbook is opened
    Set wbRawHost = mCompMan.WbkGetOpen(raw.HostWrkbkFullName) ' Wrkbk is the Workbook hosting the raw Component
    Application.EnableEvents = True
    
    mRenew.ByImport rn_wb:=wbRawHost _
                  , rn_comp_name:=sCompName _
                  , rn_exp_file_full_name:=Me.ExpFileFullName
    Log.Entry = "Raw-Component (hosted in '" & wbRawHost.Name & "' replaced by (re-)import of the clone's export file '" & Me.ExpFileFullName & "'"

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ReplaceRawWithCloneWhenConfirmed( _
                                      ByRef raw As clsRaw, _
                                      ByRef rwu_updated As Boolean)
' -------------------------------------------------------------------
' Provided confirmed by the user, replaces the raw code in the
'  hosting Workbook with the cloned raw component's code in its
' Export-File. Code replacement for class modules, satandard modules
' and UserForms is done via a rename, re-import, remove renamed
' procedure, for a data module the code is replaced line by line
' because a (re-)import is impossible for such a module.
' -------------------------------------------------------------------
    Const PROC = "ReplaceRawWithCloneWhenConfirmed"
    
    On Error GoTo eh
    
    If ReplaceRawWithCloneConfirmed(raw) Then
            '~~ Get the origin host Workbook open
        With raw
            .HostWrkbkFullName = mComCompsSaved.HostFullName(Me.CompName)
            .CompName = Me.CompName
            .ExpFileExt = Me.ExpFileExt ' required to build the export file's full name
            Set .ExpFile = fso.GetFile(.ExpFileFullName)
            Set .Wrkbk = mCompMan.WbkGetOpen(.HostWrkbkFullName)
            Set .VBComp = .Wrkbk.VBProject.VBComponents(.CompName)
            .TypeString = Me.TypeString
            Select Case .VBComp.Type
                Case vbext_ct_Document
                    '~~ A data module cannot be replaced like the other ones. Thus only its code
                    '~~ is synchronized - disregarding any consequences this may have when the
                    '~~ structure and design of the underlying Workbook or Worksheet has changed.
                    mSync.ByCodeLines sync_target_comp_name:=.CompName _
                                    , wb_source_full_name:=.HostWrkbkFullName _
                                    , sync_source_codelines:=.CodeLines
                Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm
                    Me.ReplaceRawWithClone raw
            End Select
            rwu_updated = True
        End With ' RawComp
    End If

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

