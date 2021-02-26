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
' Properties:
' - ExpFile      R/W File object created by menas of the BackUpCode method
' - HostedRaws  W   String with named Components hosted in the Workbook
' - IsRawHostedElsewhere           R   TRUE when a Component is regarded a Common one
' - CodeChanged     R   TRUE when the code of the CodeModule differs from the code
'                       backed up in the corresponding ExportFile
' - KindOfChange    R
' - KindOfComp      R
'
' Methods:
' - BackUpCode
'
'
' -----------------------------------------------------------------------------------
Private dctChanges          As Dictionary   ' Dictionary of code changes
Private flExpFile           As File
Private foExport            As Folder
Private lKindOfComp         As enKindOfComp
Private sCompName           As String       ' The class module's global component name
Private sExpFileExtension   As String
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
Private lMaxLenType         As Long
Private lMaxLenComp         As Long

Public Property Get MaxLenComp() As Long:           MaxLenComp = lMaxLenComp:   End Property

Public Property Get MaxLenType() As Long:           MaxLenType = lMaxLenType:   End Property

Public Property Get Service(Optional ByVal wb_full_name) As String
    If sService = vbNullString Then
        sService = Split(Me.Wrkbk.Path, "\")(UBound(Split(Me.Wrkbk.Path, "\")))
    End If
    Service = sService
End Property

Friend Property Get CompName() As String:           CompName = sCompName:       End Property

Friend Property Let CompName(ByVal s As String)
' ----------------------------------------------------------------
' Provides the properties: - CompName (sCompName)
'                          - ExpFileFullName (sExpFileFullName)
' ----------------------------------------------------------------
    sCompName = s
    If Not wb Is Nothing Then
        sExpFileFullName = Me.ExpFilePath & "\" & sCompName & Me.ExpFileExtension
    End If
End Property

Friend Property Get ExpFileFullName() As String
    If sExpFileFullName = vbNullString Then
        sExpFileFullName = Me.ExpFilePath & "\" & Me.CompName & Me.ExpFileExtension
    End If
    ExpFileFullName = sExpFileFullName
End Property

Friend Property Get ExpFile() As File
    
    Dim fso As New FileSystemObject
    
    With Me
        If flExpFile Is Nothing Then
            If fso.FileExists(.ExpFileFullName) Then
                Set flExpFile = GetFile(.ExpFileFullName)
                '~~ The component apparently has yet no never been exported
                If Not Me.VBComp Is Nothing Then
                    Me.VBComp.Export .ExpFileFullName
                    Set flExpFile = GetFile(.ExpFileFullName)
                End If
            End If
        End If
    End With
    
    Set ExpFile = flExpFile
    Set fso = Nothing
    
End Property

Friend Property Set ExpFile(ByVal fl As File):      Set flExpFile = fl:     End Property

Friend Property Get ExpFilePath() As String
    Dim fso As New FileSystemObject
    If sExpFilePath = vbNullString Then
        If Not Me.Wrkbk Is Nothing Then
            sExpFilePath = fso.GetParentFolderName(Me.Wrkbk.FullName)
        ElseIf Me.WrkbkFullName <> vbNullString Then
            sExpFilePath = fso.GetParentFolderName(Me.WrkbkFullName)
        End If
    End If
    ExpFilePath = sExpFilePath
End Property

Friend Property Get ExpFolder() As Folder:      Set ExpFolder = foExport:   End Property

Friend Property Get ExpFileExtension() As String
    
    If sExpFileExtension = vbNullString Then
        If vbc Is Nothing Then Set vbc = wb.VbProject.VBComponents(sCompName)
        With Me
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExtension = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExtension = ".cls"
                Case vbext_ct_MSForm:       sExpFileExtension = ".frm"
                Case vbext_ct_Document:     sExpFileExtension = ".cls"
            End Select
        End With
    End If
    ExpFileExtension = sExpFileExtension

End Property

Friend Property Get DsplyAllChanges() As Long
' -------------------------------------------
' Display of all differences between the
' clone's and the raw's Export-File.
' -------------------------------------------
    Const PROC = "DisplayAllChanges"
    
    On Error GoTo eh
    Dim sMsg As tMsg
    
    If Me.Changed Then
        mFile.Compare fc_file_left:=sTmpExpFileFullName _
                    , fc_file_right:=Me.ExpFileFullName _
                    , fc_left_title:="'Clone' component's export file '" & sTmpExpFileFullName & "'" _
                    , fc_right_title:="'Raw' component's export file '" & Me.ExpFileFullName & "'"
    Else
        sMsg.Section(1).sText = "The Export Files of the component '" & Me.CompName & "'"
        sMsg.Section(2).sText = "Temp Export:    " & sTmpExpFileFullName & vbLf & _
                                "Current Exüort: " & Me.ExpFileFullName
        sMsg.Section(2).bMonspaced = True
        sMsg.Section(3).sText = "are identical. The clone component has thus not changed."
        sMsg.Section(4).sLabel = "Please note!"
        sMsg.Section(4).sText = "Differences in empty code lines or upper/lower case differences do not constitute a relevant code change."
        
        mMsg.Dsply msg_title:="Display of code changes failed!" _
                 , msg:=sMsg
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get KindOfComp() As enKindOfComp
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    If lKindOfComp = enKindOfComp.enUnknown Then
        If mHostedRaws.Exists(sCompName) Then
            If mHostedRaws.HostFullName(comp_name:=sCompName) = Me.Wrkbk.FullName _
            Then lKindOfComp = enHostedRaw _
            Else lKindOfComp = enRawClone
        ElseIf lKindOfComp = enKindOfComp.enUnknown Then
            lKindOfComp = enInternal
        End If
        cLog.Entry = Me.TypeString & " (" & KoCStrng(lKindOfComp) & ")"
    End If
    
xt: KindOfComp = lKindOfComp
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property


Private Function KoCStrng(ByVal koc As enKindOfComp) As String
    Select Case koc
        Case enKindOfComp.enUnknown:        KoCStrng = "Unknown"
        Case enKindOfComp.enInternal:       KoCStrng = "Internal, neither Raw nor Clone"
        Case enKindOfComp.enRawClone:       KoCStrng = "Clone"
        Case enKindOfComp.enHostedRaw:      KoCStrng = "Raw (Hosted)"
    End Select
End Function

Private Property Get ReplaceRawWithCloneConfirmed() As Boolean
' -------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' --------------------------------------------------------------
    Const PROC = "ReplaceRawWithCloneConfirmed"
    
    On Error GoTo eh
    Dim sReplyUpdateOrigin      As String
    Dim sReplyDsplyDiff         As String
    Dim sReplyIgnore            As String
    Dim sMsg                    As tMsg
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
    Set cllButtons = mMsg.Buttons(sReplyUpdateOrigin, sReplyDsplyDiff, sReplyIgnore)
    
    '~~ Offer options
    sTitle = "A ""Remotely hosted raw's clone has been modified in this VB-Project!"
    With sMsg
        .Section(1).sLabel = "About this incident:"
        .Section(1).sText = "The module '" & Me.CompName & "' is a clone of the raw component hosted in Workbook '" & cRaw.WrkbkBaseName & "' " & _
                            "This clone component's code has usually has been synched with the remote raw componet's code. " & _
                            "The remote raw hasn't been changed but the cloned component's code has."
        .Section(2).sLabel = "About the choices:"
        .Section(2).sText = "When the code modification is ignored it will be regarded temporarily only and overwrtitten with the next open of this Workbook. " & _
                            "When the remote The module, updated when this Workbook has opened, is no longer identical with its origin. " & _
                            "I.e. the component's code has been changed in this Workbook using it instead of in the " & _
                            "Workbook the origin is ""hosted""."
        .Section(3).sText = "Take your decission or postpone it after having displayed the difference."
    End With
        
    sReply = sReplyDsplyDiff
    Do
        sReply = mMsg.Dsply(msg_title:=sTitle, _
                            msg:=sMsg, _
                            msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sReplyUpdateOrigin
                ReplaceRawWithCloneConfirmed = True
                Exit Do
            Case sReplyDsplyDiff
                mFile.Compare fc_file_left:=Me.ExpFileFullName _
                            , fc_file_right:=cRaw.ExpFileFullName _
                            , fc_left_title:="The cloned raw's current code in Workbook/VBProject " & Me.WrkbkBaseName & " (" & Me.ExpFileFullName & ")" _
                            , fc_right_title:="The remote raw's current code in Workbook/VBProject " & cRaw.WrkbkBaseName & " (" & cRaw.ExpFileFullName & ")"
                ' The question is re-displayed for a decison
            Case sReplyIgnore
                ReplaceRawWithCloneConfirmed = False
                Exit Do
        End Select
    Loop
                                            
xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get VBComp() As VBComponent
    Const PROC = "VBComp_Get"
    
    On Error GoTo eh
    
    If vbc Is Nothing Then
        If Not mCompMan.CompExists(ce_wb:=wb, ce_comp_name:=sCompName) _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "A Component named '" & sCompName & "' does not exist in Workbook '" & wb.Name & "'!"
        Set vbc = wb.VbProject.VBComponents(sCompName)
    End If
    Set VBComp = vbc

xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Set VBComp(ByRef vb_comp As VBComponent):       Set vbc = vb_comp:                      End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If wb Is Nothing Then
        If Not fso.FileExists(sWrkbkFullName) _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
        Set wb = mCompMan.WbkGetOpen(sWrkbkFullName)
    Else
        Set Wrkbk = wb
    End If
    
xt: Set fso = Nothing
    Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
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
    
    Dim fso As New FileSystemObject
    Dim vbc As VBComponent
    
    Set wb = wbk
    sWrkbkFullName = wb.FullName
    sWrkbkBaseName = fso.GetBaseName(sWrkbkFullName)
    sExpFilePath = fso.GetParentFolderName(sWrkbkFullName)
    Set foExport = fso.GetFolder(sExpFilePath)
    sService = Split(wb.Path, "\")(UBound(Split(wb.Path, "\")))
    For Each vbc In wb.VbProject.VBComponents
        lMaxLenComp = mBasic.Max(lMaxLenComp, Len(vbc.Name))
    Next vbc
    Set fso = Nothing
    
End Property

Friend Property Get WrkbkBaseName() As String:                 WrkbkBaseName = BaseName(sWrkbkFullName):        End Property

Friend Property Get WrkbkFullName() As String:                 WrkbkFullName = sWrkbkFullName:                  End Property

Friend Property Let WrkbkFullName(ByVal s As String):          sWrkbkFullName = s:                              End Property

Private Function BaseName(ByVal s As String) As String
    With New FileSystemObject:  BaseName = .GetBaseName(s): End With
End Function

Friend Property Get Changed() As Boolean
' --------------------------------------
' returns TRUE when the temporary Export
' File differs from the regular Export
' File which indicates that the code had
' changed.
' --------------------------------------
    Const PROC = "Changed"
    
    On Error GoTo eh
    
    If dctChanges Is Nothing Then
        With New FileSystemObject
            If Not .FileExists(Me.ExpFileFullName) Then
                Me.VBComp.Export Me.ExpFileFullName
            End If
            sTmpFolder = Me.ExpFilePath & "\Temp"
            If Not .FolderExists(sTmpFolder) Then
                DoEvents
                .CreateFolder sTmpFolder
                DoEvents
            End If
            sTmpExpFileFullName = sTmpFolder & "\" & vbc.Name & sExpFileExtension
            vbc.Export sTmpExpFileFullName
            DoEvents
    
            Set dctChanges = _
            mFile.Differs(fd_file1:=.GetFile(sTmpExpFileFullName) _
                        , fd_file2:=.GetFile(Me.ExpFileFullName) _
                        , fd_ignore_empty_records:=True _
                        , fd_compare:=vbTextCompare _
                        , fd_stop_after:=1 _
                         )
        End With

    End If
    Changed = dctChanges.Count <> 0
    
xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Private Function DebugComp() As String
    If lMaxCompLength > 0 _
    Then DebugComp = sCompName & " " & String(lMaxCompLength - Len(sCompName), ".") & ": " _
    Else DebugComp = sCompName & " : "
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
End Function

Public Function Exists(ByVal sComp As String) As Boolean
    Dim vbc As VBComponent

    On Error Resume Next
    Set vbc = Me.Wrkbk.VbProject.VBComponents(sComp)
    Exists = Err.Number = 0
    Set vbc = Nothing

End Function


Public Function ExpFileToDict(ByVal fl As File) As Dictionary
' --------------------------------------------------------------
' Returns the content of the export file (fl) as Dictionary.
' All lines preceeding "Attribute " lines and "Attribute " lines
' are excempted.
' --------------------------------------------------------------
    Const PROC = "ExpFileToArray"
    Const ATTRIB_LINE = "Attribute "

    On Error GoTo eh
    Dim i               As Long
    Dim dct             As Dictionary
    Dim v               As Variant
    Dim cllRemove       As New Collection
    Dim sLine           As String
    Dim cll             As Collection
    
    Set cll = New Collection
    Set dct = mFile.Dict(fl)
    
    '~~ Remove the extra module information lines
    For Each v In dct
        If Left(sLine, Len(ATTRIB_LINE)) = ATTRIB_LINE _
        And Left(dct(v), Len(ATTRIB_LINE)) <> ATTRIB_LINE _
        Then Exit For
        cllRemove.Add v
        sLine = dct(v)
    Next v
    For Each v In cllRemove
        dct.Remove v
    Next v
    Set cllRemove = Nothing
    
    '~~ Remove any leading empty items/lines
    Set cllRemove = New Collection
    For Each v In dct
        If dct(v) <> vbNullString Then
            Exit For
        End If
        Debug.Print "Collect fo removeal " & i & " item with content >>" & dct.Item(i) & "<<"
        cllRemove.Add i:    i = i + 1
        sLine = dct(v)
    Next v
    While Len(dct.Items(dct.Count - 1)) < 2
        dct.Remove dct.Keys((dct.Count - 1))
        Debug.Print "Leading empty item/line removed"
    Wend
    
    '~~ Remove any trailing empty items/lines
    For i = 1 To cllRemove.Count
        dct.Remove cllRemove.Item(i)
        Debug.Print "Trailing empty item/line removed"
    Next i
    
    Set ExpFileToDict = dct

xt: Set cll = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function GetFile(ByVal sFile As String) As File
    Dim fso As New FileSystemObject
    With fso:  Set GetFile = .GetFile(sFile):  End With
    Set fso = Nothing
End Function

Public Sub ReplaceRawWithClone()
' --------------------------------------------------------------
' Replace the raw component's code with the Export-File of the
' clone component.
' --------------------------------------------------------------
    Const PROC = "ReplaceRawWithClone"

    On Error GoTo eh
    Dim wbRawHost   As Workbook
        
    Application.EnableEvents = False ' prevent the UpdateRawClones service
    Set wbRawHost = mCompMan.WbkGetOpen(cRaw.HostFullName) ' Wrkbk is the Workbook hosting the raw Component
    Application.EnableEvents = True
    
    mRenew.ByImport rn_wb:=wbRawHost _
                  , rn_comp_name:=sCompName _
                  , rn_exp_file_full_name:=Me.ExpFileFullName
    cLog.Entry = "Remote raw (hosted in '" & wbRawHost.Name & "' replaced by (re-)import of the clone's export file '" & Me.ExpFileFullName & "'"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub ReplaceRawWithCloneWhenConfirmed( _
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
    Dim fso As New FileSystemObject
    
    If ReplaceRawWithCloneConfirmed Then
            '~~ Get the origin host Workbook open
        With cRaw
            .HostFullName = mHostedRaws.HostFullName(Me.CompName)
            .CompName = Me.CompName
            .ExpFileExtension = Me.ExpFileExtension ' required to build the export file's full name
            Set .ExpFile = fso.GetFile(.ExpFileFullName)
            Set .RawHost = mCompMan.WbkGetOpen(.HostFullName)
            Set .RawVbc = .RawHost.VbProject.VBComponents(.CompName)
            .TypeString = Me.TypeString
            Select Case .RawType
                Case vbext_ct_Document
                    '~~ A data module cannot be replaced like the other ones. Thus only its code
                    '~~ is synchronized - disregarding any consequences this may have when the
                    '~~ structure and design of the underlying Workbook or Worksheet has changed.
                    mSync.ByCodeLines bcl_clone_wb:=cComp.Wrkbk _
                                    , bcl_comp_name:=.CompName _
                                    , bcl_raw_host_full_name:=.HostFullName _
                                    , bcl_raw_codelines:=.CodeLines
                Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm
                    Me.ReplaceRawWithClone
            End Select
            rwu_updated = True
        End With ' cRaw
    End If

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function CompType() As Dictionary
    Set CompType = dctType
End Function

Public Function TypeString( _
            Optional ByVal vbc As VBComponent = Nothing) As String
' ----------------------------------------------------------------
' Returns the component type derived from sComp.
' ----------------------------------------------------------------
    Const PROC = "TypeString"
    
    On Error GoTo eh
    
    If vbc Is Nothing Then Set vbc = Me.VBComp
    If dctType.Exists(vbc.Type) _
    Then TypeString = dctType(vbc.Type) _
    Else TypeString = "Type '" & CStr(vbc.Type) & "' unknown"
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

'Private Function UserFormSizeIsDifferent( _
'                                   ByVal sid_file_1 As File, _
'                                   ByVal sid_file_2 As File) As Boolean
'    Dim sFrxFile1Path  As String
'    Dim sFrxFile2Path  As String
'    Dim fso             As New FileSystemObject
'
'    sFrxFile1Path = Replace$(sid_file_1.Path, ".frm", ".frx")
'    sFrxFile2Path = Replace$(sid_file_2.Path, ".frm", ".frx")
'
'    With fso
'        If sid_file_1.Size <> sid_file_2.Size Then
'            UserFormSizeIsDifferent = True
'        ElseIf .GetFile(sFrxFile1Path).Size <> .GetFile(sFrxFile2Path).Size Then
'            UserFormSizeIsDifferent = True
'        End If
'    End With
'    Set fso = Nothing
'
'End Function

Private Sub Class_Initialize()
    If dctType Is Nothing Then Set dctType = New Dictionary
    dctType.Add vbext_ct_ActiveXDesigner, "ActiveX Designer":   lMaxLenType = Max(Len(lMaxLenType), Len("ActiveX Designer"))
    dctType.Add vbext_ct_ClassModule, "Class Module":           lMaxLenType = Max(Len(lMaxLenType), Len("Class Module"))
    dctType.Add vbext_ct_Document, "Document Module":           lMaxLenType = Max(Len(lMaxLenType), Len("Document Module"))
    dctType.Add vbext_ct_MSForm, "UserForm":                    lMaxLenType = Max(Len(lMaxLenType), Len("UserForm"))
    dctType.Add vbext_ct_StdModule, "Standard Module":          lMaxLenType = Max(Len(lMaxLenType), Len("Standard Module"))
End Sub

Private Sub Class_Terminate()
    Set dctChanges = Nothing
    Set wb = Nothing
    Set vbc = Nothing
    With New FileSystemObject
        If .FolderExists(sTmpFolder) Then .DeleteFolder sTmpFolder
    End With
    Set flExpFile = Nothing
End Sub
